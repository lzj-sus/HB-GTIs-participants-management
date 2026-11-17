library(shiny)
library(bslib)
library(dplyr)
library(DBI)
library(RSQLite)
library(DT)
library(lubridate)
library(tidyr)
library(stringr)
library(openxlsx)
library(tibble)
library(purrr)
library(htmltools)
library(shinyWidgets)
library(readr)
library(ggplot2)

# =========================================================
# 0) Utility functions
# =========================================================
sanitize_filename <- function(x) {
  x <- as.character(x %||% "")
  x <- gsub("[\\\\/:*?\"<>|]", "_", x)
  x <- gsub("\\s+", "_", x)
  x
}

`%||%` <- function(a, b) {
  if (!is.null(a) && length(a) > 0 && !all(is.na(a))) a else b
}

safe_date <- function(x) suppressWarnings(as.Date(x))
now_chr <- function() format(Sys.time(), "%Y-%m-%d %H:%M:%S")
na2empty <- function(x) ifelse(is.na(x), "", x)
scalar_chr <- function(x, default = "") {
  if (is.null(x) || length(x) == 0 || all(is.na(x))) return(default)
  as.character(x[[1]])
}
numeric_or_na <- function(x) {
  out <- suppressWarnings(as.numeric(x))
  if (length(out) == 0 || is.na(out[[1]])) return(NA_real_)
  out[[1]]
}

read_records_csv <- function(path) {
  if (!file.exists(path)) return(NULL)
  tryCatch(
    read_csv(path, show_col_types = FALSE, locale = locale(encoding = "UTF-8")),
    error = function(e) NULL
  )
}

write_master_csv <- function(df, path, excel_csv = TRUE) {
  dir.create(dirname(normalizePath(path, winslash = "/", mustWork = FALSE)), recursive = TRUE, showWarnings = FALSE)
  tmp <- tempfile(pattern = paste0(basename(path), "_"), tmpdir = dirname(path), fileext = ".tmp")
  on.exit(unlink(tmp), add = TRUE)
  if (isTRUE(excel_csv)) {
    write_excel_csv(df, tmp)
  } else {
    write_csv(df, tmp)
  }
  ok <- file.copy(tmp, path, overwrite = TRUE)
  if (!isTRUE(ok)) {
    stop("Could not write the master file: ", path, ". Close it if it is open in Excel or WPS, then try again.")
  }
  invisible(TRUE)
}

append_record_csv <- function(path, df_row, excel_csv = TRUE) {
  stopifnot(is.data.frame(df_row), nrow(df_row) == 1)
  if (!file.exists(path)) {
    write_master_csv(df_row, path, excel_csv = excel_csv)
  } else {
    old <- read_csv(path, show_col_types = FALSE, locale = locale(encoding = "UTF-8"), col_types = cols(.default = col_character()))
    df_row <- df_row %>% mutate(across(everything(), as.character))
    all_names <- union(names(old), names(df_row))
    for (nm in setdiff(all_names, names(old))) old[[nm]] <- NA
    for (nm in setdiff(all_names, names(df_row))) df_row[[nm]] <- NA
    old <- old[, all_names, drop = FALSE]
    df_row <- df_row[, all_names, drop = FALSE]
    write_master_csv(bind_rows(old, df_row), path, excel_csv = excel_csv)
  }
}

save_individual_record_csv <- function(df_row, pid, visit_date, out_dir, create_pid_dir = TRUE) {
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  target_dir <- out_dir
  if (isTRUE(create_pid_dir)) {
    target_dir <- file.path(out_dir, sanitize_filename(pid))
    if (!dir.exists(target_dir)) dir.create(target_dir, recursive = TRUE, showWarnings = FALSE)
  }
  stamp <- format(Sys.time(), "%Y-%m-%d_%H-%M-%S")
  file_name <- paste0(sanitize_filename(pid), "_", sanitize_filename(as.character(visit_date)), "_", stamp, ".csv")
  out_path <- file.path(target_dir, file_name)
  write_excel_csv(df_row, out_path)
  out_path
}


id_aliases_for_module <- function(pid, module = c("default", "u3")) {
  module <- match.arg(module)
  pid <- scalar_chr(pid, "")
  if (pid == "") return(character(0))
  num <- stringr::str_extract(pid, "\\d+")
  unique(stats::na.omit(c(
    pid,
    if (!is.na(num) && nzchar(num)) paste0("PD", num) else NA_character_,
    if (!is.na(num) && nzchar(num)) paste0("P", num) else NA_character_,
    if (!is.na(num) && nzchar(num)) paste0("S", num) else NA_character_
  )))
}

list_individual_record_files <- function(individual_dir, pid = NULL, aliases = NULL) {
  ids <- unique(c(aliases %||% character(0), pid %||% character(0)))
  ids <- ids[!is.na(ids) & nzchar(ids)]
  if (length(ids) == 0) return(character(0))

  files <- character(0)
  for (one_id in ids) {
    pid_clean <- sanitize_filename(one_id)

    pid_dir <- file.path(individual_dir, pid_clean)
    if (dir.exists(pid_dir)) {
      files <- c(files, list.files(pid_dir, pattern = "\\.csv$", full.names = TRUE, recursive = TRUE))
    }

    if (dir.exists(individual_dir)) {
      root_files <- list.files(individual_dir, pattern = "\\.csv$", full.names = TRUE, recursive = FALSE)
      root_files <- root_files[startsWith(basename(root_files), paste0(pid_clean, "_"))]
      files <- c(files, root_files)
    }
  }

  unique(files[file.exists(files)])
}

read_individual_history_csv <- function(individual_dir, pid, aliases = NULL) {
  files <- list_individual_record_files(individual_dir, pid = pid, aliases = aliases)
  if (length(files) == 0) return(NULL)

  out_list <- lapply(files, function(f) {
    df <- tryCatch(
      read_csv(
        f,
        show_col_types = FALSE,
        locale = locale(encoding = "UTF-8"),
        col_types = cols(.default = col_character())
      ),
      error = function(e) NULL
    )
    if (is.null(df) || nrow(df) == 0) return(NULL)
    df <- df %>% mutate(across(everything(), as.character))
    df$source_file <- basename(f)
    df$file_mtime <- as.character(file.info(f)$mtime)
    df
  })

  out_list <- Filter(Negate(is.null), out_list)
  if (length(out_list) == 0) return(NULL)

  bind_rows(out_list)
}

order_history_df <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(df)

  parse_any_time <- function(x) {
    if (is.null(x)) return(as.POSIXct(rep(NA, nrow(df))))
    suppressWarnings(parse_date_time(
      as.character(x),
      orders = c("Y-m-d H:M:S", "Y/m/d H:M:S", "Y-m-d H:M", "Y/m/d H:M", "Y-m-d", "Y/m/d")
    ))
  }

  saved_time <- if ("saved_at_local" %in% names(df)) parse_any_time(df$saved_at_local) else as.POSIXct(rep(NA, nrow(df)))
  file_time  <- if ("file_mtime" %in% names(df)) parse_any_time(df$file_mtime) else as.POSIXct(rep(NA, nrow(df)))
  diary_time <- if ("diary_date" %in% names(df)) suppressWarnings(as.POSIXct(as.Date(df$diary_date))) else as.POSIXct(rep(NA, nrow(df)))
  date_time  <- if ("date" %in% names(df)) suppressWarnings(as.POSIXct(as.Date(df$date))) else as.POSIXct(rep(NA, nrow(df)))
  visit_time <- if ("visit_date" %in% names(df)) suppressWarnings(as.POSIXct(as.Date(df$visit_date))) else as.POSIXct(rep(NA, nrow(df)))

  df$.sort_time <- saved_time
  miss <- is.na(df$.sort_time); df$.sort_time[miss] <- file_time[miss]
  miss <- is.na(df$.sort_time); df$.sort_time[miss] <- diary_time[miss]
  miss <- is.na(df$.sort_time); df$.sort_time[miss] <- date_time[miss]
  miss <- is.na(df$.sort_time); df$.sort_time[miss] <- visit_time[miss]

  if (!"source_file" %in% names(df)) df$source_file <- ""
  df %>% arrange(desc(.sort_time), desc(source_file))
}

history_line_plot <- function(df, date_col, value_col, title_text, y_label = NULL) {
  if (is.null(df) || nrow(df) == 0 || !(date_col %in% names(df)) || !(value_col %in% names(df))) {
    return(blank_plot("No history yet"))
  }
  plot_df <- df %>%
    mutate(
      .x_date = safe_date(.data[[date_col]]),
      .y_value = suppressWarnings(as.numeric(.data[[value_col]]))
    ) %>%
    filter(!is.na(.x_date), !is.na(.y_value)) %>%
    arrange(.x_date)

  if (nrow(plot_df) == 0) return(blank_plot("No data to plot"))

  ggplot(plot_df, aes(x = .x_date, y = .y_value, group = 1)) +
    geom_line(linewidth = 0.9) +
    geom_point(size = 2.2) +
    labs(title = title_text, x = NULL, y = y_label %||% value_col) +
    theme_minimal(base_size = 14)
}


# =========================================================
# 1) Participant information system (core logic)
# =========================================================
db_file <- Sys.getenv("SQLITE_DB_FILE", "tis_pd_visit_manager.sqlite")

db_driver <- function() {
  tolower(Sys.getenv("DB_DRIVER", "sqlite"))
}

get_con <- function() {
  driver <- db_driver()
  if (identical(driver, "postgres")) {
    if (!requireNamespace("RPostgres", quietly = TRUE)) {
      stop("RPostgres package is required when DB_DRIVER=postgres.")
    }
    return(DBI::dbConnect(
      RPostgres::Postgres(),
      host = Sys.getenv("PGHOST", "127.0.0.1"),
      port = as.integer(Sys.getenv("PGPORT", "5432")),
      dbname = Sys.getenv("PGDATABASE", "tis_pd"),
      user = Sys.getenv("PGUSER", "postgres"),
      password = Sys.getenv("PGPASSWORD", ""),
      sslmode = Sys.getenv("PGSSLMODE", "prefer")
    ))
  }

  DBI::dbConnect(RSQLite::SQLite(), db_file)
}

status_levels <- c("planned", "done", "missed", "rescheduled", "paused", "dropout")
status_label_map <- c(
  planned = "Pending",
  done = "Completed",
  missed = "Not completed",
  rescheduled = "Rescheduled",
  paused = "Rest/paused",
  dropout = "Withdrawn"
)

event_type_label_map <- c(
  onsite = "On-site visit",
  home = "Home session",
  online = "Online assessment",
  followup = "Follow-up",
  pause = "Rest"
)

get_event_order <- function(event_name, event_type, session_no = NA) {
  case_when(
    grepl("OFF-BASELINE", event_name %||% "") ~ 1,
    grepl("ON-1 PRE/POST", event_name %||% "") ~ 2,
    identical(event_type, "home") ~ 3,
    identical(event_type, "pause") ~ 4,
    grepl("4-week post-intervention assessment", event_name %||% "") ~ 6,
    grepl("4-week post-intervention follow-up", event_name %||% "") ~ 8,
    grepl("8-week post-intervention follow-up", event_name %||% "") ~ 9,
    TRUE ~ 99
  )
}

get_display_date_vec <- function(status, planned_date, actual_date) {
  out <- ifelse(status == "done" & !is.na(actual_date), actual_date, planned_date)
  as.Date(out, origin = "1970-01-01")
}

init_db <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)

  dbExecute(con, "
    CREATE TABLE IF NOT EXISTS subjects (
      subject_id TEXT PRIMARY KEY,
      initials TEXT,
      visit1_date TEXT,
      visit2_date TEXT,
      note TEXT,
      is_dropout INTEGER DEFAULT 0,
      dropout_date TEXT,
      dropout_reason TEXT,
      created_at TEXT
    )
  ")

  visit_id_sql <- if (identical(db_driver(), "postgres")) "id SERIAL PRIMARY KEY" else "id INTEGER PRIMARY KEY AUTOINCREMENT"
  dbExecute(con, paste0("
    CREATE TABLE IF NOT EXISTS visits (
      ", visit_id_sql, ",
      subject_id TEXT,
      event_name TEXT,
      event_type TEXT,
      session_no INTEGER,
      planned_date TEXT,
      status TEXT,
      actual_date TEXT,
      note TEXT,
      completed_by TEXT,
      created_at TEXT,
      updated_at TEXT
    )
  "))

  subject_cols <- dbListFields(con, "subjects")
  if (!"full_name" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN full_name TEXT")
  if (!"sex" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN sex TEXT")
  if (!"birth_month" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN birth_month TEXT")
  if (!"height_cm" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN height_cm REAL")
  if (!"weight_kg" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN weight_kg REAL")
  if (!"education" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN education TEXT")
  if (!"occupation" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN occupation TEXT")
  if (!"disease_years" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN disease_years REAL")
  if (!"is_active" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN is_active INTEGER DEFAULT 1")
  if (!"is_dropout" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN is_dropout INTEGER DEFAULT 0")
  if (!"dropout_date" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN dropout_date TEXT")
  if (!"dropout_reason" %in% subject_cols) dbExecute(con, "ALTER TABLE subjects ADD COLUMN dropout_reason TEXT")
  dbExecute(con, "UPDATE subjects SET is_active = '1' WHERE is_active IS NULL")
  dbExecute(con, "UPDATE subjects SET is_dropout = '0' WHERE is_dropout IS NULL")

  create_index_if_possible <- function(table_name, index_name, cols) {
    if (!dbExistsTable(con, table_name)) return(invisible(FALSE))
    existing_cols <- dbListFields(con, table_name)
    if (!all(cols %in% existing_cols)) return(invisible(FALSE))
    sql <- paste0(
      "CREATE INDEX IF NOT EXISTS ", index_name,
      " ON ", table_name, " (", paste(cols, collapse = ", "), ")"
    )
    tryCatch(dbExecute(con, sql), error = function(e) NULL)
    invisible(TRUE)
  }

  create_index_if_possible("subjects", "idx_subjects_dropout", c("is_active", "is_dropout", "subject_id"))
  create_index_if_possible("visits", "idx_visits_subject_status_date", c("subject_id", "status", "planned_date"))
  create_index_if_possible("family_diary_records", "idx_family_diary_pid_date", c("pid", "diary_date"))
  create_index_if_possible("scale_records", "idx_scale_records_type_pid_date", c("scale_type", "pid", "record_date"))
  create_index_if_possible("updrs3_records", "idx_updrs3_pid_visit_date", c("pid", "visit_date"))
}

read_subjects <- function() {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  if (!dbExistsTable(con, "subjects")) return(tibble())
  dbReadTable(con, "subjects") %>%
    as_tibble() %>%
    mutate(
      visit1_date = safe_date(visit1_date),
      visit2_date = safe_date(visit2_date),
      dropout_date = safe_date(dropout_date),
      full_name = as.character(full_name),
      sex = as.character(sex),
      birth_month = as.character(birth_month),
      height_cm = suppressWarnings(as.numeric(height_cm)),
      weight_kg = suppressWarnings(as.numeric(weight_kg)),
      education = as.character(education),
      occupation = as.character(occupation),
      disease_years = suppressWarnings(as.numeric(disease_years)),
      is_active = if ("is_active" %in% names(.)) ifelse(is.na(is_active), 1L, as.integer(is_active)) else 1L,
      is_dropout = ifelse(is.na(is_dropout), 0L, as.integer(is_dropout))
    ) %>%
    arrange(subject_id)
}

read_visits <- function() {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  if (!dbExistsTable(con, "visits")) return(tibble())
  dbReadTable(con, "visits") %>%
    as_tibble() %>%
    mutate(
      planned_date = safe_date(planned_date),
      actual_date = safe_date(actual_date),
      status = as.character(status),
      event_order = get_event_order(event_name, event_type, session_no),
      display_date = get_display_date_vec(status, planned_date, actual_date)
    ) %>%
    arrange(subject_id, display_date, event_order, session_no, event_name)
}

generate_schedule <- function(subject_id, visit1_date, visit2_date) {
  visit1_date <- as.Date(visit1_date)
  visit2_date <- as.Date(visit2_date)
  t0 <- now_chr()
  fixed_events <- tibble(
    subject_id = subject_id,
    event_name = c(
      "OFF-BASELINE (Visit 1)",
      "ON-1 PRE/POST (Visit 2) | First stimulation session",
      "4-week post-intervention assessment (Visit 3)",
      "4-week post-intervention follow-up",
      "8-week post-intervention follow-up"
    ),
    event_type = c("onsite", "onsite", "onsite", "followup", "followup"),
    session_no = c(NA, 1, NA, NA, NA),
    planned_date = c(visit1_date, visit2_date, visit2_date + 28, visit2_date + 56, visit2_date + 84),
    status = "planned",
    actual_date = as.Date(NA),
    note = NA_character_,
    completed_by = NA_character_,
    created_at = t0,
    updated_at = t0
  )
  home_events <- tibble(
    subject_id = subject_id,
    event_name = paste0("Home session ", 1:27),
    event_type = "home",
    session_no = 1:27,
    planned_date = visit2_date + 1:27,
    status = "planned",
    actual_date = as.Date(NA),
    note = NA_character_,
    completed_by = NA_character_,
    created_at = t0,
    updated_at = t0
  )
  bind_rows(fixed_events, home_events) %>%
    mutate(planned_date = as.character(planned_date), actual_date = as.character(actual_date))
}

insert_subject_with_schedule <- function(subject_id, initials, full_name, sex, birth_month, height_cm, weight_kg, education, occupation, disease_years, visit1_date, visit2_date, note) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  existing <- dbGetQuery(con, "SELECT COUNT(*) AS n FROM subjects WHERE subject_id = ?", params = list(subject_id))$n
  if (existing > 0) stop("This participant ID already exists.")

  dbWriteTable(con, "subjects", data.frame(
    subject_id = subject_id,
    initials = initials,
    full_name = full_name,
    sex = sex,
    birth_month = birth_month,
    height_cm = numeric_or_na(height_cm),
    weight_kg = numeric_or_na(weight_kg),
    education = education,
    occupation = occupation,
    disease_years = numeric_or_na(disease_years),
    visit1_date = as.character(as.Date(visit1_date)),
    visit2_date = as.character(as.Date(visit2_date)),
    note = note,
    created_at = now_chr(),
    is_active = 1,
    is_dropout = 0,
    dropout_date = NA,
    dropout_reason = NA,
    stringsAsFactors = FALSE
  ), append = TRUE)

  dbWriteTable(con, "visits", generate_schedule(subject_id, visit1_date, visit2_date), append = TRUE)
}

initialize_subject_progress <- function(subject_id, visit1_done = FALSE, visit1_actual_date = NA, visit2_done = FALSE, visit2_actual_date = NA,
                                        home_done_n = 0, visit3_done = FALSE, fu4_done = FALSE, fu8_done = FALSE) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  t0 <- now_chr()
  mark_done_by_name <- function(event_nm, actual_date = NA) {
    row <- dbGetQuery(con, "SELECT planned_date FROM visits WHERE subject_id = ? AND event_name = ? LIMIT 1", params = list(subject_id, event_nm))
    if (nrow(row) == 0) return(invisible(FALSE))
    final_actual <- safe_date(actual_date)
    if (is.na(final_actual)) final_actual <- safe_date(row$planned_date[1])
    dbExecute(con, "UPDATE visits SET status = 'done', actual_date = ?, updated_at = ? WHERE subject_id = ? AND event_name = ?",
              params = list(as.character(final_actual), t0, subject_id, event_nm))
  }
  if (visit1_done) mark_done_by_name("OFF-BASELINE (Visit 1)", visit1_actual_date)
  if (visit2_done) mark_done_by_name("ON-1 PRE/POST (Visit 2) | First stimulation session", visit2_actual_date)
  home_done_n <- min(max(as.integer(home_done_n), 0), 27)
  if (home_done_n > 0) for (k in seq_len(home_done_n)) mark_done_by_name(paste0("Home session ", k))
  if (visit3_done) mark_done_by_name("4-week post-intervention assessment (Visit 3)")
  if (fu4_done) mark_done_by_name("4-week post-intervention follow-up")
  if (fu8_done) mark_done_by_name("8-week post-intervention follow-up")
}

update_subject_info <- function(old_subject_id, new_subject_id, initials, full_name, sex, birth_month, height_cm, weight_kg, education, occupation, disease_years, visit1_date, visit2_date, note) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  if (old_subject_id != new_subject_id) {
    existing <- dbGetQuery(con, "SELECT COUNT(*) AS n FROM subjects WHERE subject_id = ?", params = list(new_subject_id))$n
    if (existing > 0) stop("The new participant ID already exists.")
  }
  dbExecute(con,
            "UPDATE subjects SET subject_id = ?, initials = ?, full_name = ?, sex = ?, birth_month = ?, height_cm = ?, weight_kg = ?, education = ?, occupation = ?, disease_years = ?, visit1_date = ?, visit2_date = ?, note = ? WHERE subject_id = ?",
            params = list(
              new_subject_id, initials, full_name, sex, birth_month,
              numeric_or_na(height_cm), numeric_or_na(weight_kg),
              education, occupation, numeric_or_na(disease_years),
              as.character(as.Date(visit1_date)), as.character(as.Date(visit2_date)),
              note, old_subject_id
            ))
  if (old_subject_id != new_subject_id) {
    dbExecute(con, "UPDATE visits SET subject_id = ? WHERE subject_id = ?", params = list(new_subject_id, old_subject_id))
  }
}

apply_subject_dropout <- function(subject_id, is_dropout = FALSE, dropout_date = NA, dropout_reason = NA) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  is_dropout <- isTRUE(is_dropout)
  dropout_date <- safe_date(dropout_date)
  t0 <- now_chr()
  dbExecute(con, "UPDATE subjects SET is_dropout = ?, dropout_date = ?, dropout_reason = ? WHERE subject_id = ?",
            params = list(ifelse(is_dropout, 1, 0), ifelse(is.na(dropout_date), NA, as.character(dropout_date)), dropout_reason, subject_id))
  if (!is_dropout || is.na(dropout_date)) {
    dbExecute(con,
              "UPDATE visits SET status = CASE WHEN status = 'dropout' THEN 'planned' ELSE status END, updated_at = ?
               WHERE subject_id = ? AND status = 'dropout' AND (note IS NULL OR note LIKE '[AUTO_DROPOUT]%')",
              params = list(t0, subject_id))
    return(invisible(TRUE))
  }
  dbExecute(con,
            "UPDATE visits SET status = 'dropout', actual_date = NULL,
             note = CASE WHEN note IS NULL OR note = '' THEN '[AUTO_DROPOUT] Participant withdrawn; remaining tasks cancelled'
                         ELSE '[AUTO_DROPOUT] Participant withdrawn; remaining tasks cancelled. ' || note END,
             updated_at = ? WHERE subject_id = ? AND planned_date >= ? AND status != 'done'",
            params = list(t0, subject_id, as.character(dropout_date)))
  invisible(TRUE)
}

delete_subject <- function(subject_id) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  dbExecute(con, "DELETE FROM visits WHERE subject_id = ?", params = list(subject_id))
  dbExecute(con, "DELETE FROM subjects WHERE subject_id = ?", params = list(subject_id))
}

shift_future_visit_dates <- function(con, old_row, new_anchor_date) {
  if (nrow(old_row) != 1) return(invisible(FALSE))

  sid <- old_row$subject_id[1]
  old_status <- scalar_chr(old_row$status[1], "planned")
  old_event_type <- scalar_chr(old_row$event_type[1], "")
  old_event_name <- scalar_chr(old_row$event_name[1], "")
  old_session <- suppressWarnings(as.integer(old_row$session_no[1]))
  old_actual <- safe_date(old_row$actual_date[1])
  old_planned <- safe_date(old_row$planned_date[1])
  old_anchor_date <- if (identical(old_status, "done") && !is.na(old_actual)) old_actual else old_planned
  new_anchor_date <- safe_date(new_anchor_date)

  if (is.na(old_anchor_date) || is.na(new_anchor_date)) return(invisible(FALSE))
  delta_days <- as.integer(new_anchor_date - old_anchor_date)
  if (is.na(delta_days) || delta_days == 0L) return(invisible(FALSE))

  visits <- dbReadTable(con, "visits") %>% as_tibble() %>%
    filter(subject_id == !!sid) %>%
    mutate(
      planned_date = safe_date(planned_date),
      session_no = suppressWarnings(as.integer(session_no)),
      status = as.character(status)
    )

  future_fixed_events <- c(
    "4-week post-intervention assessment (Visit 3)",
    "4-week post-intervention follow-up",
    "8-week post-intervention follow-up"
  )

  if (identical(old_event_type, "home") && !is.na(old_session)) {
    rows_to_shift <- visits %>%
      filter(event_type == "home", session_no > old_session, !(status %in% c("done", "dropout")))
  } else if (identical(old_event_name, "ON-1 PRE/POST (Visit 2) | First stimulation session")) {
    rows_to_shift <- visits %>%
      filter(
        (event_type == "home" | event_name %in% future_fixed_events),
        !(status %in% c("done", "dropout"))
      )
  } else {
    rows_to_shift <- tibble()
  }

  if (nrow(rows_to_shift) == 0) return(invisible(FALSE))

  t0 <- now_chr()
  for (i in seq_len(nrow(rows_to_shift))) {
    pd <- rows_to_shift$planned_date[i]
    if (is.na(pd)) next
    dbExecute(
      con,
      "UPDATE visits SET planned_date = ?, updated_at = ? WHERE id = ?",
      params = list(as.character(pd + delta_days), t0, rows_to_shift$id[i])
    )
  }
  invisible(TRUE)
}

recalc_subject_schedule <- function(subject_id) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  visits <- dbReadTable(con, "visits") %>% as_tibble() %>%
    filter(subject_id == !!subject_id) %>%
    mutate(planned_date = safe_date(planned_date), actual_date = safe_date(actual_date), status = as.character(status))
  if (nrow(visits) == 0) return(invisible(TRUE))

  dbExecute(con, "DELETE FROM visits WHERE subject_id = ? AND event_type = 'pause' AND note LIKE '[AUTO_PAUSE]%'", params = list(subject_id))
  visits <- visits %>% filter(!(event_type == "pause" & str_detect(na2empty(note), "^\\[AUTO_PAUSE\\]")))

  v1 <- visits %>% filter(event_name == "OFF-BASELINE (Visit 1)") %>% slice(1)
  v2 <- visits %>% filter(event_name == "ON-1 PRE/POST (Visit 2) | First stimulation session") %>% slice(1)
  visit3 <- visits %>% filter(event_name == "4-week post-intervention assessment (Visit 3)") %>% slice(1)
  fu4 <- visits %>% filter(event_name == "4-week post-intervention follow-up") %>% slice(1)
  fu8 <- visits %>% filter(event_name == "8-week post-intervention follow-up") %>% slice(1)
  homes <- visits %>% filter(event_type == "home") %>% arrange(session_no, id)
  if (nrow(v2) == 0) return(invisible(TRUE))

  is_done_scalar <- function(x) identical(scalar_chr(x, ""), "done")
  visit2_anchor <- if (!is.na(v2$actual_date[1]) && is_done_scalar(v2$status[1])) v2$actual_date[1] else v2$planned_date[1]
  if (is.na(visit2_anchor)) visit2_anchor <- Sys.Date()
  visit1_anchor <- if (nrow(v1) == 1) {
    if (!is.na(v1$actual_date[1]) && is_done_scalar(v1$status[1])) v1$actual_date[1] else v1$planned_date[1]
  } else {
    NA
  }

  t0 <- now_chr()
  prev_date <- visit2_anchor
  pause_rows <- list()

  if (nrow(homes) > 0) {
    for (i in seq_len(nrow(homes))) {
      r <- homes[i, ]
      r_status <- scalar_chr(r$status, "planned")
      new_plan <- if (identical(r_status, "done") && !is.na(r$actual_date[1])) {
        r$actual_date[1]
      } else {
        desired <- r$planned_date[1]
        if (is.na(desired)) desired <- prev_date + 1
        max(desired, prev_date + 1)
      }
      if (new_plan > prev_date + 1) {
        gap_days <- seq.Date(prev_date + 1, new_plan - 1, by = "day")
        for (gd in gap_days) {
          gd <- as.Date(gd, origin = "1970-01-01")
          pause_rows[[length(pause_rows) + 1]] <- data.frame(
            subject_id = subject_id,
            event_name = "Rest/paused",
            event_type = "pause",
            session_no = NA,
            planned_date = as.character(gd),
            status = "paused",
            actual_date = NA,
            note = "[AUTO_PAUSE] Created automatically when the intervention was postponed",
            completed_by = r$completed_by %||% NA_character_,
            created_at = t0,
            updated_at = t0,
            stringsAsFactors = FALSE
          )
        }
      }
      dbExecute(con, "UPDATE visits SET planned_date = ?, updated_at = ? WHERE id = ?", params = list(as.character(new_plan), t0, r$id[1]))
      prev_date <- new_plan
    }
  }

  visit3_date <- prev_date + 1
  fu4_date <- visit3_date + 28
  fu8_date <- visit3_date + 56

  update_fixed_if_exists <- function(df_row, new_date) {
    if (nrow(df_row) == 1) dbExecute(con, "UPDATE visits SET planned_date = ?, updated_at = ? WHERE id = ?", params = list(as.character(new_date), t0, df_row$id[1]))
  }
  update_fixed_if_exists(v1, visit1_anchor)
  update_fixed_if_exists(v2, visit2_anchor)
  update_fixed_if_exists(visit3, visit3_date)
  update_fixed_if_exists(fu4, fu4_date)
  update_fixed_if_exists(fu8, fu8_date)
  if (length(pause_rows) > 0) dbWriteTable(con, "visits", bind_rows(pause_rows), append = TRUE)
  dbExecute(con, "UPDATE subjects SET visit1_date = ?, visit2_date = ? WHERE subject_id = ?",
            params = list(as.character(visit1_anchor), as.character(visit2_anchor), subject_id))
  invisible(TRUE)
}

update_visit_full <- function(id, event_name, event_type, session_no, planned_date, status, actual_date, note, completed_by, recalc_after = TRUE) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  old_row <- dbGetQuery(con, "SELECT * FROM visits WHERE id = ?", params = list(id))
  if (nrow(old_row) != 1) stop("Task not found.")
  sid <- old_row$subject_id[1]

  planned_date <- safe_date(planned_date)
  actual_date <- safe_date(actual_date)
  status <- scalar_chr(status, "planned")
  event_type <- scalar_chr(event_type, scalar_chr(old_row$event_type, "onsite"))

  if (identical(status, "done") && is.na(actual_date)) actual_date <- planned_date
  if (!identical(status, "done")) actual_date <- NA
  if (!identical(event_type, "home")) session_no <- NA
  new_anchor_date <- if (identical(status, "done") && !is.na(actual_date)) actual_date else planned_date

  dbExecute(con,
            "UPDATE visits SET event_name = ?, event_type = ?, session_no = ?, planned_date = ?, status = ?, actual_date = ?, note = ?, completed_by = ?, updated_at = ? WHERE id = ?",
            params = list(event_name, event_type, ifelse(is.na(session_no), NA, as.integer(session_no)),
                          as.character(planned_date), status, ifelse(is.na(actual_date), NA, as.character(actual_date)),
                          note, completed_by, now_chr(), id))
  shift_future_visit_dates(con, old_row, new_anchor_date)
  if (isTRUE(recalc_after)) recalc_subject_schedule(sid)
}

complete_daily_home_tasks <- function(con, task_date, task_ids, operator) {
  task_date <- as.Date(task_date)
  if (length(task_date) != 1L || is.na(task_date) || task_date > Sys.Date()) {
    stop("Select today or a past date. Future tasks cannot be completed in a batch.")
  }
  task_ids <- unique(as.integer(task_ids))
  if (anyNA(task_ids)) stop("Invalid task ID. Refresh the page and try again.")
  operator <- if (length(operator) == 1L && !is.na(operator)) as.character(operator) else ""
  # Recheck eligibility inside the transaction; never recalculate the schedule here.
  DBI::dbWithTransaction(con, {
    changed <- 0L
    for (id in task_ids) {
      changed <- changed + DBI::dbExecute(con,
        paste0("UPDATE visits SET status = 'done', actual_date = ?, completed_by = ?, updated_at = ? ",
               "WHERE id = ? AND event_type = 'home' AND planned_date = ? ",
               "AND status IN ('planned', 'rescheduled') ",
               "AND subject_id IN (SELECT subject_id FROM subjects ",
               "WHERE COALESCE(is_active, 1) = 1 AND COALESCE(is_dropout, 0) = 0)"),
        params = list(as.character(task_date), operator, now_chr(), id, as.character(task_date)))
    }
    changed
  })
}

make_month_calendar <- function(month_date, visits_df) {
  month_date <- floor_date(as.Date(month_date), "month")
  start_day <- floor_date(month_date, "month")
  end_day <- ceiling_date(month_date, "month") - days(1)
  month_start_wday <- wday(start_day, week_start = 1)
  calendar_start <- start_day - days(month_start_wday - 1)
  calendar_end <- end_day + days(7 - wday(end_day, week_start = 1))
  all_days <- seq.Date(calendar_start, calendar_end, by = "day")

  daily_info <- tibble(date = all_days) %>% left_join(
    visits_df %>% filter(!is.na(display_date)) %>% mutate(date = as.Date(display_date)) %>%
      group_by(date) %>% summarise(
        total_n = n(),
        done_n = sum(status == "done", na.rm = TRUE),
        missed_n = sum(status == "missed", na.rm = TRUE),
        planned_n = sum(status == "planned", na.rm = TRUE),
        rescheduled_n = sum(status == "rescheduled", na.rm = TRUE),
        paused_n = sum(status == "paused", na.rm = TRUE),
        dropout_n = sum(status == "dropout", na.rm = TRUE),
        labels = paste0("<div style='font-size:11px; line-height:1.3;'>", paste0(event_name, " | ", status_label_map[status], collapse = "<br>"), "</div>"),
        .groups = "drop"
      ), by = "date"
  ) %>% mutate(across(c(total_n, done_n, missed_n, planned_n, rescheduled_n, paused_n, dropout_n), ~replace_na(., 0)), labels = replace_na(labels, ""))

  mat <- matrix(all_days, ncol = 7, byrow = TRUE)
  week_names <- c("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
  rows <- lapply(seq_len(nrow(mat)), function(i) {
    tds <- lapply(seq_len(ncol(mat)), function(j) {
      d <- as.Date(mat[i, j], origin = "1970-01-01")
      info <- daily_info %>% filter(date == d)
      in_month <- format(d, "%Y-%m") == format(month_date, "%Y-%m")
      bg <- if (in_month) "#FFFFFF" else "#F5F5F5"
      date_style <- if (in_month) "font-weight:700;" else "color:#999999;"
      summary_html <- if (nrow(info) > 0 && info$total_n[1] > 0) {
        HTML(paste0(
          "<div style='margin-top:4px; font-size:11px;'>",
          "<span style='margin-right:6px;'>All: ", info$total_n[1], "</span>",
          "<span style='margin-right:6px; color:#157347;'>Done: ", info$done_n[1], "</span>",
          "<span style='margin-right:6px; color:#B02A37;'>Missed: ", info$missed_n[1], "</span>",
          "<span style='margin-right:6px; color:#6C757D;'>Due: ", info$planned_n[1], "</span>",
          "<span style='margin-right:6px; color:#B58105;'>Moved: ", info$rescheduled_n[1], "</span>",
          "<span style='margin-right:6px; color:#7A4FA3;'>Rest: ", info$paused_n[1], "</span>",
          "<span style='color:#8B0000;'>W/drawn: ", info$dropout_n[1], "</span>",
          "</div><div style='margin-top:6px; max-height:90px; overflow:auto;'>", info$labels[1], "</div>"
        ))
      } else HTML("<div style='margin-top:8px; color:#BBBBBB; font-size:11px;'>No tasks</div>")
      tags$td(style = paste0("vertical-align:top; min-width:150px; height:150px; padding:8px; background:", bg, "; border:1px solid #E5E7EB;"),
              div(style = date_style, format(d, "%d")), summary_html)
    })
    tags$tr(tds)
  })
  tags$div(style = "margin-bottom:18px;", tags$h5(style = "margin-bottom:8px;", format(month_date, "%Y-%m")),
           tags$table(style = "border-collapse:collapse; width:100%; table-layout:fixed;",
                      tags$thead(tags$tr(lapply(week_names, function(x) tags$th(style = "border:1px solid #E5E7EB; background:#F8F9FA; padding:8px; text-align:center;", x)))),
                      tags$tbody(rows)))
}

make_full_subject_calendar <- function(visits_df) {
  visits_df <- visits_df %>% filter(!is.na(display_date))
  if (nrow(visits_df) == 0) return(div("This participant has no schedule yet."))
  min_month <- floor_date(min(visits_df$display_date), "month")
  max_month <- floor_date(max(visits_df$display_date), "month")
  months_seq <- seq.Date(min_month, max_month, by = "month")
  tagList(lapply(months_seq, function(m) make_month_calendar(m, visits_df)))
}

build_trajectory_table <- function(df) {
  if (nrow(df) == 0) return(tibble())
  df %>% mutate(type_label = event_type_label_map[event_type],
                status_label = status_label_map[status],
                display_date_chr = as.character(display_date)) %>%
    transmute(`Planned date` = as.character(planned_date),
              `Display date` = display_date_chr,
              Event = event_name,
              Type = type_label,
              Session = session_no,
              Status = status_label,
              Notes = note)
}

# =========================================================
# 2) Home diary
# =========================================================
diary_record_file <- "PD_Family_Diary_records.csv"
diary_individual_dir <- "records_diary_individual"

time_choices <- format(seq.POSIXt(as.POSIXct("2026-01-01 00:00:00"), as.POSIXct("2026-01-01 23:30:00"), by = "30 min"), "%H:%M")
time_to_minutes <- function(x) {
  if (is.null(x) || is.na(x) || x == "") return(NA_real_)
  sp <- strsplit(x, ":")[[1]]
  as.numeric(sp[1]) * 60 + as.numeric(sp[2])
}
calc_duration_hours <- function(start_time, end_time) {
  s <- time_to_minutes(start_time); e <- time_to_minutes(end_time)
  if (is.na(s) || is.na(e)) return(NA_real_)
  if (e < s) e <- e + 24 * 60
  round((e - s) / 60, 2)
}
format_slot_minute <- function(x) {
  x <- x %% (24 * 60)
  sprintf("%02d:%02d", x %/% 60, x %% 60)
}
generate_awake_slots <- function(start_time, end_time, nap_start = "", nap_end = "") {
  s <- time_to_minutes(start_time); e <- time_to_minutes(end_time)
  if (is.na(s) || is.na(e)) return(character(0))
  if (e <= s) e <- e + 24 * 60
  starts <- seq(s, e - 30, by = 30)
  ends <- starts + 30
  ns <- time_to_minutes(nap_start); ne <- time_to_minutes(nap_end)
  if (!is.na(ns) && !is.na(ne)) {
    if (ne <= ns) ne <- ne + 24 * 60
    if (ns < s) {
      ns <- ns + 24 * 60
      ne <- ne + 24 * 60
    }
    keep <- !(starts < ne & ends > ns)
    starts <- starts[keep]; ends <- ends[keep]
  }
  paste0(format_slot_minute(starts), "-", format_slot_minute(ends))
}

# =========================================================
# 3) Non-motor symptom scales (logic unchanged)
# =========================================================
# -------------------------
# Utility functions
# -------------------------
sanitize_filename <- function(x) {
  x <- gsub("[\\\\/:*?\"<>|]", "_", x)
  x <- gsub("\\s+", "_", x)
  x
}

read_records <- function(path) {
  if (!file.exists(path)) return(NULL)
  tryCatch(
    read_csv(path, show_col_types = FALSE, locale = locale(encoding = "UTF-8")),
    error = function(e) NULL
  )
}

append_record <- function(path, df_row) {
  append_record_csv(path, df_row, excel_csv = TRUE)
}

save_individual_record <- function(df_row, pid, visit_date, out_dir) {
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  
  # Create a separate subfolder for each participant inside the main folder
  pid_dir <- file.path(out_dir, sanitize_filename(pid))
  if (!dir.exists(pid_dir)) dir.create(pid_dir, recursive = TRUE, showWarnings = FALSE)
  
  stamp <- format(Sys.time(), "%Y-%m-%d_%H-%M-%S")
  file_name <- paste0(
    sanitize_filename(pid), "_",
    sanitize_filename(as.character(visit_date)), "_",
    stamp, ".csv"
  )
  
  out_path <- file.path(pid_dir, file_name)
  write_excel_csv(df_row, out_path)
  out_path
}

save_scale_record <- function(out, pid, visit_date, individual_dir, record_file,
                              records_reactive, session, success_prefix = "Saved",
                              scale_type = NULL) {
  db_ok <- TRUE
  if (!is.null(scale_type) && nzchar(scale_type)) {
    db_ok <- tryCatch({
      append_scale_db_record(scale_type, out)
      TRUE
    }, error = function(e) {
      showNotification(paste0("Database write failed: ", e$message), type = "error", duration = 8)
      FALSE
    })
  }
  if (!db_ok) return(FALSE)

  individual_ok <- FALSE
  individual_path <- NULL
  
  individual_ok <- tryCatch({
    individual_path <- save_individual_record(out, pid, visit_date, individual_dir)
    TRUE
  }, error = function(e) {
    showNotification(paste0("Could not save the individual record: ", e$message), type = "error", duration = 8)
    FALSE
  })
  
  if (!individual_ok) return(FALSE)
  
  master_ok <- tryCatch({
    append_record(record_file, out)
    TRUE
  }, error = function(e) {
    showNotification(
      paste0("Individual record saved, but the master file could not be updated: ", e$message,
             ". Close any open CSV, Excel or WPS file, then refresh the history."),
      type = "warning",
      duration = 10
    )
    FALSE
  })
  
  if (master_ok) {
    records_reactive(if (!is.null(scale_type) && nzchar(scale_type)) read_scale_records_db(scale_type, record_file, pid) else read_records(record_file))
    showNotification(
      paste0(
        success_prefix, ".\nDatabase updated.\nIndividual record: ",
        normalizePath(individual_path, winslash = "/", mustWork = FALSE),
        "\nMaster file: ",
        normalizePath(record_file, winslash = "/", mustWork = FALSE)
      ),
      type = "message",
      duration = 6
    )
  } else {
    showNotification(
      paste0("Individual record saved to: ",
             normalizePath(individual_path, winslash = "/", mustWork = FALSE)),
      type = "message",
      duration = 6
    )
  }
  TRUE
}

render_hist_table <- function(df, pid, keep_cols) {
  if (is.null(df) || nrow(df) == 0) return(NULL)
  aliases <- id_aliases_for_module(pid, module = "default")
  out <- df %>% filter(pid %in% aliases)
  if (nrow(out) == 0) return(NULL)
  keep_cols <- intersect(keep_cols, names(out))
  out %>% select(all_of(keep_cols))
}

filter_records_for_pid <- function(df, pid) {
  if (is.null(df) || nrow(df) == 0 || !"pid" %in% names(df)) return(df)
  aliases <- id_aliases_for_module(pid, module = "default")
  df %>% filter(pid %in% aliases)
}

read_db_wide_records <- function(table_name) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  if (!dbExistsTable(con, table_name)) return(NULL)
  tryCatch(dbReadTable(con, table_name) %>% as_tibble(), error = function(e) NULL)
}

append_db_wide_record <- function(table_name, df_row, source_label = "runtime_shiny") {
  stopifnot(is.data.frame(df_row), nrow(df_row) == 1)
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)

  df_out <- df_row %>%
    mutate(
      migration_source_file = source_label,
      migration_source_row = 0L,
      migration_key = paste0(source_label, "::", sanitize_filename(pid), "::", format(Sys.time(), "%Y%m%d%H%M%OS3"), "::", sample.int(1e6, 1)),
      migrated_at = now_chr()
    )

  dbWithTransaction(con, {
    db_ensure_columns(con, table_name, df_out)
    df_out <- db_align_to_table(con, table_name, df_out)
    dbAppendTable(con, table_name, df_out)
  })
  invisible(TRUE)
}

read_diary_records_db <- function(pid = NULL) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  db_df <- NULL
  if (dbExistsTable(con, "family_diary_records")) {
    aliases <- if (!is.null(pid)) id_aliases_for_module(pid, module = "default") else character(0)
    db_df <- tryCatch({
      if (length(aliases) > 0) {
        quoted_aliases <- paste(as.character(dbQuoteString(con, aliases)), collapse = ", ")
        dbGetQuery(
          con,
          paste0("SELECT * FROM family_diary_records WHERE pid IN (", quoted_aliases, ")")
        ) %>% as_tibble()
      } else {
        dbReadTable(con, "family_diary_records") %>% as_tibble()
      }
    }, error = function(e) NULL)
  }
  if (!is.null(db_df) && nrow(db_df) > 0) return(db_df)
  read_records_csv(diary_record_file)
}

payload_json_to_tibble <- function(payload_json) {
  rows <- lapply(payload_json, function(txt) {
    one <- tryCatch(jsonlite::fromJSON(txt, simplifyVector = FALSE), error = function(e) NULL)
    if (is.null(one)) return(NULL)
    one <- lapply(one, function(x) {
      if (is.null(x) || length(x) == 0) return(NA_character_)
      as.character(x[[1]])
    })
    as_tibble(one)
  })
  rows <- Filter(Negate(is.null), rows)
  if (length(rows) == 0) return(NULL)
  bind_rows(rows)
}

read_scale_records_db <- function(scale_type, fallback_file = NULL, pid = NULL) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  db_df <- NULL
  if (dbExistsTable(con, "scale_records")) {
    aliases <- if (!is.null(pid)) id_aliases_for_module(pid, module = "default") else character(0)
    db_df <- tryCatch(
      {
        if (length(aliases) > 0) {
          quoted_aliases <- paste(as.character(dbQuoteString(con, aliases)), collapse = ", ")
          dbGetQuery(
            con,
            paste0(
              "SELECT * FROM scale_records WHERE scale_type = ",
              as.character(dbQuoteString(con, scale_type)),
              " AND pid IN (", quoted_aliases, ")"
            )
          ) %>% as_tibble()
        } else {
          dbGetQuery(
            con,
            paste0("SELECT * FROM scale_records WHERE scale_type = ", as.character(dbQuoteString(con, scale_type)))
          ) %>% as_tibble()
        }
      },
      error = function(e) NULL
    )
  }
  if (!is.null(db_df) && nrow(db_df) > 0 && all(c("scale_type", "payload_json") %in% names(db_df))) {
    if (!is.null(pid) && "pid" %in% names(db_df)) db_df <- filter_records_for_pid(db_df, pid)
    if (nrow(db_df) > 0) {
      payload_df <- payload_json_to_tibble(db_df$payload_json)
      if (!is.null(payload_df) && nrow(payload_df) > 0) return(payload_df)
    }
  }
  if (!is.null(fallback_file)) read_records(fallback_file) else NULL
}

scale_total_column <- function(df_row) {
  candidates <- intersect(
    c("Total_GAI", "Total_GDS", "Total_ESS", "Total_PDSS2", "Total_SCOPA_AUT",
      "Total_UPDRS_I", "Total_UPDRS_II", "Total_NMSS", "Total_PDQ39", "Total_ISI"),
    names(df_row)
  )
  if (length(candidates) == 0) return(NA_real_)
  suppressWarnings(as.numeric(df_row[[candidates[1]]][1]))
}

append_scale_db_record <- function(scale_type, df_row) {
  level <- if ("ISI_Level" %in% names(df_row)) as.character(df_row$ISI_Level[1]) else NA_character_
  compact <- tibble(
    scale_type = scale_type,
    pid = as.character(df_row$pid[1]),
    record_date = as.character(df_row$date[1]),
    visit = as.character(df_row$visit[1]),
    total_score = scale_total_column(df_row),
    level = level,
    saved_at_local = as.character(df_row$saved_at_local[1]),
    payload_json = jsonlite::toJSON(as.list(df_row[1, , drop = TRUE]), auto_unbox = TRUE, null = "null", na = "null"),
    migration_source_file = "runtime_shiny",
    migration_source_row = 0L,
    migration_key = paste0("runtime_shiny::", scale_type, "::", sanitize_filename(df_row$pid[1]), "::", format(Sys.time(), "%Y%m%d%H%M%OS3"), "::", sample.int(1e6, 1)),
    migrated_at = now_chr()
  )

  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  dbWithTransaction(con, {
    db_ensure_columns(con, "scale_records", compact)
    compact <- db_align_to_table(con, "scale_records", compact)
    dbAppendTable(con, "scale_records", compact)
  })
  invisible(TRUE)
}

safe_input_num <- function(x) {
  if (is.null(x) || length(x) == 0 || identical(x, "") || is.na(x)) return(NA_real_)
  as.numeric(x)
}

get_scale_vals <- function(input, prefix, n) {
  vapply(seq_len(n), function(i) {
    safe_input_num(input[[paste0(prefix, i)]])
  }, numeric(1))
}

all_answered <- function(x) {
  !any(is.na(x))
}

score_sum_na <- function(x) sum(as.numeric(x), na.rm = TRUE)

# -------------------------
# Scale definitions
# -------------------------
GAI_ITEMS <- c(
  "I worry a lot of the time.","I find it difficult to make decisions.","I often feel nervous.","I find it hard to relax.","I often cannot enjoy life because of my worrying.",
  "Even small things upset me a great deal.","I often feel jittery.","I am a worrier.","Even small things make me worry in spite of myself.",
  "I often feel tense.","My thoughts often make me feel anxious.","My worrying makes my stomach upset.","I am a nervous person.","I always expect the worst to happen.",
  "I often feel panicky.","I feel that my worries interfere with my life.","I am often overwhelmed by my worries.","Sometimes I get stomach cramps from anxiety.","My excessive worrying has made me miss out on things.","I often feel irritable."
)

GDS_ITEMS <- c(
  "Are you basically satisfied with your life?","Have you dropped many of your activities and interests?","Do you feel that your life is empty?","Do you often get bored?","Are you hopeful about the future?",
  "Are you bothered by thoughts you cannot get out of your head?","Are you in good spirits most of the time?","Are you afraid that something bad is going to happen to you?","Do you feel happy most of the time?",
  "Do you feel helpless?","Do you often get restless and fidgety?","Do you prefer to stay at home rather than going out and doing new things?","Do you frequently worry about the future?","Do you feel you have more problems with memory than most?",
  "Do you think it is wonderful to be alive now?","Do you often feel downhearted and blue?","Do you feel worthless the way you are now?","Do you worry a lot about the past?","Do you find life very exciting?",
  "Is it hard for you to get started on new projects?","Do you feel full of energy?","Do you feel that your situation is hopeless?","Do you think that most people are better off than you are?","Do you frequently get upset over little things?",
  "Do you frequently feel like crying?","Do you have trouble concentrating?","Do you enjoy getting up in the morning?","Do you prefer to avoid social gatherings?","Is it easy for you to make decisions?","Is your mind as clear as it used to be?"
)
GDS_REVERSE_YES0 <- c(1,5,7,9,15,19,21,27,29,30)

ESS_ITEMS <- c(
  "Sitting and reading","Watching TV","Sitting inactive in a public place (e.g. a theatre or a meeting)","As a passenger in a car for an hour without a break",
  "Lying down to rest in the afternoon when circumstances permit","Sitting and talking to someone","Sitting quietly after a lunch without alcohol","In a car, while stopped for a few minutes in traffic"
)

PDSS2_ITEMS <- c(
  "Overall, did you sleep well over the past week?","Do you have difficulty falling asleep at night?","Do you have difficulty staying asleep?","Is your sleep interrupted at night by restless arms or legs?",
  "Is your sleep interrupted because you feel a strong urge to move your arms or legs?","Do you have nightmares at night?","Do you have hallucinations at night (seeing or hearing things that are not there)?","Do you get up at night to pass urine?",
  "Do you feel uncomfortable at night because difficulty moving stops you turning over or changing position?","Do you wake at night because of pain in your arms or legs?","Do you wake at night because of cramps in your arms or legs?",
  "Do you wake in the early morning because pain is caused by the position of your arms or legs?","Do you tremble when you wake up?","After waking in the morning, do you feel tired or sleepy?","Do you wake at night because of snoring or difficulty breathing?"
)

SCOPA_ITEMS <- c(
  "In the past month, have you had difficulty swallowing or choking?","In the past month, have you had drooling?","In the past month, has food stuck in your throat?","In the past month, have you felt bloated during meals?","In the past month, have you been constipated (fewer than 2 bowel movements a week)?",
  "In the past month, have you had difficulty passing stools (straining)?","In the past month, have you had faecal incontinence?","In the past month, have you had urinary urgency?","In the past month, have you had urinary incontinence?","In the past month, have you had a feeling of incomplete bladder emptying?",
  "In the past month, have you had a weak urinary stream?","In the past month, have you had to urinate again within 2 hours of urinating?","In the past month, have you got up at night to pass urine?","In the past month, have you felt dizzy or had blurred vision when changing position (from lying to sitting, or sitting to standing)?","Have you felt dizzy after standing for some time?",
  "In the past 6 months, have you fainted?","In the past month, have you had excessive sweating during the day?","In the past month, have you had excessive sweating at night?","In the past month, have you been unusually sensitive to cold?","In the past month, have you been unusually sensitive to heat?",
  "In the past month, have your eyes been oversensitive to bright light?","In the past month, have you had erectile dysfunction? (men)","In the past month, have you been unable to ejaculate? (men)","In the past month, have you had excessive vaginal dryness during intercourse? (women)","In the past month, have you had difficulty reaching orgasm? (women)"
)
SCOPA_GROUP <- c(rep("Gastrointestinal", 7), rep("Urinary", 6), rep("Cardiovascular", 3), rep("Thermoregulatory", 4), "Pupillomotor", rep("Sexual dysfunction", 4))

UPDRS1_ITEMS <- c(
  "Difficulties with memory, conversation, attention, clear thinking, or finding your way around the home or neighbourhood","Seeing, hearing, smelling or feeling things that do not really exist","Feeling depressed, sad, hopeless or unable to feel pleasure","Feeling nervous, worried or tense",
  "Showing little interest in activities or in interacting with other people","Unusually strong urges that are hard to control (e.g. gambling, cleaning, taking extra medication)","Difficulty falling asleep or staying asleep through the night","Not always staying awake during the day","Uncomfortable sensations such as pain, tingling or aching",
  "Difficulty controlling urine, such as urgency, frequency or incontinence","Constipation causing difficulty passing stools","Feeling light-headed, dizzy or faint on standing up from a chair or bed","Feeling fatigued, not the kind caused by drowsiness or sadness"
)

UPDRS2_ITEMS <- c(
  "Do you have problems with your speech?","Do you usually have too much saliva, either when awake or asleep?","Do you usually have trouble taking pills or eating? Do you need to cut up or crush your medication, or soften, chop or blend your food to avoid choking?","Do you usually have difficulty eating and using utensils, such as a knife and fork, a spoon or chopsticks?",
  "Do you usually have difficulty dressing, for example dressing slowly or needing help with buttons, zips, putting on clothes or jewellery?","Do you often find washing, bathing, shaving, brushing your teeth, combing your hair or other personal hygiene slow, or do you need help with them?","Do other people often find your handwriting hard to read?","Do you usually have difficulty doing your hobbies or the things you enjoy?",
  "Do you often have difficulty turning over in bed?","Do you often have shaking or tremor?","Do you often have difficulty getting out of bed, out of a car, or up from a low chair?","Do you often have difficulty walking and keeping your balance?","When walking, do you suddenly stop or freeze, as if your feet are stuck to the ground?"
)

NMSS_ITEMS <- c(
  "Light-headedness, dizziness or weakness on standing up from lying or sitting","Falls caused by dizziness or loss of consciousness","Falling asleep during the day in situations such as talking, eating, watching TV or reading","Fatigue or lack of energy affecting daytime activities","Difficulty falling asleep or waking frequently at night",
  "Discomfort in the legs when sitting or lying down, relieved by moving them","Loss of interest in what is happening around you","Reduced motivation to do things, reluctance to try new things","Looking or feeling sad or low in mood","Feeling anxious, tense or panicky",
  "Feeling emotionally flat, with a lack of normal emotional responses","Lack of pleasure in everyday life","Seeing or hearing things that are not there","Delusions, such as believing that someone is trying to harm you, that you have been robbed, or that someone is being unfaithful","Double vision, seeing one object as two",
  "Difficulty concentrating, for example when reading or in conversation","Difficulty remembering recent events","Forgetting to do things, such as taking medication","Drooling during the day","Difficulty swallowing or choking",
  "Constipation (fewer than 3 bowel movements a week)","Urinary urgency","Urinary frequency (less than 2 hours between voids)","Getting up regularly at night to pass urine","Change in sexual interest, increased or decreased","Difficulty with sexual function","Unexplained pain (whether related to medication, or relieved by antiparkinsonian medication)","Reduced sense of taste or smell","Unexplained change in weight (not explained by diet)","Excessive sweating (not explained by hot weather)"
)

PDQ39_ITEMS <- c(
  "Difficulty doing the leisure activities you normally enjoy?","Difficulty doing household chores, such as cooking or tidying up?","Difficulty carrying bags when you go shopping?","Difficulty walking 1 km on your own?","Difficulty walking 100 m on your own?",
  "Difficulty getting around the house?","Difficulty getting around outside?","Do you need someone to accompany you when you go out?","Are you afraid or worried about falling when you go out?","Do you want to go out but are confined to the house?",
  "Difficulty bathing yourself?","Difficulty dressing yourself?","Difficulty doing up buttons or tying shoelaces?","Difficulty writing clearly?","Difficulty cutting up your food?","Difficulty holding a drink without spilling it?",
  "Feeling depressed?","Feeling lonely and isolated?","Feeling like crying?","Feeling angry or bitter?","Feeling anxious?","Worrying about your future?","Feeling that you need to hide your Parkinson's from other people?",
  "Avoiding eating or drinking in public?","Feeling embarrassed in public because of your Parkinson's?","Worrying about how other people react to you?","Difficulty with close personal relationships?","Lack of support from your spouse or partner when you need it?",
  "Lack of support from family or friends when you need it?","Falling asleep unexpectedly during the day?","Difficulty concentrating, for example when watching TV or reading?","Feeling that your memory is poor?","Having nightmares or hallucinations?","Difficulty with your speech?",
  "Feeling unable to communicate with other people?","Feeling ignored by others?","Muscle cramps, or pain caused by cramps?","Aching or pain in your body or joints?","Feeling uncomfortably hot or cold?"
)

ISI_ITEMS <- c(
  "1A Difficulty falling asleep",
  "1B Difficulty staying asleep",
  "1C Waking up too early",
  "2 Satisfaction with your current sleep pattern",
  "3 Interference of sleep problems with daytime functioning",
  "4 Impact of sleep problems on quality of life",
  "5 Worry or distress about your current sleep problems"
)

ISI_INTERPRET <- function(total_score) {
  dplyr::case_when(
    total_score >= 0  & total_score <= 7  ~ "No clinically significant insomnia",
    total_score >= 8  & total_score <= 14 ~ "Subthreshold insomnia",
    total_score >= 15 & total_score <= 21 ~ "Clinical insomnia, moderate",
    total_score >= 22 & total_score <= 28 ~ "Clinical insomnia, severe",
    TRUE ~ NA_character_
  )
}

# -------------------------
# File paths
# -------------------------
gai_record_file <- "GAI_records.csv"
gai_individual_dir <- "records_gai_individual"

gds_record_file <- "GDS_records.csv"
gds_individual_dir <- "records_gds_individual"

ess_record_file <- "ESS_records.csv"
ess_individual_dir <- "records_ess_individual"

pdss2_record_file <- "PDSS2_records.csv"
pdss2_individual_dir <- "records_pdss2_individual"

scopa_record_file <- "SCOPA_AUT_records.csv"
scopa_individual_dir <- "records_scopa_aut_individual"

updrs1_record_file <- "MDS_UPDRS_I_records.csv"
updrs1_individual_dir <- "records_updrs1_individual"

updrs2_record_file <- "MDS_UPDRS_II_records.csv"
updrs2_individual_dir <- "records_updrs2_individual"

nmss_record_file <- "NMSS_records.csv"
nmss_individual_dir <- "records_nmss_individual"

pdq39_record_file <- "PDQ39_records.csv"
pdq39_individual_dir <- "records_pdq39_individual"

isi_record_file <- "ISI_records.csv"
isi_individual_dir <- "records_isi_individual"

# -------------------------
# Shared UI components
# -------------------------
common_sidebar <- function(prefix, show_visit = TRUE, extra_ui = NULL) {
  tagList(
    selectInput(paste0(prefix, "_pid"), "Participant ID", choices = character(0), selected = NULL),
    dateInput(paste0(prefix, "_date"), "Assessment date", value = Sys.Date()),
    if (show_visit) textInput(paste0(prefix, "_visit"), "Visit / phase", value = "Baseline"),
    if (!is.null(extra_ui)) extra_ui,
    tags$hr(),
    actionButton(paste0(prefix, "_save"), "Calculate and save", class = "btn-primary btn-lg"),
    actionButton(paste0(prefix, "_load"), "Refresh history", class = "btn-lg"),
    actionButton(paste0(prefix, "_reset"), "Reset", class = "btn-lg"),
    tags$hr(),
    downloadButton(paste0(prefix, "_csv"), "Export master CSV", class = "btn-lg"),
    downloadButton(paste0(prefix, "_xlsx"), "Export master Excel", class = "btn-lg")
  )
}

radio_question <- function(id, label, choices, inline = TRUE, selected = character(0)) {
  radioButtons(id, label, choices = choices, selected = selected, inline = inline)
}

# -------------------------
# Shared download handlers
# -------------------------
register_downloads <- function(output, prefix, record_file, scale_type = NULL) {
  output[[paste0(prefix, "_csv")]] <- downloadHandler(
    filename = function() paste0(prefix, "_records_", Sys.Date(), ".csv"),
    content = function(file) {
      df <- if (!is.null(scale_type) && nzchar(scale_type)) read_scale_records_db(scale_type, record_file) else read_records(record_file)
      if (is.null(df) || nrow(df) == 0) {
        write_excel_csv(tibble(Note = "No data yet."), file)
      } else {
        write_excel_csv(df, file)
      }
    }
  )
  
  output[[paste0(prefix, "_xlsx")]] <- downloadHandler(
    filename = function() paste0(prefix, "_records_", Sys.Date(), ".xlsx"),
    content = function(file) {
      df <- if (!is.null(scale_type) && nzchar(scale_type)) read_scale_records_db(scale_type, record_file) else read_records(record_file)
      wb <- createWorkbook()
      addWorksheet(wb, "Records")
      if (is.null(df)) {
        writeData(wb, "Records", data.frame(Note = "No data yet."))
      } else {
        writeData(wb, "Records", df)
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}



# =========================================================
# 4) UPDRS III
# =========================================================
u3_record_file <- "MDS_UPDRS_III_records.csv"
u3_individual_dir <- "records_individual"

u3_items <- tibble(
  item = c("3.1","3.2","3.3a","3.3b","3.3c","3.3d","3.3e","3.4a","3.4b","3.5a","3.5b","3.6a","3.6b","3.7a","3.7b","3.8a","3.8b","3.9","3.10","3.11","3.12","3.13","3.14","3.15a","3.15b","3.16a","3.16b","3.17a","3.17b","3.17c","3.17d","3.17e","3.18"),
  label = c("Speech","Facial expression","Rigidity, neck","Rigidity, left upper limb","Rigidity, right upper limb","Rigidity, left lower limb","Rigidity, right lower limb","Finger tapping, left hand","Finger tapping, right hand","Hand movements, left hand","Hand movements, right hand","Pronation-supination, left hand","Pronation-supination, right hand","Toe tapping, left foot","Toe tapping, right foot","Leg agility, left leg","Leg agility, right leg","Arising from chair","Gait","Freezing of gait","Postural stability","Posture","Global spontaneity of movement","Postural tremor, left hand","Postural tremor, right hand","Kinetic tremor, left hand","Kinetic tremor, right hand","Rest tremor amplitude, left upper limb","Rest tremor amplitude, right upper limb","Rest tremor amplitude, left lower limb","Rest tremor amplitude, right lower limb","Rest tremor amplitude, lip/jaw","Rest tremor constancy")
) %>% mutate(group = case_when(
  item %in% c("3.1","3.2") ~ "Speech/Facial",
  grepl("^3\\.3", item) ~ "Rigidity",
  item %in% c("3.4a","3.4b","3.5a","3.5b","3.6a","3.6b") ~ "Upper limb bradykinesia",
  item %in% c("3.7a","3.7b","3.8a","3.8b") ~ "Lower limb bradykinesia",
  item %in% c("3.9","3.10","3.11","3.12","3.13","3.14") ~ "Gait/Posture/Body",
  item %in% c("3.15a","3.15b","3.16a","3.16b") ~ "Postural/Kinetic tremor",
  grepl("^3\\.17", item) ~ "Rest tremor amplitude",
  item == "3.18" ~ "Rest tremor constancy",
  TRUE ~ "Other"
))

u3_read_records <- function(path) {
  if (!file.exists(path)) return(NULL)
  tryCatch(read_csv(path, show_col_types = FALSE), error = function(e) NULL)
}

db_quote_id <- function(con, x) as.character(dbQuoteIdentifier(con, x))

db_ensure_columns <- function(con, table_name, df) {
  if (!dbExistsTable(con, table_name)) {
    dbWriteTable(con, table_name, df[0, , drop = FALSE], overwrite = TRUE)
    return(invisible(TRUE))
  }
  existing <- dbListFields(con, table_name)
  missing <- setdiff(names(df), existing)
  for (nm in missing) {
    dbExecute(con, paste0(
      "ALTER TABLE ", db_quote_id(con, table_name),
      " ADD COLUMN ", db_quote_id(con, nm), " TEXT"
    ))
  }
  invisible(TRUE)
}

db_align_to_table <- function(con, table_name, df) {
  existing <- dbListFields(con, table_name)
  for (nm in setdiff(existing, names(df))) df[[nm]] <- NA
  df[, existing, drop = FALSE]
}

u3_dedupe_records <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(df)
  key_cols <- intersect(c("pid", "visit_date", "saved_at_local"), names(df))
  if (length(key_cols) == 0) return(df)

  df %>%
    mutate(
      .source_priority = case_when(
        "migration_source_file" %in% names(.) & str_detect(na2empty(migration_source_file), "U3_All_Subjects_Combined") ~ 3L,
        "migration_source_file" %in% names(.) & str_detect(na2empty(migration_source_file), "^MDS_UPDRS_III_records\\.csv$") ~ 2L,
        TRUE ~ 1L
      )
    ) %>%
    arrange(.source_priority) %>%
    distinct(across(all_of(key_cols)), .keep_all = TRUE) %>%
    select(-.source_priority)
}

u3_read_db_records <- function(pid = NULL) {
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
  if (!dbExistsTable(con, "updrs3_records")) return(NULL)
  out <- tryCatch(dbReadTable(con, "updrs3_records") %>% as_tibble(), error = function(e) NULL)
  out <- u3_dedupe_records(out)
  if (!is.null(out) && nrow(out) > 0 && !is.null(pid) && "pid" %in% names(out)) {
    aliases <- id_aliases_for_module(pid, module = "u3")
    out <- out %>% filter(pid %in% aliases)
  }
  out
}

u3_append_db_record <- function(df_row) {
  stopifnot(is.data.frame(df_row), nrow(df_row) == 1)
  con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)

  df_out <- df_row %>%
    mutate(
      migration_source_file = "runtime_shiny",
      migration_source_row = 0L,
      migration_key = paste0("runtime_shiny::", sanitize_filename(pid), "::", format(Sys.time(), "%Y%m%d%H%M%OS3"), "::", sample.int(1e6, 1)),
      migrated_at = now_chr()
    )

  dbWithTransaction(con, {
    db_ensure_columns(con, "updrs3_records", df_out)
    df_out <- db_align_to_table(con, "updrs3_records", df_out)
    dbAppendTable(con, "updrs3_records", df_out)
  })
  invisible(TRUE)
}

u3_append_record <- function(path, df_row) {
  append_record_csv(path, df_row, excel_csv = FALSE)
}
u3_save_individual <- function(df_row, pid, visit_date, out_dir) {
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  pid_dir <- file.path(out_dir, sanitize_filename(pid))
  if (!dir.exists(pid_dir)) dir.create(pid_dir, recursive = TRUE, showWarnings = FALSE)
  file_name <- paste0(sanitize_filename(pid), "_", sanitize_filename(as.character(visit_date)), "_", format(Sys.time(), "%Y-%m-%d_%H-%M-%S"), ".csv")
  out_path <- file.path(pid_dir, file_name)
  write_csv(df_row, out_path)
  out_path
}

derive_scores <- function(df_long) {
  total <- sum(df_long$value, na.rm = TRUE)
  left_items <- c("3.3b","3.4a","3.5a","3.6a","3.7a","3.8a","3.15a","3.16a","3.17a","3.17c")
  right_items <- c("3.3c","3.4b","3.5b","3.6b","3.7b","3.8b","3.15b","3.16b","3.17b","3.17d")
  left_score <- sum(df_long$value[df_long$item %in% left_items], na.rm = TRUE)
  right_score <- sum(df_long$value[df_long$item %in% right_items], na.rm = TRUE)
  rigidity_total <- sum(df_long$value[grepl("^3\\.3", df_long$item)], na.rm = TRUE)
  tremor_total <- sum(df_long$value[df_long$item %in% c("3.15a","3.15b","3.16a","3.16b","3.17a","3.17b","3.17c","3.17d","3.17e","3.18")], na.rm = TRUE)
  group_order <- c(
    "Speech/Facial",
    "Rigidity",
    "Upper limb bradykinesia",
    "Lower limb bradykinesia",
    "Gait/Posture/Body",
    "Postural/Kinetic tremor",
    "Rest tremor amplitude",
    "Rest tremor constancy"
  )
  
  group_sum <- df_long %>%
    mutate(group = factor(group, levels = group_order)) %>%
    group_by(group) %>%
    summarise(score = sum(value, na.rm = TRUE), .groups = "drop") %>%
    arrange(group) %>%
    mutate(group = as.character(group)) %>%
    pivot_wider(names_from = group, values_from = score)
  
  for (nm in setdiff(group_order, names(group_sum))) {
    group_sum[[nm]] <- 0
  }
  
  group_sum <- group_sum[, group_order, drop = FALSE]
  bind_cols(tibble(Total_III = total, Left_sum = left_score, Right_sum = right_score, RminusL = right_score - left_score,
                   Abs_LminusR = abs(left_score - right_score), More_affected_side = case_when(abs(left_score - right_score) < 1 ~ "Symmetric", left_score > right_score ~ "Left more affected", right_score > left_score ~ "Right more affected", TRUE ~ "Symmetric"),
                   Rigidity_sum = rigidity_total, Tremor_sum = tremor_total), group_sum)
}

get_previous_row <- function(df_all, pid) {
  if (is.null(df_all) || nrow(df_all) == 0 || !"pid" %in% names(df_all)) return(NULL)
  aliases <- id_aliases_for_module(pid, module = "u3")
  dfp <- df_all %>% filter(pid %in% aliases)
  if (nrow(dfp) < 1) return(NULL)
  if ("visit_date" %in% names(dfp)) dfp$visit_date <- suppressWarnings(as.Date(dfp$visit_date)) else dfp$visit_date <- as.Date(NA)
  if ("saved_at_local" %in% names(dfp)) {
    dfp$saved_at_local_dt <- suppressWarnings(ymd_hms(dfp$saved_at_local))
    if (all(is.na(dfp$saved_at_local_dt))) dfp$saved_at_local_dt <- suppressWarnings(parse_date_time(dfp$saved_at_local, orders = c("Y-m-d H:M:S", "Y/m/d H:M:S", "Y-m-d H:M", "Y/m/d H:M")))
  } else dfp$saved_at_local_dt <- as.POSIXct(NA)
  dfp %>% arrange(visit_date, saved_at_local_dt) %>% slice_tail(n = 1)
}

calc_delta <- function(current_summary_row, prev_row) {
  if (is.null(prev_row)) return(NULL)
  common <- intersect(names(current_summary_row), names(prev_row))
  common_num <- common[sapply(common, function(nm) is.numeric(current_summary_row[[nm]]) && is.numeric(prev_row[[nm]]))]
  if (length(common_num) == 0) return(NULL)
  tibble(metric = common_num, delta = sapply(common_num, function(nm) as.numeric(current_summary_row[[nm]]) - as.numeric(prev_row[[nm]]))) %>% arrange(desc(abs(delta)))
}
blank_plot <- function(title_text) ggplot(data.frame(x = 1, y = 1), aes(x, y)) + geom_blank() + labs(title = title_text) + theme_void()

# Non-motor symptom history plot: display only, saving logic unchanged
scale_metric_choices <- list(
  gai = c("Total_GAI" = "Total_GAI", "Positive_n" = "Positive_n"),
  gds = c("Total_GDS" = "Total_GDS"),
  ess = c("Total_ESS" = "Total_ESS"),
  pdss2 = c("Total_PDSS2" = "Total_PDSS2"),
  scopa = c("Total_SCOPA_AUT" = "Total_SCOPA_AUT", "N_items" = "N_items"),
  updrs1 = c("Total_UPDRS_I" = "Total_UPDRS_I"),
  updrs2 = c("Total_UPDRS_II" = "Total_UPDRS_II"),
  nmss = c("Total_NMSS" = "Total_NMSS", "Positive_items" = "Positive_items"),
  pdq39 = c("Total_PDQ39" = "Total_PDQ39"),
  isi = c("Total_ISI" = "Total_ISI")
)

render_scale_history_plot_ui <- function(prefix, title = "History") {
  tagList(
    tags$hr(),
    h4(title),
    selectInput(paste0(prefix, "_hist_metric"), "Select measure",
                choices = scale_metric_choices[[prefix]],
                selected = unname(scale_metric_choices[[prefix]][1])),
    plotOutput(paste0(prefix, "_hist_plot"), height = 300)
  )
}


# =========================================================
# 5) UI
# =========================================================
ui <- fluidPage(
  theme = bs_theme(version = 5, bootswatch = "flatly"),
  tags$head(tags$style(HTML(" 
      .small-note { color:#666; font-size:12px; }
      .card-like { background:#fff; border:1px solid #E5E7EB; border-radius:14px; padding:14px; margin-bottom:12px; }
      .slot-box { border: 1px solid #ddd; border-radius: 8px; padding: 10px; margin-bottom: 8px; background: #fafafa; }
      body { font-size: 17px; }
      .u3-score-group { margin-bottom: 14px; }
      .u3-score-group .shiny-options-group { display:flex; flex-wrap:wrap; column-gap:28px; row-gap:10px; }
      .u3-score-group .form-check-inline { margin-right: 0 !important; }
      .u3-score-group .control-label { font-weight: 600; margin-bottom: 8px; display:block; }
  "))),
  titlePanel("TIS-PD Clinical Management Suite"),
  sidebarLayout(
    sidebarPanel(width = 3,
      h4("Add participant"),
      textInput("subject_id", "Participant ID (e.g. PD001)"),
      textInput("initials", "Initials"),
      textInput("full_name", "Name"),
      selectInput("sex", "Sex", choices = c("Not recorded" = "", "Male" = "Male", "Female" = "Female", "Other" = "Other"), selected = ""),
      textInput("birth_month", "Month and year of birth", placeholder = "e.g. 1965-03"),
      fluidRow(
        column(6, numericInput("height_cm", "Height (cm)", value = NA, min = 0, step = 0.1)),
        column(6, numericInput("weight_kg", "Weight (kg)", value = NA, min = 0, step = 0.1))
      ),
      textInput("education", "Education"),
      textInput("occupation", "Occupation"),
      numericInput("disease_years", "Disease duration (years)", value = NA, min = 0, step = 0.1),
      dateInput("visit1_date", "Visit 1 planned date", value = Sys.Date()),
      checkboxInput("init_visit1_done", "Visit 1 completed", FALSE),
      conditionalPanel("input.init_visit1_done == true", airDatepickerInput("visit1_done_date", "Visit 1 completion date", value = NULL, clearButton = TRUE, autoClose = TRUE)),
      dateInput("visit2_date", "Visit 2 planned date", value = Sys.Date() + 1),
      checkboxInput("init_visit2_done", "Visit 2 completed", FALSE),
      conditionalPanel("input.init_visit2_done == true", airDatepickerInput("visit2_done_date", "Visit 2 completion date", value = NULL, clearButton = TRUE, autoClose = TRUE)),
      textAreaInput("subject_note", "Notes", rows = 3),
      checkboxInput("is_dropout_new", "Participant withdrawn", FALSE),
      conditionalPanel("input.is_dropout_new == true",
                       airDatepickerInput("dropout_date_new", "Withdrawal date", value = NULL, clearButton = TRUE, autoClose = TRUE),
                       textAreaInput("dropout_reason_new", "Reason for withdrawal", rows = 2)),
      numericInput("init_home_done_n", "Home sessions completed (0–27)", 0, min = 0, max = 27),
      checkboxInput("init_visit3_done", "Visit 3 completed", FALSE),
      checkboxInput("init_fu4_done", "4-week follow-up completed", FALSE),
      checkboxInput("init_fu8_done", "8-week follow-up completed", FALSE),
      actionButton("add_subject", "Save and generate schedule", class = "btn-primary"),
      hr(),
      dateInput("filter_date", "View tasks for a date", Sys.Date()),
      textInput("operator_name", "Recorded by / operator", value = ""),
      hr(),
      h4("Delete participant"),
      selectInput("delete_subject_id", "Select participant", choices = character(0)),
      actionButton("delete_subject_btn", "Permanently delete this participant", class = "btn-danger"),
      div(class = "small-note", "All of this participant's task records will also be deleted. This cannot be undone.")
    ),
    mainPanel(width = 9,
      tabsetPanel(
        tabPanel("Participants", br(),
          div(class = "card-like", h4("All participants"), actionButton("edit_subject_btn", "Edit selected participant", class = "btn-info"), br(), br(), DTOutput("subjects_table")),
          div(class = "card-like", h4(textOutput("selected_subject_header")), uiOutput("selected_subject_calendar")),
          div(class = "card-like", h4("All tasks for this participant (click a row to edit)"), fluidRow(column(6, actionButton("edit_subject_task_btn", "Edit selected task", class = "btn-warning"))), br(), DTOutput("subject_detail_table"))
        ),
        tabPanel("Today's tasks", br(),
                 div(class = "card-like",
                     h4(textOutput("task_header")),
                     fluidRow(
                       column(6, actionButton("edit_daily_task_btn", "Edit selected task", class = "btn-warning")),
                       column(6, actionButton("complete_daily_home_btn", "Complete all of today's sessions (excluding assessments)", icon = icon("check"), class = "btn-success"))
                     ),
                     br(),
                     DTOutput("daily_tasks")),
                 div(class = "card-like",
                     h4(textOutput("tomorrow_task_header")),
                     DTOutput("tomorrow_tasks")),
                 div(class = "card-like", h4("Overdue tasks (click a row to change status)"), DTOutput("overdue_tasks"))),
        tabPanel("Home diary", br(),
                 div(class = "card-like",
                     h4("Home diary entry"),
                     fluidRow(column(4, selectInput("diary_pid", "Participant ID", choices = character(0), selected = NULL)), column(4, dateInput("diary_date", "Date", value = Sys.Date())), column(4, sliderInput("self_score", "Today's self-rating (0–100)", min = 0, max = 100, value = 80, step = 1))),
                     radioButtons("fall_event", "Did you fall today?", choices = c("No" = 0, "Yes" = 1), selected = 0, inline = TRUE),
                     textAreaInput("fall_detail", "If you fell, give the time and circumstances", value = "", rows = 2),
                     radioButtons("adverse_event", "Did anything else adverse happen today?", choices = c("No" = 0, "Yes" = 1), selected = 0, inline = TRUE),
                     textAreaInput("adverse_detail", "If so, give the time and circumstances", value = "", rows = 2),
                     radioButtons("mood_score", "Rate your overall mood today", choices = c("Very poor" = 1, "Poor" = 2, "Fair" = 3, "Good" = 4, "Very good" = 5), selected = 3),
                     radioButtons("exercise", "Did you do any physical exercise or activity today?", choices = c("No" = 0, "Yes" = 1), selected = 0, inline = TRUE),
                     textAreaInput("exercise_detail", "If so, note the type and duration", value = "", rows = 2),
                     radioButtons("constipation", "Did you have constipation or difficulty passing stools today?", choices = c("No" = 0, "Yes" = 1), selected = 0, inline = TRUE),
                     textInput("med_onset_mins", "Time to medication onset (min)", value = ""),
                     fluidRow(column(3, selectInput("night_sleep_start", "Night sleep start time", choices = time_choices, selected = "21:00")),
                              column(3, selectInput("night_sleep_end", "Night sleep end time", choices = time_choices, selected = "06:00")),
                              column(3, selectInput("nap_start", "Nap start time (optional)", choices = c("Not recorded" = "", time_choices), selected = "")),
                              column(3, selectInput("nap_end", "Nap end time (optional)", choices = c("Not recorded" = "", time_choices), selected = ""))),
                     fluidRow(column(6, selectInput("awake_start", "Awake period start", choices = time_choices, selected = "06:00")), column(6, selectInput("awake_end", "Awake period end", choices = time_choices, selected = "22:00"))),
                     textAreaInput("other_symptoms", "Any other discomfort or unusual symptoms", value = "", rows = 3),
                     actionButton("diary_calc_save", "Calculate and save", class = "btn-primary"),
                     actionButton("diary_load_hist", "Refresh history"),
                     actionButton("diary_reset_all", "Reset"),
                     hr(), h4("Half-hourly state chart for awake periods"), uiOutput("slot_inputs"),
                     hr(), h4("Calculated automatically"), tableOutput("summary_table"),
                     hr(), h4("History (current participant)"),
                     selectInput("diary_hist_metric", "Select measure", choices = c(
                       "Self-rating" = "self_score",
                       "Good control (% of awake time)" = "good_control_pct",
                       "Non-troublesome dyskinesia (% of awake time)" = "non_troublesome_dyskinesia_pct",
                       "Troublesome dyskinesia (% of awake time)" = "troublesome_dyskinesia_pct",
                       "Severe discomfort (% of awake time)" = "very_uncomfortable_pct"
                     ), selected = "self_score"),
                     plotOutput("hist_plot", height = 320))),
        tabPanel("Non-motor symptoms", br(),
                 tabsetPanel(
tabPanel("GAI",
           sidebarLayout(
             sidebarPanel(common_sidebar("gai")),
             mainPanel(
               h4("GAI (20 items, yes/no)"),
               uiOutput("gai_questions"),
               tags$hr(), h4("Results"), tableOutput("gai_summary"),
               tags$hr(), h4("History"), tableOutput("gai_hist"),
               render_scale_history_plot_ui("gai")
             )
           )
  ),
  
  tabPanel("GDS",
           sidebarLayout(
             sidebarPanel(common_sidebar("gds")),
             mainPanel(
               h4("GDS (30 items, yes/no)"),
               uiOutput("gds_questions"),
               tags$hr(), h4("Results"), tableOutput("gds_summary"),
               tags$hr(), h4("History"), tableOutput("gds_hist"),
               render_scale_history_plot_ui("gds")
             )
           )
  ),
  
  tabPanel("ESS",
           sidebarLayout(
             sidebarPanel(common_sidebar("ess")),
             mainPanel(
               h4("ESS (8 items, scored 0–3)"),
               uiOutput("ess_questions"),
               tags$hr(), h4("Results"), tableOutput("ess_summary"),
               tags$hr(), h4("History"), tableOutput("ess_hist"),
               render_scale_history_plot_ui("ess")
             )
           )
  ),
  
  tabPanel("PDSS-2",
           sidebarLayout(
             sidebarPanel(common_sidebar("pdss2")),
             mainPanel(
               h4("PDSS-2 (15 items, scored 0–4)"),
               uiOutput("pdss2_questions"),
               tags$hr(), h4("Results"), tableOutput("pdss2_summary"),
               tags$hr(), h4("History"), tableOutput("pdss2_hist"),
               render_scale_history_plot_ui("pdss2")
             )
           )
  ),
  
  tabPanel("SCOPA-AUT",
           sidebarLayout(
             sidebarPanel(
               common_sidebar("scopa", extra_ui = tagList(
                 selectInput("scopa_sex", "Sex", choices = c("Not recorded" = "", "Male" = "Male", "Female" = "Female"), selected = "")
               ))
             ),
             mainPanel(
               h4("SCOPA-AUT (scored 0–3)"),
               uiOutput("scopa_questions"),
               tags$hr(), h4("Results"), tableOutput("scopa_summary"),
               tags$hr(), h4("History"), tableOutput("scopa_hist"),
               render_scale_history_plot_ui("scopa")
             )
           )
  ),
  
  tabPanel("MDS-UPDRS I",
           sidebarLayout(
             sidebarPanel(common_sidebar("updrs1")),
             mainPanel(
               h4("MDS-UPDRS Part I (scored 0–4)"),
               uiOutput("updrs1_questions"),
               tags$hr(), h4("Results"), tableOutput("updrs1_summary"),
               tags$hr(), h4("History"), tableOutput("updrs1_hist"),
               render_scale_history_plot_ui("updrs1")
             )
           )
  ),
  
  tabPanel("MDS-UPDRS II",
           sidebarLayout(
             sidebarPanel(common_sidebar("updrs2")),
             mainPanel(
               h4("MDS-UPDRS Part II (scored 0–4)"),
               uiOutput("updrs2_questions"),
               tags$hr(), h4("Results"), tableOutput("updrs2_summary"),
               tags$hr(), h4("History"), tableOutput("updrs2_hist"),
               render_scale_history_plot_ui("updrs2")
             )
           )
  ),
  
  tabPanel("NMSS",
           sidebarLayout(
             sidebarPanel(common_sidebar("nmss")),
             mainPanel(
               h4("NMSS (answer no/yes; if yes, rate severity and frequency)"),
               uiOutput("nmss_questions"),
               tags$hr(), h4("Results"), tableOutput("nmss_summary"),
               tags$hr(), h4("History"), tableOutput("nmss_hist"),
               render_scale_history_plot_ui("nmss")
             )
           )
  ),
  
  tabPanel("PDQ-39",
           sidebarLayout(
             sidebarPanel(common_sidebar("pdq39")),
             mainPanel(
               h4("PDQ-39 (scored 0–4)"),
               uiOutput("pdq39_questions"),
               tags$hr(), h4("Results"), tableOutput("pdq39_summary"),
               tags$hr(), h4("History"), tableOutput("pdq39_hist"),
               render_scale_history_plot_ui("pdq39")
             )
           )
  ),
  
  tabPanel("ISI",
           sidebarLayout(
             sidebarPanel(common_sidebar("isi")),
             mainPanel(
               h4("ISI — Insomnia Severity Index (7 items, scored 0–4)"),
               helpText("Total 0–28: 0–7 no clinically significant insomnia; 8–14 subthreshold insomnia; 15–21 moderate clinical insomnia; 22–28 severe clinical insomnia."),
               uiOutput("isi_questions"),
               tags$hr(), h4("Results"), tableOutput("isi_summary"),
               tags$hr(), h4("History"), tableOutput("isi_hist"),
               render_scale_history_plot_ui("isi")
             )
           )
  )
                 )
        ),
        tabPanel("MDS-UPDRS-III", br(),
                 div(class = "card-like",
                     fluidRow(column(3, selectInput("u3_pid", "Participant ID", choices = character(0), selected = NULL)), column(3, dateInput("u3_visit_date", "Assessment date", value = Sys.Date())), column(3, textInput("u3_visit", "Visit / phase", value = "Baseline")), column(3, selectInput("med_state", "Medication state", choices = c("NA","ON","OFF"), selected = "NA"))),
                     fluidRow(column(3, selectInput("dbs_state", "DBS state", choices = c("NA","ON","OFF"), selected = "NA")), column(3, radioButtons("dyskinesia_present", "Dyskinesia present?", choices = c("0","1"), selected = "0", inline = TRUE)), column(3, radioButtons("dyskinesia_interfere", "Dyskinesia interfering with scoring?", choices = c("0","1"), selected = "0", inline = TRUE))),
                     uiOutput("u3_question_inputs"),
                     hr(), actionButton("u3_calc_save", "Calculate and save (database + local backup)", class = "btn-primary"), actionButton("u3_load_hist", "Refresh / load history"), actionButton("u3_reset_all", "Set all items to 0"),
                     hr(), h4("Current score (Part III)"), tableOutput("u3_score_table"),
                     hr(), h4("Dominant side (this assessment)"), tableOutput("u3_side_table"),
                     hr(), h4("Current vs previous: difference plot"), plotOutput("u3_delta_plot", height = 380),
                     hr(), h4("History (current participant)"),
                     selectInput("u3_hist_metric", "Select measure", choices = c("Total III" = "Total_III", "Left Sum" = "Left_sum", "Right Sum" = "Right_sum", "Rigidity Sum" = "Rigidity_sum", "Tremor Sum" = "Tremor_sum"), selected = "Total_III"),
                     plotOutput("u3_hist_plot", height = 320)
                 )
        ),
        tabPanel("Export", br(), div(class = "card-like", h4("Export data"), downloadButton("download_excel", "Export Excel"), tags$span("  "), downloadButton("download_csv", "Export CSV")))
      )
    )
  )
)

# =========================================================
# 6) Server
# =========================================================
server <- function(input, output, session) {
  init_db()
  subjects_rv <- reactiveVal(read_subjects())
  visits_rv <- reactiveVal(read_visits())
  editing_subject_id <- reactiveVal(NULL)
  current_subject_id <- reactiveVal(NULL)
  editing_event_id <- reactiveVal(NULL)
  selected_subject_task_id <- reactiveVal(NULL)
  selected_daily_task_id <- reactiveVal(NULL)
  pending_daily_home <- reactiveVal(NULL)

  dt_display_start <- function(rows_current, n_rows) {
    if (is.null(n_rows) || is.na(n_rows) || n_rows <= 0) return(0L)
    rows_current <- suppressWarnings(as.integer(rows_current))
    rows_current <- rows_current[!is.na(rows_current)]
    if (length(rows_current) == 0) return(0L)
    min(max(min(rows_current) - 1L, 0L), max(n_rows - 1L, 0L))
  }

  active_task_subject_ids <- reactive({
    s_all <- subjects_rv()
    if (is.null(s_all) || nrow(s_all) == 0) return(character(0))
    s_all %>%
      mutate(
        is_active = if ("is_active" %in% names(.)) replace_na(as.integer(is_active), 1L) else 1L,
        is_dropout = if ("is_dropout" %in% names(.)) replace_na(as.integer(is_dropout), 0L) else 0L
      ) %>%
      filter(is_active == 1L, is_dropout != 1L) %>%
      pull(subject_id)
  })

  add_subject_names <- function(df) {
    s_all <- subjects_rv()
    if (is.null(df) || nrow(df) == 0 || is.null(s_all) || nrow(s_all) == 0) {
      if (!"full_name" %in% names(df)) df$full_name <- character(0)
      return(df)
    }
    if (!"full_name" %in% names(s_all)) s_all$full_name <- NA_character_
    df %>%
      left_join(
        s_all %>% transmute(subject_id = as.character(subject_id), full_name = na2empty(full_name)),
        by = "subject_id"
      )
  }

  refresh_all <- function() {
    subjects_rv(read_subjects())
    visits_rv(read_visits())
    s_all <- subjects_rv()
    updateSelectInput(session, "delete_subject_id", choices = s_all$subject_id)
    if (is.null(current_subject_id()) && nrow(s_all) > 0) current_subject_id(s_all$subject_id[1])
    if (!is.null(current_subject_id()) && !(current_subject_id() %in% s_all$subject_id)) current_subject_id(if (nrow(s_all) > 0) s_all$subject_id[1] else NULL)
  }
  observeEvent(TRUE, refresh_all(), once = TRUE)

  # ---- Sync the current participant to the sub-module selectors ----
  observe({
    s_all <- subjects_rv()
    ids <- if (is.null(s_all) || nrow(s_all) == 0) character(0) else s_all$subject_id
    sid <- current_subject_id()
    default_sid <- if (!is.null(sid) && !is.na(sid) && sid %in% ids) sid else if (length(ids) > 0) ids[1] else character(0)

    updateSelectInput(session, "diary_pid", choices = ids, selected = default_sid)
    updateSelectInput(session, "u3_pid", choices = ids, selected = default_sid)
    for (pref in c("gai","gds","ess","pdss2","scopa","updrs1","updrs2","nmss","pdq39","isi")) {
      updateSelectInput(session, paste0(pref, "_pid"), choices = ids, selected = default_sid)
    }
  })

  # ---- Add participant ----
  observeEvent(input$is_dropout_new, {
    if (!isTRUE(input$is_dropout_new)) {
      updateAirDateInput(session, "dropout_date_new", value = NULL)
      updateTextAreaInput(session, "dropout_reason_new", value = "")
    }
  })
  observeEvent(input$init_visit1_done, if (!isTRUE(input$init_visit1_done)) updateAirDateInput(session, "visit1_done_date", value = NULL))
  observeEvent(input$init_visit2_done, if (!isTRUE(input$init_visit2_done)) updateAirDateInput(session, "visit2_done_date", value = NULL))

  observeEvent(input$add_subject, {
    req(input$subject_id, input$visit1_date, input$visit2_date)
    sid <- str_trim(input$subject_id)
    ini <- str_trim(input$initials)
    if (sid == "") return(showNotification("Participant ID cannot be empty.", type = "error"))
    if (as.Date(input$visit2_date) < as.Date(input$visit1_date)) return(showNotification("The Visit 2 planned date cannot be earlier than the Visit 1 planned date.", type = "error"))
    tryCatch({
      insert_subject_with_schedule(
        sid, ini,
        input$full_name, input$sex, input$birth_month,
        input$height_cm, input$weight_kg,
        input$education, input$occupation, input$disease_years,
        input$visit1_date, input$visit2_date, input$subject_note
      )
      initialize_subject_progress(
        sid,
        visit1_done = isTRUE(input$init_visit1_done),
        visit1_actual_date = if (isTRUE(input$init_visit1_done)) input$visit1_done_date else NA,
        visit2_done = isTRUE(input$init_visit2_done),
        visit2_actual_date = if (isTRUE(input$init_visit2_done)) input$visit2_done_date else NA,
        home_done_n = input$init_home_done_n,
        visit3_done = isTRUE(input$init_visit3_done),
        fu4_done = isTRUE(input$init_fu4_done),
        fu8_done = isTRUE(input$init_fu8_done)
      )
      if (isTRUE(input$is_dropout_new)) {
        apply_subject_dropout(sid, TRUE, input$dropout_date_new, input$dropout_reason_new)
      }
      recalc_subject_schedule(sid)
      refresh_all(); current_subject_id(sid)
      showNotification("Participant added.", type = "message")
    }, error = function(e) showNotification(e$message, type = "error"))
  })

  # ---- Main table ----
  output$subjects_table <- renderDT({
    df <- subjects_rv()
    display_cols <- c(
      "subject_id", "full_name", "sex", "birth_month", "height_cm", "weight_kg",
      "education", "occupation", "disease_years", "initials", "visit1_date",
      "visit2_date", "is_dropout", "dropout_date", "dropout_reason", "note",
      "created_at", "is_active"
    )
    datatable(
      df %>% mutate(
        is_active = ifelse(is.na(is_active), 1L, as.integer(is_active)),
        visit1_date = as.character(visit1_date),
        visit2_date = as.character(visit2_date),
        dropout_date = as.character(dropout_date)
      ) %>% select(any_of(display_cols), everything()),
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 10, scrollX = TRUE, stateSave = TRUE, displayStart = dt_display_start(isolate(input$subjects_table_rows_current), nrow(df)))
    )
  }, server = FALSE)
  observeEvent(input$subjects_table_rows_selected, {
    idx <- input$subjects_table_rows_selected
    df <- subjects_rv()
    if (length(idx) == 1 && nrow(df) >= idx) current_subject_id(df$subject_id[idx])
  })

  observeEvent(current_subject_id(), {
    selected_subject_task_id(NA_integer_)
  }, ignoreInit = TRUE)

  observeEvent(input$filter_date, {
    selected_daily_task_id(NA_integer_)
  }, ignoreInit = TRUE)

  output$selected_subject_header <- renderText({ paste0("Current participant: ", current_subject_id() %||% "None selected") })
  subject_detail_df <- reactive({
    sid <- current_subject_id(); req(sid)
    visits_rv() %>%
      filter(subject_id == sid) %>%
      mutate(planned_date = as.character(planned_date), actual_date = as.character(actual_date), display_date = as.character(display_date))
  })
  daily_tasks_df <- reactive({
    d <- as.Date(input$filter_date)
    visits_rv() %>%
      filter(display_date == d, subject_id %in% active_task_subject_ids()) %>%
      mutate(planned_date = as.character(planned_date), actual_date = as.character(actual_date), display_date = as.character(display_date)) %>%
      add_subject_names()
  })
  observeEvent(input$complete_daily_home_btn, {
    d <- as.Date(input$filter_date)
    if (length(d) != 1L || is.na(d) || d > Sys.Date()) {
      return(showNotification("Select today or a past date. Future tasks cannot be completed in a batch.", type = "warning"))
    }
    refresh_all()
    tasks <- daily_tasks_df() %>%
      filter(event_type == "home", status %in% c("planned", "rescheduled"))
    if (nrow(tasks) == 0L) {
      return(showNotification("There are no outstanding home sessions for this date.", type = "message"))
    }
    pending_daily_home(list(date = d, ids = tasks$id))
    showModal(modalDialog(
      title = "Confirm that all of today's sessions are complete",
      p(paste0("Date: ", d, "; ", nrow(tasks), " home session(s) in total. Please confirm that all of them were actually completed.")),
      p("This does not include the assessments before and after the first session, other assessments, follow-ups, paused tasks or tasks already marked as not completed, and it does not change the schedule that follows."),
      tags$ul(lapply(seq_len(nrow(tasks)), function(i) {
        tags$li(paste(tasks$subject_id[i], tasks$full_name[i], tasks$event_name[i]))
      })),
      footer = tagList(modalButton("Cancel"), actionButton("confirm_daily_home_done", "Confirm all complete", class = "btn-success")),
      easyClose = FALSE
    ))
  })
  observeEvent(input$confirm_daily_home_done, {
    pending <- pending_daily_home()
    req(pending)
    pending_daily_home(NULL)
    tryCatch({
      con <- get_con()
      n <- tryCatch(complete_daily_home_tasks(con, pending$date, pending$ids, input$operator_name),
                    finally = dbDisconnect(con))
      removeModal()
      refresh_all()
      showNotification(paste0("Marked ", n, " home session(s) for this date as completed."), type = "message")
    }, error = function(e) {
      removeModal()
      showNotification(paste0("Batch completion failed: ", e$message), type = "error")
    })
  })

  tomorrow_tasks_df <- reactive({
    start_d <- as.Date(input$filter_date) + 1
    end_d <- as.Date(input$filter_date) + 3
    visits_rv() %>%
      filter(
        display_date >= start_d,
        display_date <= end_d,
        subject_id %in% active_task_subject_ids(),
        coalesce(event_type, "") %in% c("onsite", "online", "followup"),
        !(status %in% c("done", "dropout"))
      ) %>%
      mutate(planned_date = as.character(planned_date), actual_date = as.character(actual_date), display_date = as.character(display_date)) %>%
      add_subject_names()
  })
  overdue_tasks_df <- reactive({
    visits_rv() %>%
      filter(
        display_date < Sys.Date(),
        subject_id %in% active_task_subject_ids(),
        coalesce(event_type, "") != "pause",
        !str_detect(na2empty(note), "^\\[AUTO_PAUSE\\]"),
        !(status %in% c("done", "dropout"))
      ) %>%
      mutate(planned_date = as.character(planned_date), actual_date = as.character(actual_date), display_date = as.character(display_date)) %>%
      add_subject_names()
  })

  observeEvent(input$subject_detail_table_rows_selected, {
    idx <- input$subject_detail_table_rows_selected
    selected_subject_task_id(get_selected_visit_id(subject_detail_df(), idx))
  }, ignoreInit = TRUE)

  observeEvent(input$daily_tasks_rows_selected, {
    idx <- input$daily_tasks_rows_selected
    selected_id <- get_selected_visit_id(daily_tasks_df(), idx)
    selected_daily_task_id(selected_id)
    if (length(selected_id) != 1 || is.na(selected_id)) return(invisible(NULL))
    open_event_modal(selected_id)
  }, ignoreInit = TRUE)

  observeEvent(input$overdue_tasks_rows_selected, {
    idx <- input$overdue_tasks_rows_selected
    selected_id <- get_selected_visit_id(overdue_tasks_df(), idx)
    if (length(selected_id) != 1 || is.na(selected_id)) return(invisible(NULL))
    open_event_modal(selected_id)
  }, ignoreInit = TRUE)

  output$selected_subject_calendar <- renderUI({
    sid <- current_subject_id(); req(sid)
    make_full_subject_calendar(visits_rv() %>% filter(subject_id == sid))
  })
  output$subject_detail_table <- renderDT({
    df <- subject_detail_df()
    datatable(
      df %>% select(id, event_name, event_type, session_no, planned_date, status, actual_date, note, completed_by),
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 20, scrollX = TRUE, stateSave = TRUE, displayStart = dt_display_start(isolate(input$subject_detail_table_rows_current), nrow(df)))
    )
  }, server = FALSE)
  output$task_header <- renderText({ paste0("Date: ", as.character(input$filter_date)) })
  output$daily_tasks <- renderDT({
    df <- daily_tasks_df()
    datatable(
      df %>% select(id, subject_id, full_name, event_name, event_type, session_no, planned_date, status, actual_date, note),
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 15, scrollX = TRUE, stateSave = TRUE, displayStart = dt_display_start(isolate(input$daily_tasks_rows_current), nrow(df)))
    )
  }, server = FALSE)
  output$tomorrow_task_header <- renderText({
    paste0("Assessment tasks in the next three days: ", as.character(as.Date(input$filter_date) + 1), " to ", as.character(as.Date(input$filter_date) + 3))
  })
  output$tomorrow_tasks <- renderDT({
    df <- tomorrow_tasks_df()
    datatable(
      df %>% select(id, subject_id, full_name, event_name, event_type, session_no, planned_date, status, actual_date, note),
      rownames = FALSE,
      options = list(pageLength = 15, scrollX = TRUE, stateSave = TRUE, displayStart = dt_display_start(isolate(input$tomorrow_tasks_rows_current), nrow(df)))
    )
  }, server = FALSE)
  output$overdue_tasks <- renderDT({
    df <- overdue_tasks_df()
    datatable(
      df %>% select(id, subject_id, full_name, event_name, event_type, session_no, planned_date, status, actual_date, note),
      selection = "single",
      rownames = FALSE,
      options = list(pageLength = 15, scrollX = TRUE, stateSave = TRUE, displayStart = dt_display_start(isolate(input$overdue_tasks_rows_current), nrow(df)))
    )
  }, server = FALSE)

  # ---- Edit participant ----
  observeEvent(input$edit_subject_btn, {
    sid <- current_subject_id(); req(sid)
    subj <- subjects_rv() %>% filter(subject_id == sid)
    v1 <- visits_rv() %>% filter(subject_id == sid, event_name == "OFF-BASELINE (Visit 1)") %>% slice(1)
    v2 <- visits_rv() %>% filter(subject_id == sid, event_name == "ON-1 PRE/POST (Visit 2) | First stimulation session") %>% slice(1)
    editing_subject_id(sid)
    showModal(modalDialog(
      title = paste0("Edit participant: ", sid), size = "l", easyClose = TRUE,
      textInput("edit_subject_id", "Participant ID", value = subj$subject_id[1]),
      textInput("edit_initials", "Initials", value = na2empty(subj$initials[1])),
      textInput("edit_full_name", "Name", value = na2empty(subj$full_name[1])),
      selectInput("edit_sex", "Sex", choices = c("Not recorded" = "", "Male" = "Male", "Female" = "Female", "Other" = "Other"), selected = scalar_chr(subj$sex, "")),
      textInput("edit_birth_month", "Month and year of birth", value = na2empty(subj$birth_month[1]), placeholder = "e.g. 1965-03"),
      fluidRow(
        column(6, numericInput("edit_height_cm", "Height (cm)", value = numeric_or_na(subj$height_cm[1]), min = 0, step = 0.1)),
        column(6, numericInput("edit_weight_kg", "Weight (kg)", value = numeric_or_na(subj$weight_kg[1]), min = 0, step = 0.1))
      ),
      textInput("edit_education", "Education", value = na2empty(subj$education[1])),
      textInput("edit_occupation", "Occupation", value = na2empty(subj$occupation[1])),
      numericInput("edit_disease_years", "Disease duration (years)", value = numeric_or_na(subj$disease_years[1]), min = 0, step = 0.1),
      dateInput("edit_visit1_date", "Visit 1 planned date", value = subj$visit1_date[1]),
      checkboxInput("edit_visit1_done", "Visit 1 completed", identical(scalar_chr(v1$status, "planned"), "done")),
      conditionalPanel("input.edit_visit1_done == true", airDatepickerInput("edit_visit1_done_date", "Visit 1 completion date", value = if (nrow(v1)==1) v1$actual_date[1] else NULL, clearButton = TRUE, autoClose = TRUE)),
      dateInput("edit_visit2_date", "Visit 2 planned date", value = subj$visit2_date[1]),
      checkboxInput("edit_visit2_done", "Visit 2 completed", identical(scalar_chr(v2$status, "planned"), "done")),
      conditionalPanel("input.edit_visit2_done == true", airDatepickerInput("edit_visit2_done_date", "Visit 2 completion date", value = if (nrow(v2)==1) v2$actual_date[1] else NULL, clearButton = TRUE, autoClose = TRUE)),
      textAreaInput("edit_subject_note", "Notes", value = na2empty(subj$note[1]), rows = 3),
      checkboxInput("edit_is_dropout", "Participant withdrawn", isTRUE(subj$is_dropout[1] == 1)),
      conditionalPanel("input.edit_is_dropout == true", airDatepickerInput("edit_dropout_date", "Withdrawal date", value = subj$dropout_date[1], clearButton = TRUE, autoClose = TRUE), textAreaInput("edit_dropout_reason", "Reason for withdrawal", value = na2empty(subj$dropout_reason[1]), rows = 2)),
      checkboxInput("edit_recalc_subject", "Regenerate this participant's schedule after saving", TRUE),
      footer = tagList(modalButton("Cancel"), actionButton("save_subject_edit", "Save changes", class = "btn-primary"))
    ))
  })

  observeEvent(input$save_subject_edit, {
    old_sid <- editing_subject_id(); req(old_sid)
    tryCatch({
      new_sid <- str_trim(input$edit_subject_id)
      update_subject_info(
        old_sid, new_sid, input$edit_initials,
        input$edit_full_name, input$edit_sex, input$edit_birth_month,
        input$edit_height_cm, input$edit_weight_kg,
        input$edit_education, input$edit_occupation, input$edit_disease_years,
        input$edit_visit1_date, input$edit_visit2_date, input$edit_subject_note
      )
      con <- get_con(); on.exit(dbDisconnect(con), add = TRUE)
      v1_actual <- if (isTRUE(input$edit_visit1_done)) safe_date(input$edit_visit1_done_date %||% input$edit_visit1_date) else NA
      dbExecute(con, "UPDATE visits SET planned_date = ?, status = ?, actual_date = ?, updated_at = ? WHERE subject_id = ? AND event_name = 'OFF-BASELINE (Visit 1)'",
                params = list(as.character(safe_date(input$edit_visit1_date)), ifelse(isTRUE(input$edit_visit1_done), "done", "planned"), ifelse(is.na(v1_actual), NA, as.character(v1_actual)), now_chr(), new_sid))
      v2_actual <- if (isTRUE(input$edit_visit2_done)) safe_date(input$edit_visit2_done_date %||% input$edit_visit2_date) else NA
      dbExecute(con, "UPDATE visits SET planned_date = ?, status = ?, actual_date = ?, updated_at = ? WHERE subject_id = ? AND event_name = 'ON-1 PRE/POST (Visit 2) | First stimulation session'",
                params = list(as.character(safe_date(input$edit_visit2_date)), ifelse(isTRUE(input$edit_visit2_done), "done", "planned"), ifelse(is.na(v2_actual), NA, as.character(v2_actual)), now_chr(), new_sid))
      if (isTRUE(input$edit_recalc_subject)) recalc_subject_schedule(new_sid)
      apply_subject_dropout(new_sid, isTRUE(input$edit_is_dropout), if (isTRUE(input$edit_is_dropout)) input$edit_dropout_date else NA, if (isTRUE(input$edit_is_dropout)) input$edit_dropout_reason else NA)
      editing_subject_id(NULL); removeModal(); refresh_all(); current_subject_id(new_sid)
      showNotification("Participant details updated.", type = "message")
    }, error = function(e) showNotification(e$message, type = "error"))
  })

  # ---- Edit task ----
  get_selected_visit_id <- function(df_now, idx) {
    if (length(idx) != 1) return(NA_integer_)
    if (is.null(df_now) || nrow(df_now) < idx) return(NA_integer_)
    out <- suppressWarnings(as.integer(df_now$id[[idx]]))
    if (length(out) != 1 || is.na(out)) return(NA_integer_)
    out
  }

  open_event_modal <- function(selected_id) {
    selected_id <- suppressWarnings(as.integer(selected_id))
    if (length(selected_id) != 1 || is.na(selected_id)) {
      showNotification("Could not locate that task. Select a row again and retry.", type = "error")
      return(invisible(FALSE))
    }

    visit_row <- visits_rv() %>% filter(id == selected_id)
    if (nrow(visit_row) != 1) {
      showNotification("Task record not found. Refresh and try again.", type = "error")
      return(invisible(FALSE))
    }

    editing_event_id(selected_id)
    showModal(modalDialog(
      title = paste0("Edit task: ", visit_row$subject_id, " | ", visit_row$event_name),
      textInput("edit_event_name", "Event name", value = na2empty(visit_row$event_name)),
      selectInput("edit_event_type", "Event type", choices = c("onsite", "home", "online", "followup", "pause"), selected = scalar_chr(visit_row$event_type, "onsite")),
      numericInput("edit_session_no", "Session number (home sessions only)", value = ifelse(is.na(visit_row$session_no[1]), NA, visit_row$session_no[1]), min = 1, step = 1),
      dateInput("edit_planned_date", "Planned date", value = visit_row$planned_date[1]),
      selectInput("edit_status", "Status", choices = status_levels, selected = scalar_chr(visit_row$status, "planned")),
      dateInput("edit_actual_date", "Actual date", value = ifelse(is.na(visit_row$actual_date[1]), visit_row$planned_date[1], visit_row$actual_date[1])),
      textAreaInput("edit_note", "Notes", value = na2empty(visit_row$note[1]), rows = 3),
      textInput("edit_operator", "Recorded by / operator", value = na2empty(visit_row$completed_by[1] %||% input$operator_name)),
      checkboxInput("edit_recalc_after", "Downstream dates shift automatically after saving", TRUE),
      footer = tagList(modalButton("Cancel"), actionButton("save_event_edit", "Save task changes", class = "btn-primary"))
    ))
    invisible(TRUE)
  }

  observeEvent(input$edit_subject_task_btn, {
    idx <- input$subject_detail_table_rows_selected
    selected_id <- get_selected_visit_id(subject_detail_df(), idx)

    if (length(selected_id) != 1 || is.na(selected_id)) {
      # Fallback: rows_selected is sometimes not written to the reactiveVal promptly
      # after switching participant, so read it again from the visible table data
      if (!is.null(idx) && length(idx) == 1) {
        selected_id <- suppressWarnings(as.integer(subject_detail_df()$id[idx]))
      }
    }

    if (length(selected_id) != 1 || is.na(selected_id)) {
      return(showNotification("Select a row in 'All tasks for this participant' first.", type = "warning"))
    }

    selected_subject_task_id(selected_id)
    open_event_modal(selected_id)
  })

  observeEvent(input$edit_daily_task_btn, {
    idx <- input$daily_tasks_rows_selected
    selected_id <- get_selected_visit_id(daily_tasks_df(), idx)

    if (length(selected_id) != 1 || is.na(selected_id)) {
      if (!is.null(idx) && length(idx) == 1) {
        selected_id <- suppressWarnings(as.integer(daily_tasks_df()$id[idx]))
      }
    }

    if (length(selected_id) != 1 || is.na(selected_id)) {
      return(showNotification("Select a row in 'Today's tasks' first.", type = "warning"))
    }

    selected_daily_task_id(selected_id)
    open_event_modal(selected_id)
  })
  observeEvent(input$save_event_edit, {
    eid <- editing_event_id(); req(eid)
    tryCatch({
      update_visit_full(eid, input$edit_event_name, input$edit_event_type, input$edit_session_no, input$edit_planned_date,
                        input$edit_status, input$edit_actual_date, input$edit_note, input$edit_operator,
                        recalc_after = isTRUE(input$edit_recalc_after))
      editing_event_id(NULL); removeModal(); refresh_all()
      selected_subject_task_id(NA_integer_)
      selected_daily_task_id(NA_integer_)
      dataTableProxy("subject_detail_table") %>% selectRows(NULL)
      dataTableProxy("daily_tasks") %>% selectRows(NULL)
      dataTableProxy("overdue_tasks") %>% selectRows(NULL)
      showNotification("Task updated.", type = "message")
    }, error = function(e) showNotification(e$message, type = "error"))
  })

  observeEvent(input$delete_subject_btn, {
    req(input$delete_subject_id)
    tryCatch({ delete_subject(input$delete_subject_id); refresh_all(); showNotification("Participant deleted.", type = "message") }, error = function(e) showNotification(e$message, type = "error"))
  })

  # ---- Home diary ----
  diary_records <- reactiveVal(NULL)
  latest_summary <- reactiveVal(NULL)
  mood_label_map <- c("1" = "Very poor", "2" = "Poor", "3" = "Fair", "4" = "Good", "5" = "Very good")
  awake_slots <- reactive(generate_awake_slots(input$awake_start, input$awake_end, input$nap_start, input$nap_end))
  output$slot_inputs <- renderUI({
    slots <- awake_slots(); if (length(slots) == 0) return(tags$p("Choose a valid awake period: the end time must be later than the start time."))
    tagList(lapply(seq_along(slots), function(i) {
      div(class = "slot-box", radioButtons(inputId = paste0("slot_", i), label = slots[i],
                                          choices = c("1 Good control" = 1, "2 Non-troublesome dyskinesia" = 2, "3 Troublesome dyskinesia" = 3, "4 Severe discomfort" = 4), selected = 1, inline = TRUE))
    }))
  })
  observeEvent(input$diary_reset_all, {
    updateDateInput(session, "diary_date", value = Sys.Date()); updateSliderInput(session, "self_score", value = 80)
    updateRadioButtons(session, "fall_event", selected = 0); updateTextAreaInput(session, "fall_detail", value = "")
    updateRadioButtons(session, "adverse_event", selected = 0); updateTextAreaInput(session, "adverse_detail", value = "")
    updateRadioButtons(session, "mood_score", selected = 3); updateRadioButtons(session, "exercise", selected = 0)
    updateTextAreaInput(session, "exercise_detail", value = ""); updateRadioButtons(session, "constipation", selected = 0)
    updateTextInput(session, "med_onset_mins", value = ""); updateSelectInput(session, "night_sleep_start", selected = "21:00")
    updateSelectInput(session, "night_sleep_end", selected = "06:00"); updateSelectInput(session, "nap_start", selected = "")
    updateSelectInput(session, "nap_end", selected = ""); updateSelectInput(session, "awake_start", selected = "06:00")
    updateSelectInput(session, "awake_end", selected = "22:00"); updateTextAreaInput(session, "other_symptoms", value = "")
    showNotification("Reset.", type = "message")
  })
  summary_data <- reactive({
    slots <- awake_slots(); if (length(slots) == 0) return(NULL)
    state_values <- sapply(seq_along(slots), function(i) { val <- input[[paste0("slot_", i)]]; if (is.null(val) || val == "") return(1); as.numeric(val) })
    night_sleep_hours <- calc_duration_hours(input$night_sleep_start, input$night_sleep_end)
    nap_hours <- if (input$nap_start == "" || input$nap_end == "") 0 else calc_duration_hours(input$nap_start, input$nap_end)
    tibble(
      n_slots = length(slots),
      good_control_n = sum(state_values == 1, na.rm = TRUE),
      non_troublesome_dyskinesia_n = sum(state_values == 2, na.rm = TRUE),
      troublesome_dyskinesia_n = sum(state_values == 3, na.rm = TRUE),
      very_uncomfortable_n = sum(state_values == 4, na.rm = TRUE),
      good_control_pct = round(mean(state_values == 1) * 100, 1),
      dyskinesia_pct = round(mean(state_values %in% c(2, 3)) * 100, 1),
      poor_state_pct = round(mean(state_values %in% c(3, 4)) * 100, 1),
      night_sleep_hours = night_sleep_hours,
      nap_hours = nap_hours,
      total_sleep_hours = round(night_sleep_hours + nap_hours, 2)
    )
  })
  output$summary_table <- renderTable(summary_data(), striped = TRUE, bordered = TRUE)
  observeEvent(input$diary_calc_save, {
    req(input$diary_pid)
    slots <- awake_slots(); req(length(slots) > 0)
    state_values <- sapply(seq_along(slots), function(i) { val <- input[[paste0("slot_", i)]]; if (is.null(val) || val == "") return(1); as.numeric(val) })
    slot_wide <- tibble(slot = slots, state = state_values) %>% pivot_wider(names_from = slot, values_from = state, names_prefix = "slot_")
    meta <- tibble(pid = input$diary_pid, diary_date = as.character(input$diary_date), self_score = as.numeric(input$self_score),
                   fall_event = as.numeric(input$fall_event), fall_detail = input$fall_detail,
                   adverse_event = as.numeric(input$adverse_event), adverse_detail = input$adverse_detail,
                   mood_score = as.numeric(input$mood_score), mood_label = unname(mood_label_map[as.character(input$mood_score)]),
                   exercise = as.numeric(input$exercise), exercise_detail = input$exercise_detail,
                   constipation = as.numeric(input$constipation), med_onset_mins = input$med_onset_mins,
                   night_sleep_start = input$night_sleep_start, night_sleep_end = input$night_sleep_end,
                   nap_start = input$nap_start, nap_end = input$nap_end, awake_start = input$awake_start, awake_end = input$awake_end,
                   other_symptoms = input$other_symptoms, saved_at_local = now_chr())
    out <- bind_cols(meta, summary_data(), slot_wide)
    db_ok <- tryCatch({
      append_db_wide_record("family_diary_records", out)
      TRUE
    }, error = function(e) {
      showNotification(paste0("Database write failed: ", e$message), type = "error", duration = 8)
      FALSE
    })
    if (!db_ok) return(NULL)

    individual_path <- tryCatch(
      save_individual_record_csv(out, input$diary_pid, input$diary_date, diary_individual_dir, create_pid_dir = TRUE),
      error = function(e) NA_character_
    )
    master_ok <- tryCatch({ append_record_csv(diary_record_file, out); TRUE }, error = function(e) { showNotification(paste0("Saved to the database, but the local master-file backup failed: ", e$message), type = "warning"); FALSE })
    diary_records(read_diary_records_db(input$diary_pid))
    latest_summary(summary_data())
    backup_msg <- if (!is.na(individual_path)) normalizePath(individual_path, winslash = "/", mustWork = FALSE) else "Local individual backup failed"
    showNotification(paste0("Saved to the database; local backup: ", backup_msg), type = "message")
  })
  observeEvent(input$diary_load_hist, {
    diary_records(read_diary_records_db(input$diary_pid))
    showNotification("Home diary history reloaded from the database.", type = "message")
  })
  output$hist_plot <- renderPlot({
    input$diary_load_hist
    req(input$diary_pid, input$diary_hist_metric)
    df <- diary_records()
    aliases <- id_aliases_for_module(input$diary_pid, module = "default")
    if (!is.null(df) && nrow(df) > 0 && "pid" %in% names(df)) df <- df %>% filter(pid %in% aliases)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    out <- order_history_df(df)

    out <- out %>%
      mutate(
        good_control_n = suppressWarnings(as.numeric(good_control_n)),
        non_troublesome_dyskinesia_n = suppressWarnings(as.numeric(non_troublesome_dyskinesia_n)),
        troublesome_dyskinesia_n = suppressWarnings(as.numeric(troublesome_dyskinesia_n)),
        very_uncomfortable_n = suppressWarnings(as.numeric(very_uncomfortable_n))
      ) %>%
      mutate(
        awake_slot_n = good_control_n + non_troublesome_dyskinesia_n + troublesome_dyskinesia_n + very_uncomfortable_n,
        good_control_pct = ifelse(!is.na(awake_slot_n) & awake_slot_n > 0, good_control_n / awake_slot_n * 100, NA_real_),
        non_troublesome_dyskinesia_pct = ifelse(!is.na(awake_slot_n) & awake_slot_n > 0, non_troublesome_dyskinesia_n / awake_slot_n * 100, NA_real_),
        troublesome_dyskinesia_pct = ifelse(!is.na(awake_slot_n) & awake_slot_n > 0, troublesome_dyskinesia_n / awake_slot_n * 100, NA_real_),
        very_uncomfortable_pct = ifelse(!is.na(awake_slot_n) & awake_slot_n > 0, very_uncomfortable_n / awake_slot_n * 100, NA_real_)
      )

    metric_labels <- c(
      self_score = "Self-rating",
      good_control_pct = "Good control (% of awake time)",
      non_troublesome_dyskinesia_pct = "Non-troublesome dyskinesia (% of awake time)",
      troublesome_dyskinesia_pct = "Troublesome dyskinesia (% of awake time)",
      very_uncomfortable_pct = "Severe discomfort (% of awake time)"
    )

    y_lab <- if (identical(input$diary_hist_metric, "self_score")) "Self Score" else "%"

    history_line_plot(
      out, "diary_date", input$diary_hist_metric,
      paste0("Diary History: ", metric_labels[[input$diary_hist_metric]]),
      y_lab
    )
  })

# ---- records cache ----
  gai_records <- reactiveVal(NULL)
  gds_records <- reactiveVal(NULL)
  ess_records <- reactiveVal(NULL)
  pdss2_records <- reactiveVal(NULL)
  scopa_records <- reactiveVal(NULL)
  updrs1_records <- reactiveVal(NULL)
  updrs2_records <- reactiveVal(NULL)
  nmss_records <- reactiveVal(NULL)
  pdq39_records <- reactiveVal(NULL)
  isi_records <- reactiveVal(NULL)
  
  register_downloads(output, "gai", gai_record_file, "GAI")
  register_downloads(output, "gds", gds_record_file, "GDS")
  register_downloads(output, "ess", ess_record_file, "ESS")
  register_downloads(output, "pdss2", pdss2_record_file, "PDSS2")
  register_downloads(output, "scopa", scopa_record_file, "SCOPA_AUT")
  register_downloads(output, "updrs1", updrs1_record_file, "UPDRS_I")
  register_downloads(output, "updrs2", updrs2_record_file, "UPDRS_II")
  register_downloads(output, "nmss", nmss_record_file, "NMSS")
  register_downloads(output, "pdq39", pdq39_record_file, "PDQ39")
  register_downloads(output, "isi", isi_record_file, "ISI")
  

  # ---- Non-motor symptom history plot (built from the saved records; display only) ----
  output$gai_hist_plot <- renderPlot({
    req(input$gai_pid, input$gai_hist_metric)
    df <- filter_records_for_pid(gai_records(), input$gai_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$gai_hist_metric, paste0("GAI: ", input$gai_hist_metric), input$gai_hist_metric)
  })
  output$gds_hist_plot <- renderPlot({
    req(input$gds_pid, input$gds_hist_metric)
    df <- filter_records_for_pid(gds_records(), input$gds_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$gds_hist_metric, paste0("GDS: ", input$gds_hist_metric), input$gds_hist_metric)
  })
  output$ess_hist_plot <- renderPlot({
    req(input$ess_pid, input$ess_hist_metric)
    df <- filter_records_for_pid(ess_records(), input$ess_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$ess_hist_metric, paste0("ESS: ", input$ess_hist_metric), input$ess_hist_metric)
  })
  output$pdss2_hist_plot <- renderPlot({
    req(input$pdss2_pid, input$pdss2_hist_metric)
    df <- filter_records_for_pid(pdss2_records(), input$pdss2_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$pdss2_hist_metric, paste0("PDSS-2: ", input$pdss2_hist_metric), input$pdss2_hist_metric)
  })
  output$scopa_hist_plot <- renderPlot({
    req(input$scopa_pid, input$scopa_hist_metric)
    df <- filter_records_for_pid(scopa_records(), input$scopa_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$scopa_hist_metric, paste0("SCOPA-AUT: ", input$scopa_hist_metric), input$scopa_hist_metric)
  })
  output$updrs1_hist_plot <- renderPlot({
    req(input$updrs1_pid, input$updrs1_hist_metric)
    df <- filter_records_for_pid(updrs1_records(), input$updrs1_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$updrs1_hist_metric, paste0("MDS-UPDRS I: ", input$updrs1_hist_metric), input$updrs1_hist_metric)
  })
  output$updrs2_hist_plot <- renderPlot({
    req(input$updrs2_pid, input$updrs2_hist_metric)
    df <- filter_records_for_pid(updrs2_records(), input$updrs2_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$updrs2_hist_metric, paste0("MDS-UPDRS II: ", input$updrs2_hist_metric), input$updrs2_hist_metric)
  })
  output$nmss_hist_plot <- renderPlot({
    req(input$nmss_pid, input$nmss_hist_metric)
    df <- filter_records_for_pid(nmss_records(), input$nmss_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$nmss_hist_metric, paste0("NMSS: ", input$nmss_hist_metric), input$nmss_hist_metric)
  })
  output$pdq39_hist_plot <- renderPlot({
    req(input$pdq39_pid, input$pdq39_hist_metric)
    df <- filter_records_for_pid(pdq39_records(), input$pdq39_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$pdq39_hist_metric, paste0("PDQ-39: ", input$pdq39_hist_metric), input$pdq39_hist_metric)
  })
  output$isi_hist_plot <- renderPlot({
    req(input$isi_pid, input$isi_hist_metric)
    df <- filter_records_for_pid(isi_records(), input$isi_pid)
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    history_line_plot(order_history_df(df), "date", input$isi_hist_metric, paste0("ISI: ", input$isi_hist_metric), input$isi_hist_metric)
  })

  # =========================
  # GAI
  # =========================
  output$gai_questions <- renderUI({
    tagList(lapply(seq_along(GAI_ITEMS), function(i) {
      radio_question(paste0("gai_q", i), paste0(i, ". ", GAI_ITEMS[i]), c("Yes" = 1, "No" = 0))
    }))
  })
  
  gai_summary_r <- reactive({
    vals <- get_scale_vals(input, "gai_q", length(GAI_ITEMS))
    if (!all_answered(vals)) return(NULL)
    tibble(Total_GAI = sum(vals), Positive_n = sum(vals == 1))
  })
  
  output$gai_summary <- renderTable(gai_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$gai_load, {
    gai_records(read_scale_records_db("GAI", gai_record_file, input$gai_pid))
  })
  
  output$gai_hist <- renderTable(
    render_hist_table(gai_records(), input$gai_pid, c("pid","date","visit","Total_GAI","Positive_n","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$gai_reset, {
    updateSelectInput(session, "gai_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "gai_date", value = Sys.Date())
    updateTextInput(session, "gai_visit", value = "Baseline")
    for (i in seq_along(GAI_ITEMS)) updateRadioButtons(session, paste0("gai_q", i), selected = character(0))
    showNotification("GAI reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$gai_save, {
    vals_num <- get_scale_vals(input, "gai_q", length(GAI_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("GAI still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    meta <- tibble(pid = input$gai_pid, date = as.character(input$gai_date), visit = input$gai_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("GAI_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_GAI = sum(vals_num), Positive_n = sum(vals_num == 1)), wide)
    save_scale_record(out, input$gai_pid, input$gai_date, gai_individual_dir, gai_record_file, gai_records, session, "GAI saved", "GAI")
  })
  
  # =========================
  # GDS
  # =========================
  output$gds_questions <- renderUI({
    tagList(lapply(seq_along(GDS_ITEMS), function(i) {
      radio_question(paste0("gds_q", i), paste0(i, ". ", GDS_ITEMS[i]), c("Yes" = 1, "No" = 0))
    }))
  })
  
  gds_summary_r <- reactive({
    vals <- get_scale_vals(input, "gds_q", length(GDS_ITEMS))
    if (!all_answered(vals)) return(NULL)
    score <- vals
    score[GDS_REVERSE_YES0] <- 1 - score[GDS_REVERSE_YES0]
    tibble(Total_GDS = sum(score))
  })
  
  output$gds_summary <- renderTable(gds_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$gds_load, {
    gds_records(read_scale_records_db("GDS", gds_record_file, input$gds_pid))
  })
  
  output$gds_hist <- renderTable(
    render_hist_table(gds_records(), input$gds_pid, c("pid","date","visit","Total_GDS","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$gds_reset, {
    updateSelectInput(session, "gds_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "gds_date", value = Sys.Date())
    updateTextInput(session, "gds_visit", value = "Baseline")
    for (i in seq_along(GDS_ITEMS)) updateRadioButtons(session, paste0("gds_q", i), selected = character(0))
    showNotification("GDS reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$gds_save, {
    vals_num <- get_scale_vals(input, "gds_q", length(GDS_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("GDS still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    score <- vals_num
    score[GDS_REVERSE_YES0] <- 1 - score[GDS_REVERSE_YES0]
    meta <- tibble(pid = input$gds_pid, date = as.character(input$gds_date), visit = input$gds_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("GDS_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_GDS = sum(score)), wide)
    save_scale_record(out, input$gds_pid, input$gds_date, gds_individual_dir, gds_record_file, gds_records, session, "GDS saved", "GDS")
  })
  
  # =========================
  # ESS
  # =========================
  output$ess_questions <- renderUI({
    tagList(lapply(seq_along(ESS_ITEMS), function(i) {
      radio_question(paste0("ess_q", i), paste0(i, ". ", ESS_ITEMS[i]), c("0" = 0, "1" = 1, "2" = 2, "3" = 3))
    }))
  })
  
  ess_summary_r <- reactive({
    vals <- get_scale_vals(input, "ess_q", length(ESS_ITEMS))
    if (!all_answered(vals)) return(NULL)
    tibble(Total_ESS = sum(vals))
  })
  
  output$ess_summary <- renderTable(ess_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$ess_load, {
    ess_records(read_scale_records_db("ESS", ess_record_file, input$ess_pid))
  })
  
  output$ess_hist <- renderTable(
    render_hist_table(ess_records(), input$ess_pid, c("pid","date","visit","Total_ESS","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$ess_reset, {
    updateSelectInput(session, "ess_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "ess_date", value = Sys.Date())
    updateTextInput(session, "ess_visit", value = "Baseline")
    for (i in seq_along(ESS_ITEMS)) updateRadioButtons(session, paste0("ess_q", i), selected = character(0))
    showNotification("ESS reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$ess_save, {
    vals_num <- get_scale_vals(input, "ess_q", length(ESS_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("ESS still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    meta <- tibble(pid = input$ess_pid, date = as.character(input$ess_date), visit = input$ess_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("ESS_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_ESS = sum(vals_num)), wide)
    save_scale_record(out, input$ess_pid, input$ess_date, ess_individual_dir, ess_record_file, ess_records, session, "ESS saved", "ESS")
  })
  
  # =========================
  # PDSS-2
  # =========================
  output$pdss2_questions <- renderUI({
    tagList(lapply(seq_along(PDSS2_ITEMS), function(i) {
      radio_question(paste0("pdss2_q", i), paste0(i, ". ", PDSS2_ITEMS[i]), c("0" = 0, "1" = 1, "2" = 2, "3" = 3, "4" = 4))
    }))
  })
  
  pdss2_summary_r <- reactive({
    vals <- get_scale_vals(input, "pdss2_q", length(PDSS2_ITEMS))
    if (!all_answered(vals)) return(NULL)
    score <- vals
    score[1] <- 4 - score[1]
    tibble(Total_PDSS2 = sum(score))
  })
  
  output$pdss2_summary <- renderTable(pdss2_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$pdss2_load, {
    pdss2_records(read_scale_records_db("PDSS2", pdss2_record_file, input$pdss2_pid))
  })
  
  output$pdss2_hist <- renderTable(
    render_hist_table(pdss2_records(), input$pdss2_pid, c("pid","date","visit","Total_PDSS2","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$pdss2_reset, {
    updateSelectInput(session, "pdss2_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "pdss2_date", value = Sys.Date())
    updateTextInput(session, "pdss2_visit", value = "Baseline")
    for (i in seq_along(PDSS2_ITEMS)) updateRadioButtons(session, paste0("pdss2_q", i), selected = character(0))
    showNotification("PDSS-2 reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$pdss2_save, {
    vals_num <- get_scale_vals(input, "pdss2_q", length(PDSS2_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("PDSS-2 still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    score <- vals_num
    score[1] <- 4 - score[1]
    meta <- tibble(pid = input$pdss2_pid, date = as.character(input$pdss2_date), visit = input$pdss2_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("PDSS2_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_PDSS2 = sum(score)), wide)
    save_scale_record(out, input$pdss2_pid, input$pdss2_date, pdss2_individual_dir, pdss2_record_file, pdss2_records, session, "PDSS-2 saved", "PDSS2")
  })
  
  # =========================
  # SCOPA-AUT
  # =========================
  output$scopa_questions <- renderUI({
    tagList(lapply(seq_along(SCOPA_ITEMS), function(i) {
      show_item <- TRUE
      if (i %in% c(22, 23) && !is.null(input$scopa_sex) && input$scopa_sex == "Female") show_item <- FALSE
      if (i %in% c(24, 25) && !is.null(input$scopa_sex) && input$scopa_sex == "Male") show_item <- FALSE
      if (!show_item) return(NULL)
      
      tagList(
        if (i == 1 || SCOPA_GROUP[i] != SCOPA_GROUP[i - 1]) tags$h4(SCOPA_GROUP[i]),
        radio_question(paste0("scopa_q", i), paste0(i, ". ", SCOPA_ITEMS[i]),
                       c("0 None / occasional" = 0, "1 Sometimes" = 1, "2 Often" = 2, "3 Persistent" = 3))
      )
    }))
  })
  
  scopa_summary_r <- reactive({
    sex <- input$scopa_sex
    active_idx <- seq_along(SCOPA_ITEMS)
    if (sex == "Male") active_idx <- setdiff(active_idx, c(24, 25))
    if (sex == "Female") active_idx <- setdiff(active_idx, c(22, 23))
    
    vals <- vapply(active_idx, function(i) {
      safe_input_num(input[[paste0("scopa_q", i)]])
    }, numeric(1))
    
    if (length(vals) == 0 || !all_answered(vals)) return(NULL)
    tibble(Total_SCOPA_AUT = sum(vals), N_items = length(vals))
  })
  
  output$scopa_summary <- renderTable(scopa_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$scopa_load, {
    scopa_records(read_scale_records_db("SCOPA_AUT", scopa_record_file, input$scopa_pid))
  })
  
  output$scopa_hist <- renderTable(
    render_hist_table(scopa_records(), input$scopa_pid, c("pid","date","visit","sex","Total_SCOPA_AUT","N_items","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$scopa_reset, {
    updateSelectInput(session, "scopa_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "scopa_date", value = Sys.Date())
    updateTextInput(session, "scopa_visit", value = "Baseline")
    updateSelectInput(session, "scopa_sex", selected = "")
    for (i in seq_along(SCOPA_ITEMS)) updateRadioButtons(session, paste0("scopa_q", i), selected = character(0))
    showNotification("SCOPA-AUT reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$scopa_save, {
    sex <- input$scopa_sex
    active_idx <- seq_along(SCOPA_ITEMS)
    if (sex == "Male") active_idx <- setdiff(active_idx, c(24, 25))
    if (sex == "Female") active_idx <- setdiff(active_idx, c(22, 23))
    
    vals_active <- vapply(active_idx, function(i) {
      safe_input_num(input[[paste0("scopa_q", i)]])
    }, numeric(1))
    
    if (!all_answered(vals_active)) {
      showNotification("SCOPA-AUT still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    
    vals_num <- vapply(seq_along(SCOPA_ITEMS), function(i) {
      safe_input_num(input[[paste0("scopa_q", i)]])
    }, numeric(1))
    
    meta <- tibble(pid = input$scopa_pid, date = as.character(input$scopa_date), visit = input$scopa_visit,
                   sex = input$scopa_sex, saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("SCOPA_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_SCOPA_AUT = sum(vals_active), N_items = length(active_idx)), wide)
    save_scale_record(out, input$scopa_pid, input$scopa_date, scopa_individual_dir, scopa_record_file, scopa_records, session, "SCOPA-AUT saved", "SCOPA_AUT")
  })
  
  # =========================
  # UPDRS I
  # =========================
  output$updrs1_questions <- renderUI({
    tagList(lapply(seq_along(UPDRS1_ITEMS), function(i) {
      radio_question(paste0("updrs1_q", i), paste0(i, ". ", UPDRS1_ITEMS[i]), c("0" = 0, "1" = 1, "2" = 2, "3" = 3, "4" = 4))
    }))
  })
  
  updrs1_summary_r <- reactive({
    vals <- get_scale_vals(input, "updrs1_q", length(UPDRS1_ITEMS))
    if (!all_answered(vals)) return(NULL)
    tibble(Total_UPDRS_I = sum(vals))
  })
  
  output$updrs1_summary <- renderTable(updrs1_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$updrs1_load, {
    updrs1_records(read_scale_records_db("UPDRS_I", updrs1_record_file, input$updrs1_pid))
  })
  
  output$updrs1_hist <- renderTable(
    render_hist_table(updrs1_records(), input$updrs1_pid, c("pid","date","visit","Total_UPDRS_I","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$updrs1_reset, {
    updateSelectInput(session, "updrs1_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "updrs1_date", value = Sys.Date())
    updateTextInput(session, "updrs1_visit", value = "Baseline")
    for (i in seq_along(UPDRS1_ITEMS)) updateRadioButtons(session, paste0("updrs1_q", i), selected = character(0))
    showNotification("MDS-UPDRS I reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$updrs1_save, {
    vals_num <- get_scale_vals(input, "updrs1_q", length(UPDRS1_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("MDS-UPDRS I still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    meta <- tibble(pid = input$updrs1_pid, date = as.character(input$updrs1_date), visit = input$updrs1_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("UPDRS1_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_UPDRS_I = sum(vals_num)), wide)
    save_scale_record(out, input$updrs1_pid, input$updrs1_date, updrs1_individual_dir, updrs1_record_file, updrs1_records, session, "MDS-UPDRS I saved", "UPDRS_I")
  })
  
  # =========================
  # UPDRS II
  # =========================
  output$updrs2_questions <- renderUI({
    tagList(lapply(seq_along(UPDRS2_ITEMS), function(i) {
      radio_question(paste0("updrs2_q", i), paste0(i, ". ", UPDRS2_ITEMS[i]), c("0" = 0, "1" = 1, "2" = 2, "3" = 3, "4" = 4))
    }))
  })
  
  updrs2_summary_r <- reactive({
    vals <- get_scale_vals(input, "updrs2_q", length(UPDRS2_ITEMS))
    if (!all_answered(vals)) return(NULL)
    tibble(Total_UPDRS_II = sum(vals))
  })
  
  output$updrs2_summary <- renderTable(updrs2_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$updrs2_load, {
    updrs2_records(read_scale_records_db("UPDRS_II", updrs2_record_file, input$updrs2_pid))
  })
  
  output$updrs2_hist <- renderTable(
    render_hist_table(updrs2_records(), input$updrs2_pid, c("pid","date","visit","Total_UPDRS_II","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$updrs2_reset, {
    updateSelectInput(session, "updrs2_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "updrs2_date", value = Sys.Date())
    updateTextInput(session, "updrs2_visit", value = "Baseline")
    for (i in seq_along(UPDRS2_ITEMS)) updateRadioButtons(session, paste0("updrs2_q", i), selected = character(0))
    showNotification("MDS-UPDRS II reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$updrs2_save, {
    vals_num <- get_scale_vals(input, "updrs2_q", length(UPDRS2_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("MDS-UPDRS II still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    meta <- tibble(pid = input$updrs2_pid, date = as.character(input$updrs2_date), visit = input$updrs2_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("UPDRS2_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_UPDRS_II = sum(vals_num)), wide)
    save_scale_record(out, input$updrs2_pid, input$updrs2_date, updrs2_individual_dir, updrs2_record_file, updrs2_records, session, "MDS-UPDRS II saved", "UPDRS_II")
  })
  
  # =========================
  # NMSS
  # =========================
  output$nmss_questions <- renderUI({
    tagList(lapply(seq_along(NMSS_ITEMS), function(i) {
      yes_id <- paste0("nmss_yes_", i)
      sev_id <- paste0("nmss_sev_", i)
      freq_id <- paste0("nmss_freq_", i)
      
      tagList(
        tags$div(
          style = "border:1px solid #ddd; padding:10px; margin-bottom:10px; border-radius:8px;",
          radioButtons(yes_id, paste0(i, ". ", NMSS_ITEMS[i]),
                       choices = c("No" = 0, "Yes" = 1), selected = character(0), inline = TRUE),
          conditionalPanel(
            condition = sprintf("input['%s'] == '1'", yes_id),
            radioButtons(sev_id, "Severity",
                         choices = c("1 Mild" = 1, "2 Moderate" = 2, "3 Severe" = 3),
                         selected = character(0), inline = TRUE),
            radioButtons(freq_id, "Frequency",
                         choices = c("1 Rarely" = 1, "2 Often" = 2, "3 Frequent" = 3, "4 Very frequent" = 4),
                         selected = character(0), inline = TRUE)
          )
        )
      )
    }))
  })
  
  nmss_summary_r <- reactive({
    yes_vals <- vapply(seq_along(NMSS_ITEMS), function(i) {
      input[[paste0("nmss_yes_", i)]]
    }, character(1), USE.NAMES = FALSE)
    
    if (any(is.na(yes_vals) | yes_vals == "")) return(NULL)
    
    scores <- vapply(seq_along(NMSS_ITEMS), function(i) {
      y <- safe_input_num(input[[paste0("nmss_yes_", i)]])
      if (is.na(y)) return(NA_real_)
      if (y == 0) return(0)
      
      s <- safe_input_num(input[[paste0("nmss_sev_", i)]])
      f <- safe_input_num(input[[paste0("nmss_freq_", i)]])
      if (is.na(s) || is.na(f)) return(NA_real_)
      
      s * f
    }, numeric(1))
    
    if (any(is.na(scores))) return(NULL)
    
    tibble(Total_NMSS = sum(scores), Positive_items = sum(as.numeric(yes_vals) == 1))
  })
  
  output$nmss_summary <- renderTable(nmss_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$nmss_load, {
    nmss_records(read_scale_records_db("NMSS", nmss_record_file, input$nmss_pid))
  })
  
  output$nmss_hist <- renderTable(
    render_hist_table(nmss_records(), input$nmss_pid, c("pid","date","visit","Total_NMSS","Positive_items","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$nmss_reset, {
    updateSelectInput(session, "nmss_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "nmss_date", value = Sys.Date())
    updateTextInput(session, "nmss_visit", value = "Baseline")
    for (i in seq_along(NMSS_ITEMS)) {
      updateRadioButtons(session, paste0("nmss_yes_", i), selected = character(0))
      updateRadioButtons(session, paste0("nmss_sev_", i), selected = character(0))
      updateRadioButtons(session, paste0("nmss_freq_", i), selected = character(0))
    }
    showNotification("NMSS reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$nmss_save, {
    yes_num <- vapply(seq_along(NMSS_ITEMS), function(i) {
      safe_input_num(input[[paste0("nmss_yes_", i)]])
    }, numeric(1))
    
    if (!all_answered(yes_num)) {
      showNotification("NMSS still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    
    scores <- vapply(seq_along(NMSS_ITEMS), function(i) {
      y <- safe_input_num(input[[paste0("nmss_yes_", i)]])
      if (y == 0) return(0)
      
      s <- safe_input_num(input[[paste0("nmss_sev_", i)]])
      f <- safe_input_num(input[[paste0("nmss_freq_", i)]])
      if (is.na(s) || is.na(f)) return(NA_real_)
      
      s * f
    }, numeric(1))
    
    if (any(is.na(scores))) {
      showNotification("NMSS cannot be saved: one or more items answered 'yes' have no severity or frequency rating.", type = "error", duration = 6)
      return(NULL)
    }
    
    sev_num <- vapply(seq_along(NMSS_ITEMS), function(i) {
      safe_input_num(input[[paste0("nmss_sev_", i)]])
    }, numeric(1))
    freq_num <- vapply(seq_along(NMSS_ITEMS), function(i) {
      safe_input_num(input[[paste0("nmss_freq_", i)]])
    }, numeric(1))
    
    meta <- tibble(pid = input$nmss_pid, date = as.character(input$nmss_date), visit = input$nmss_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide_yes <- as_tibble(as.list(setNames(yes_num, paste0("NMSS_yes_", seq_along(yes_num)))))
    wide_sev <- as_tibble(as.list(setNames(sev_num, paste0("NMSS_sev_", seq_along(sev_num)))))
    wide_freq <- as_tibble(as.list(setNames(freq_num, paste0("NMSS_freq_", seq_along(freq_num)))))
    wide_score <- as_tibble(as.list(setNames(scores, paste0("NMSS_score_", seq_along(scores)))))
    out <- bind_cols(meta, tibble(Total_NMSS = sum(scores), Positive_items = sum(yes_num == 1)), wide_yes, wide_sev, wide_freq, wide_score)
    
    save_scale_record(out, input$nmss_pid, input$nmss_date, nmss_individual_dir, nmss_record_file, nmss_records, session, "NMSS saved", "NMSS")
  })
  
  # =========================
  # PDQ-39
  # =========================
  output$pdq39_questions <- renderUI({
    tagList(lapply(seq_along(PDQ39_ITEMS), function(i) {
      radio_question(paste0("pdq39_q", i), paste0(i, ". ", PDQ39_ITEMS[i]),
                     c("0 Never" = 0, "1 Occasionally" = 1, "2 Sometimes" = 2, "3 Often" = 3, "4 Always / unable to do it at all" = 4))
    }))
  })
  
  pdq39_summary_r <- reactive({
    vals <- get_scale_vals(input, "pdq39_q", length(PDQ39_ITEMS))
    if (!all_answered(vals)) return(NULL)
    tibble(Total_PDQ39 = sum(vals))
  })
  
  output$pdq39_summary <- renderTable(pdq39_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$pdq39_load, {
    pdq39_records(read_scale_records_db("PDQ39", pdq39_record_file, input$pdq39_pid))
  })
  
  output$pdq39_hist <- renderTable(
    render_hist_table(pdq39_records(), input$pdq39_pid, c("pid","date","visit","Total_PDQ39","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$pdq39_reset, {
    updateSelectInput(session, "pdq39_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "pdq39_date", value = Sys.Date())
    updateTextInput(session, "pdq39_visit", value = "Baseline")
    for (i in seq_along(PDQ39_ITEMS)) updateRadioButtons(session, paste0("pdq39_q", i), selected = character(0))
    showNotification("PDQ-39 reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$pdq39_save, {
    vals_num <- get_scale_vals(input, "pdq39_q", length(PDQ39_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("PDQ-39 still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    meta <- tibble(pid = input$pdq39_pid, date = as.character(input$pdq39_date), visit = input$pdq39_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("PDQ39_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_PDQ39 = sum(vals_num)), wide)
    save_scale_record(out, input$pdq39_pid, input$pdq39_date, pdq39_individual_dir, pdq39_record_file, pdq39_records, session, "PDQ-39 saved", "PDQ39")
  })
  
  # =========================
  # ISI
  # =========================
  output$isi_questions <- renderUI({
    tagList(lapply(seq_along(ISI_ITEMS), function(i) {
      radio_question(paste0("isi_q", i), paste0(i, ". ", ISI_ITEMS[i]), c("0" = 0, "1" = 1, "2" = 2, "3" = 3, "4" = 4))
    }))
  })
  
  isi_summary_r <- reactive({
    vals <- get_scale_vals(input, "isi_q", length(ISI_ITEMS))
    if (!all_answered(vals)) return(NULL)
    total_score <- sum(vals)
    tibble(Total_ISI = total_score, ISI_Level = ISI_INTERPRET(total_score))
  })
  
  output$isi_summary <- renderTable(isi_summary_r(), striped = TRUE, bordered = TRUE)
  
  observeEvent(input$isi_load, {
    isi_records(read_scale_records_db("ISI", isi_record_file, input$isi_pid))
  })
  
  output$isi_hist <- renderTable(
    render_hist_table(isi_records(), input$isi_pid, c("pid","date","visit","Total_ISI","ISI_Level","saved_at_local")),
    striped = TRUE, bordered = TRUE
  )
  
  observeEvent(input$isi_reset, {
    updateSelectInput(session, "isi_pid", selected = current_subject_id() %||% "S001")
    updateDateInput(session, "isi_date", value = Sys.Date())
    updateTextInput(session, "isi_visit", value = "Baseline")
    for (i in seq_along(ISI_ITEMS)) updateRadioButtons(session, paste0("isi_q", i), selected = character(0))
    showNotification("ISI reset.", type = "message", duration = 3)
  })
  
  observeEvent(input$isi_save, {
    vals_num <- get_scale_vals(input, "isi_q", length(ISI_ITEMS))
    if (!all_answered(vals_num)) {
      showNotification("ISI still has unanswered items and cannot be saved.", type = "error", duration = 6)
      return(NULL)
    }
    total_score <- sum(vals_num)
    meta <- tibble(pid = input$isi_pid, date = as.character(input$isi_date), visit = input$isi_visit,
                   saved_at_local = format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
    wide <- as_tibble(as.list(setNames(vals_num, paste0("ISI_", seq_along(vals_num)))))
    out <- bind_cols(meta, tibble(Total_ISI = total_score, ISI_Level = ISI_INTERPRET(total_score)), wide)
    save_scale_record(out, input$isi_pid, input$isi_date, isi_individual_dir, isi_record_file, isi_records, session, "ISI saved", "ISI")
  })

  # ---- U3 ----
  u3_records <- reactiveVal(NULL); latest_delta <- reactiveVal(NULL); latest_u3_summary <- reactiveVal(NULL)
  output$u3_question_inputs <- renderUI({
    tagList(lapply(seq_len(nrow(u3_items)), function(i) {
      div(
        class = "u3-score-group",
        radioButtons(
          inputId = paste0("I_", gsub("\\.", "_", u3_items$item[i])),
          label = paste0(u3_items$item[i], "  ", u3_items$label[i]),
          choices = c("0","1","2","3","4"),
          selected = character(0),
          inline = FALSE
        )
      )
    }))
  })
  observeEvent(input$u3_load_hist, {
    u3_records(u3_read_db_records(input$u3_pid))
    showNotification("UPDRS-III history reloaded from the database.", type = "message")
  })
  observeEvent(input$u3_reset_all, {
    for (it in u3_items$item) updateRadioButtons(session, paste0("I_", gsub("\\.", "_", it)), selected = character(0))
    updateRadioButtons(session, "dyskinesia_present", selected = "0")
    updateRadioButtons(session, "dyskinesia_interfere", selected = "0")
    showNotification("Reset: all scores cleared.", type = "message")
  })
  current_long <- reactive({
    vals <- sapply(u3_items$item, function(it) { val <- input[[paste0("I_", gsub("\\.", "_", it))]]; if (is.null(val) || val == "") return(NA_real_); as.numeric(val) })
    tibble(item = u3_items$item, label = u3_items$label, group = u3_items$group, value = as.numeric(vals))
  })
  observeEvent(input$u3_calc_save, {
    df_long <- current_long()
    if (any(is.na(df_long$value))) return(showNotification("Some items are still unscored and cannot be saved.", type = "error"))
    sumrow <- derive_scores(df_long)
    prev_records <- u3_records()
    if (is.null(prev_records) || nrow(prev_records) == 0) prev_records <- u3_read_db_records(input$u3_pid)
    prev_row <- get_previous_row(prev_records, input$u3_pid)
    latest_delta(calc_delta(sumrow, prev_row)); latest_u3_summary(sumrow)
    item_wide <- df_long %>% select(item, value) %>% pivot_wider(names_from = item, values_from = value)
    out <- bind_cols(tibble(pid = input$u3_pid, visit_date = as.character(input$u3_visit_date), visit = input$u3_visit, med_state = input$med_state, dbs_state = input$dbs_state, dyskinesia_present = input$dyskinesia_present, dyskinesia_interfere = input$dyskinesia_interfere, saved_at_local = now_chr()), sumrow, item_wide)
    tryCatch({
      u3_append_db_record(out)
    }, error = function(e) {
      stop(paste0("Database write failed: ", e$message))
    })
    u3_records(u3_read_db_records(input$u3_pid))
    individual_path <- u3_save_individual(out, input$u3_pid, input$u3_visit_date, u3_individual_dir)
    master_ok <- tryCatch({ u3_append_record(u3_record_file, out); TRUE }, error = function(e) { showNotification(paste0("Individual record saved, but the master file could not be updated: ", e$message), type = "warning"); FALSE })
    if (master_ok) u3_records(u3_read_db_records(input$u3_pid))
    showNotification(paste0("UPDRS-III saved to the database; local backup: ", normalizePath(individual_path, winslash = "/", mustWork = FALSE)), type = "message")
  })
  output$u3_score_table <- renderTable(latest_u3_summary(), striped = TRUE, bordered = TRUE)
  output$u3_side_table <- renderTable({ req(latest_u3_summary()); latest_u3_summary() %>% select(Left_sum, Right_sum, RminusL, Abs_LminusR, More_affected_side) }, striped = TRUE, bordered = TRUE)
  output$u3_delta_plot <- renderPlot({
    dd <- latest_delta()
    if (is.null(dd) || nrow(dd) == 0) return(blank_plot("No previous record, so the difference (current − previous) cannot be calculated."))
    keep <- dd %>% mutate(priority = metric %in% c("Total_III","Left_sum","Right_sum","Rigidity_sum","Tremor_sum")) %>% arrange(desc(priority), desc(abs(delta))) %>% head(12)
    ggplot(keep, aes(x = reorder(metric, delta), y = delta)) + geom_col() + coord_flip() + labs(x = NULL, y = "Current - Previous") + theme_minimal(base_size = 14)
  })
  output$u3_hist_plot <- renderPlot({
    input$u3_load_hist
    req(input$u3_pid, input$u3_hist_metric)
    df <- u3_records()
    aliases <- id_aliases_for_module(input$u3_pid, module = "u3")
    if (!is.null(df) && nrow(df) > 0 && "pid" %in% names(df)) {
      df <- df %>% filter(pid %in% aliases)
      if (all(c("visit_date", "saved_at_local") %in% names(df))) {
        df <- df %>% distinct(visit_date, saved_at_local, .keep_all = TRUE)
      }
    }
    if (is.null(df) || nrow(df) == 0) return(blank_plot("No history yet"))
    out <- order_history_df(df)
    metric_labels <- c(
      Total_III = "Total III",
      Left_sum = "Left Sum",
      Right_sum = "Right Sum",
      Rigidity_sum = "Rigidity Sum",
      Tremor_sum = "Tremor Sum"
    )
    history_line_plot(out, "visit_date", input$u3_hist_metric,
                      paste0("MDS-UPDRS III History: ", metric_labels[[input$u3_hist_metric]]),
                      metric_labels[[input$u3_hist_metric]])
  })

  # ---- Export ----
  output$download_excel <- downloadHandler(
    filename = function() paste0("TIs_PD_visit_manager_", format(Sys.Date(), "%Y%m%d"), ".xlsx"),
    content = function(file) {
      wb <- createWorkbook(); addWorksheet(wb, "subjects"); writeData(wb, "subjects", read_subjects()); addWorksheet(wb, "visits"); writeData(wb, "visits", read_visits()); saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  output$download_csv <- downloadHandler(
    filename = function() paste0("TIs_PD_visits_", format(Sys.Date(), "%Y%m%d"), ".csv"),
    content = function(file) write.csv(read_visits(), file, row.names = FALSE, fileEncoding = "UTF-8")
  )
}

shinyApp(ui, server)
