# Entry point for `shiny::runApp()`.
#
# The whole application lives in a single file,
# TIS_PD_Clinical_Management_Suite_next_day_tasks.R, which ends with
# shinyApp(ui, server). Sourcing that file therefore returns the shiny app
# object, which this script passes back to runApp().
#
# Run from the repository root:
#   shiny::runApp()
# or:
#   shiny::runApp("path/to/public-release")

app <- source(
  "TIS_PD_Clinical_Management_Suite_next_day_tasks.R",
  local = TRUE,
  encoding = "UTF-8"
)$value

app
