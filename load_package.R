#Script to install and load all the packages that are needed for the full model.

# List of packages to install
ForecastModel_Packages <- c(
  "odbc", "dplyr", "dbplyr", "stringr", "DBI",
  "RSQLite", "gamlss", "lubridate", 
  "fitdistrplus", "MASS", "bizdays", "RQuantLib", 
  "timeDate", "data.table", "tidyr", "foreach", 
  "doParallel", "taskscheduleR", "openxlsx"
)

# Function to install or load packages as necessary
load_packages <- function(package_name) {
  if (!requireNamespace(package_name, quietly = TRUE)) {
    install.packages(package_name, dependencies = TRUE)
  }
  library(package_name, character.only = TRUE)
}

# Loop through the list of packages and install/load them if needed
for (package in ForecastModel_Packages) {
  load_packages(package)
}

#Load Dates from RMetrics
load_rmetrics_calendars(c(2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025,2026,2027))
