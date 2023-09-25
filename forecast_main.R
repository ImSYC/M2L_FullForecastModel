#Setting Work Directory
setwd("F:/shiyuan/M2L_FullForecastModel/M2L_FullForecastModel/")
#Setting Output path
outputpath <- file.path("./R_Output", format(Sys.Date(), "%Y-%m"), floor_date(today(), unit= "week"))
dir.create(outputpath, recursive = TRUE)


#Calling the load packages script
source("load_package.R")
source("process_data.R")

################################################################################
