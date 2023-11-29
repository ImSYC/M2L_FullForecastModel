#Setting Work Directory
setwd("F:/shiyuan/M2L_FullForecastModel/M2L_FullForecastModel/")

#Start timer
STime <- Sys.time()

#Calling the load packages script
source("load_package.R")
source("process_data.R")



################################################################################
#1. Output Setup

#Setting Output path
outputpath <- file.path("./R_Output", format(Sys.Date(), "%Y-%m"), floor_date(today(), unit= "week"))
dir.create(outputpath, recursive = TRUE)

#Setup forecast model template
inputfile <-  file.path( "./R_Input/Template_Forecast_v8.xlsm")
wb <- loadWorkbook(inputfile)

#Set up run rate assumptions
# Get a list of CSV files in the folder
csv_files <- list.files("./R_Input/RunRateInputs", pattern = "\\.csv$", full.names = TRUE)

# Check if there are any CSV files in the folder
if (length(csv_files) == 0) {
  stop("No CSV files found in the folder.")
}

# Get file information for each CSV file
file_info <- lapply(csv_files, file.info)

# Find the latest CSV file based on the modification timestamp
latest_csv_index <- which.max(sapply(file_info, function(x) x$mtime))

# Get the path to the latest CSV file
latest_csv_path <- csv_files[latest_csv_index]
inputRunRates <- read.csv(file.path(latest_csv_path))

#Setting up output name
outputfilename <- paste(today(),"Forecast_v7.2_temp.xlsm", sep = "_")
outputfilepath <- file.path(outputpath, outputfilename)



################################################################################
#2. Analyse data

#Creating a function to fit the distribution of completion times by loan purpose, process, solicitor type, 
#Calculating and storing the accompanying cdf
Fit_Distribution <- function(data){
  
  # Initialize empty lists and define necessary data subsets
  MasterDistribution <- list()
  MasterCDF <- list()
  SubDistribution <- list()
  SubCDF <- list()
  Sol_Types <- unique(data$SolicitorType)
  Processes <- c("A2C", "A2F", "O2C", "O2F")
  Loan_purposes <- c("Purchase", "Drawdown", "Remortgage", "FurtherAdvance")
  
  # Loop through loan purposes
  for (Loan_purpose in Loan_purposes) {
    
    filtered_purpose_data <- data
    
    # Filter data based on loan purpose
    if (Loan_purpose == "Remortgage" | Loan_purpose == "FurtherAdvance") {
      filtered_purpose_data <- filtered_purpose_data[filtered_purpose_data$LoanPurpose %in% c("Remortgage", "FurtherAdvance"), ]
    } else {
      filtered_purpose_data <- filtered_purpose_data[filtered_purpose_data$LoanPurpose == Loan_purpose, ]
    }
    
    # Loop through processes
    for (Process in Processes){
      filtered_data <- filtered_purpose_data
      
      # Filter data based on process type
      if (Process == "A2C" | Process == "O2C"){
        filtered_data <- filtered_data %>% filter(DPRStatus %in% CompCon)
      } else if (Process == "A2F"){
        filtered_data <- filtered_data %>% filter(DPRStatus %in% FailCon)
      } else if (Process == "O2F"){
        filtered_data <- filtered_data %>% filter(DPRStatus %in% FailCon)
        filtered_data <- filtered_data[!is.na(filtered_data$OfferIssuedDate),]
      }
      
      # Fit a distribution to the filtered data
      fit <- fitDist(filtered_data[[Process]], k = 2, type = "realplus")
      MasterDistribution[[Loan_purpose]][[Process]] <- fit
      
      # Extract distribution parameters and calculate CDF (Cumulative Distribution Function)
      Parameters <- lapply(fit$parameters, function(x) fit[[x]])
      CDF <- lapply(T, function(x) do.call(paste0("p", fit$family[1]), c(x, Parameters)))
      MasterCDF[[Loan_purpose]][[Process]] <- CDF
      
    }
    
  }
  
  # Assign results to the global environment
  assign("MasterCDF", MasterCDF, envir = .GlobalEnv)
  assign("MasterDistribution", MasterDistribution, envir = .GlobalEnv)
  
  # Loop through solicitor types
  for (SolType in Sol_Types){
    
    filtered_soltype_data  <- data[data$SolicitorType == SolType, ]
    
    # Loop through loan purposes again
    for (Loan_purpose in Loan_purposes) {
      
      filtered_purpose_data <- filtered_soltype_data 
      
      # Filter data based on loan purpose
      if (Loan_purpose == "Remortgage" | Loan_purpose == "FurtherAdvance") {
        filtered_purpose_data <- filtered_purpose_data[filtered_purpose_data$LoanPurpose %in% c("Remortgage", "FurtherAdvance"), ]
      } else {
        filtered_purpose_data <- filtered_purpose_data[filtered_purpose_data$LoanPurpose == Loan_purpose, ]
      }
      
      # Loop through processes
      for (Process in Processes){
        filtered_data <- filtered_purpose_data
        
        # Filter data based on process type
        if (Process == "A2C" | Process == "O2C"){
          filtered_data <- filtered_data %>% filter(DPRStatus %in% CompCon)
        } else if (Process == "A2F"){
          filtered_data <- filtered_data %>% filter(DPRStatus %in% FailCon)
        } else if (Process == "O2F"){
          filtered_data <- filtered_data %>% filter(DPRStatus %in% FailCon)
          filtered_data <- filtered_data[!is.na(filtered_data$OfferIssuedDate),]
        }
        
        # Check the number of rows in filtered data
        if (nrow(filtered_data) > 500){
          
          # Fit a distribution to the filtered data
          fit <- fitDist(filtered_data[[Process]], k = 2, type = "realplus")
          SubDistribution[[Loan_purpose]][[Process]][[SolType]] <- fit
          
          # Extract distribution parameters and calculate CDF
          Parameters <- lapply(fit$parameters, function(x) fit[[x]])
          CDF <- lapply(T, function(x) do.call(paste0("p", fit$family[1]), c(x, Parameters)))
          SubCDF[[Loan_purpose]][[Process]][[SolType]] <- CDF
          
        } else {
          
          # If the number of rows is small, use the MasterDistribution and MasterCDF
          SubDistribution[[Loan_purpose]][[Process]][[SolType]] <- MasterDistribution[[Loan_purpose]][[Process]]
          SubCDF[[Loan_purpose]][[Process]][[SolType]] <- MasterCDF[[Loan_purpose]][[Process]]
          
        }
        
      }
      
    }
    
  }
  
  # Assign results to the global environment
  assign("SubCDF", SubCDF, envir = .GlobalEnv)
  assign("SubDistribution", SubDistribution, envir = .GlobalEnv)
  
}

################################################################################
#The dataframe All will refer to all applications that have occurred in the past 12 months.
#This will be the bases for our analysis.

All <- AllCases %>% filter(ApplicationDate >= (today() %m+% months(-12)))

#Fitting the distribution for all applications over the past 12 months
Fit_Distribution(All)

#Investigating the conversion rates
#Samples used

Sample_Exits          <- All %>% filter(DPRStatus %in% ExitCon)
Sample_Offers         <- Sample_Exits[!is.na(Sample_Exits$OfferIssuedDate), ]
Sample_A2C            <- Sample_Exits %>% filter(DPRStatus %in% CompCon)
Sample_A2F            <- Sample_Exits %>% filter(DPRStatus %in% FailCon)
Sample_O2F            <- Sample_A2F[!is.na(Sample_A2F$OfferIssuedDate),]

#Total Loan Value by funder under each sample
TotalLoanValue_Exits  <- aggregate(LoanAmount~Division, data = Sample_Exits, sum)
TotalLoanValue_Comp   <- aggregate(LoanAmount~Division, data = Sample_A2C, sum)
TotalLoanValue_Fail   <- aggregate(LoanAmount~Division, data = Sample_A2F, sum)
TotalLoanValue_Offers <- aggregate(LoanAmount~Division, data = Sample_Offers, sum)

TotalLoanValue_Exits  <- TotalLoanValue_Exits  %>% filter(!Division == "Unknown")
TotalLoanValue_Comp   <- TotalLoanValue_Comp   %>% filter(!Division == "Unknown")
TotalLoanValue_Fail   <- TotalLoanValue_Fail   %>% filter(!Division == "Unknown")
TotalLoanValue_Offers <- TotalLoanValue_Offers %>% filter(!Division == "Unknown")


#Simple check, 
#loan value in exits should be Completions + App to Failure
#loan value in offers should be Completions + Offer to Failure
round(TotalLoanValue_Exits[, 2], 0) == round(TotalLoanValue_Comp[,2] + TotalLoanValue_Fail[, 2])


#Generating the conversion rates
#Divisions <- TotalLoanValue_Comp$Division
Divisions <- c("PA", "PG01", "PGSL","RGA1", "RL01", "UL01", "TAMI1")

Conversion <- list()
for (Division in Divisions){
  Conversion[["A2C"]][[Division]] <- TotalLoanValue_Comp[match(Division, TotalLoanValue_Comp[,1]),2] / TotalLoanValue_Exits[match(Division, TotalLoanValue_Comp[,1]),2]
}
for (Division in Divisions){
  Conversion[["O2C"]][[Division]] <- TotalLoanValue_Comp[match(Division, TotalLoanValue_Comp[,1]),2] / TotalLoanValue_Offers[match(Division, TotalLoanValue_Comp[,1]),2]
}


#Conversion rate assumptions for 777 are replaced with the average of Phoenix and Rothesay offerings.
Conversion[["A2C"]][["TAMI1"]] <- (((Conversion[["A2C"]][["PG01"]]) + (Conversion[["A2C"]][["PGSL"]]) + (Conversion[["A2C"]][["RL01"]]))/3 )
Conversion[["O2C"]][["TAMI1"]] <- (((Conversion[["O2C"]][["PG01"]]) + (Conversion[["O2C"]][["PGSL"]]) + (Conversion[["O2C"]][["RL01"]]))/3 )


Conversion_Mat <- data.frame(A2C = unlist(Conversion[["A2C"]]), 
                             O2C = unlist(Conversion[["O2C"]]),
                             row.names = Divisions)
