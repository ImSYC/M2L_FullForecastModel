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
#The dataframe "All" will refer to all applications that have occurred in the past 12 months.
#This will be the bases for our analysis.

All <- AllCases %>% filter(ApplicationDate >= (today() %m+% months(-12)))

#Fitting the distribution for all applications over the past 12 months
Fit_Distribution(All)


################################################################################

Sample_Exits <- All %>%
  filter (DPRStatus %in$% ExitCon) %>%
  mutate(PurposeType = ifelse(LoanPurpose == "Drawdown", "Drawdown", "Other")) %>%
  filter (!Division == "Unknown")

Divisions <- SampelExits$Division

LoanPurpose <- c("All", unique(Sample_Exits$PurposeType))
JourneyType <- c("A2C", "O2C")

Conversion <- list()
for(div in divisions){
  

  Sample_Exits <- Sample_Exits %>% filter (Division == div)
  
  for (lp in LoanPurpose){
    
    if(!lp == "All"){
      Sample_Exits <- Sample_Exits %>% filter (PurposeType == lp)
    }
    
    #Samples used
    Sample_Offers         <- Sample_Exits[!is.na(Sample_Exits$OfferIssuedDate), ]
    Sample_A2C            <- Sample_Exits %>% filter(DPRStatus %in% CompCon)
    Sample_A2F            <- Sample_Exits %>% filter(DPRStatus %in% FailCon)
    Sample_O2F            <- Sample_A2F[!is.na(Sample_A2F$OfferIssuedDate),]
    
    #Total Loan Value by funder under each sample
    TotalLoanValue_Exits  <- sum(Sample_Exits$LoanAmount)
    TotalLoanValue_Comp   <- sum(Sample_A2C$LoanAmount)
    TotalLoanValue_Fail   <- sum(Sample_A2F$LoanAmount)
    TotalLoanValue_Offers <- sum(Sample_Offers$LoanAmount)
    
    for (jt in JourneyType){
      if(jt == "O2C"){
        
        Conversion[[lp]][[jt]][[div]] <- TotalLoanValue_Comp / TotalLoanValue_Offers 
        
      } else {
        
        Conversion[[lp]][[jt]][[div]] <- TotalLoanValue_Comp / TotalLoanValue_Exits 
        
      }
    }
    
  }
  
}

Conversion_Mat <- data.frame(A2C = unlist(Conversion[["All"]][["A2C"]]), 
                             O2C = unlist(Conversion[["All"]][["O2C"]]),
                             row.names = Divisions)

Conversion_Mat_Full <- data.frame(A2C_DD = unlist(Conversion[["Drawdown"]][["A2C"]]), 
                                  O2C_DD = unlist(Conversion[["Drawdown"]][["O2C"]]),
                                  A2C_nonDD = unlist(Conversion[["Other"]][["A2C"]]), 
                                  O2C_nonDD = unlist(Conversion[["Other"]][["O2C"]]),
                                  row.names = Divisions)

################################################################################
#Historic App Rates
#Observation period = 8 weeks
AppL <- 8
Past_Mondays <- seq(floor_date(today(), unit = "week") + 1 - 7 * AppL, floor_date(today(), unit = "week") + 1, by = "week")

#Finding the KFIs and Apps of DD and nonDD over the past period.

KFIRates_nonDD <- sapply(Divisions, function(div) {
  KFI1 <- Recent_KFIs %>% filter(Division == div & !LoanPurpose == Drawdown)
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    KFI2 <- KFI1 %>% filter(KFIDate >= Past_Mondays[j] & KFIDate <  Past_Mondays[j + 1])
    sum(KFI2$LoanAmount)
  })
})

KFIRates_nDD <- sapply(Divisions, function(div) {
  KFI1 <- Recent_KFIs %>% filter(Division == div & LoanPurpose == Drawdown)
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    KFI2 <- KFI1 %>% filter(KFIDate >= Past_Mondays[j] & KFIDate <  Past_Mondays[j + 1])
    sum(KFI2$LoanAmount)
  })
})

AppRates_nonDD <- sapply(Divisions, function(div) {
  App1 <- Recent_Apps %>% filter(Division == div & !LoanPurpose == Drawdown)
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    App2 <- App1 %>% filter(ApplicationDate >= Past_Mondays[j] & ApplicationDate <  Past_Mondays[j + 1])
    sum(App2$LoanAmount)
  })
})

AppRates_nDD <- sapply(Divisions, function(div) {
  App1 <- Recent_Apps %>% filter(Division == div & LoanPurpose == Drawdown)
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    App2 <- App1 %>% filter(ApplicationDate >= Past_Mondays[j] & ApplicationDate <  Past_Mondays[j + 1])
    sum(App2$LoanAmount)
  })
})


################################################################################
#Finding the product composition so that we can then forecast product level new business.
#The composition will be used to provide a recommendation, after which we can apply some manual input on what is likely to happen.
Product_Composition_Summary <- AllCases %>%
  filter(ApplicationDate >= (today() - 45)) %>%
  group_by(Division, ProductNameClean) %>%
  summarise(LoanAmount = sum(LoanAmount)) %>%
  #summarise(TotalLoanAmount = sum(LoanAmount), .groups = "drop") %>%
  group_by(Division) %>%
  mutate(Composition = LoanAmount / sum(LoanAmount))

################################################################################
#Completions data table

A <- AllCases %>%
  filter(CompletionDate >= pmin(floor_date(today(), unit = "year"), today()-10),  # Filters for cases completed in the current year or later
         DPRStatus %in% "Completed")  # Filters for cases with "Completed" DPRStatus

# Summarize Loan Quantity and Loan Value
A_LoanAmount_Summary <- A %>%
  group_by(Division, ProductNameClean, LTVBand, LoanPurpose, CompMonth) %>%
  summarize(LoanAmount = sum(LoanAmount))
A_LoanCount_Summary <- A %>%
  group_by(Division, ProductNameClean, LTVBand, LoanPurpose, CompMonth) %>%
  summarize(LoanCount = n()) 

A_LoanAmount <- A_LoanAmount_Summary %>%
  spread(key = CompMonth, value = c("LoanAmount"), fill = 0)
A_LoanCount <- A_LoanCount_Summary %>% 
  spread(key = CompMonth, value = c("LoanCount"), fill = 0)

################################################################################
B <- AllCases %>%
  filter(
    StatusChangeDate >= today() - days(100),               # Filters cases with StatusChangeDate within the last 100 days
    ApplicationDate >= today() - days(365),                # Filters cases with ApplicationDate within the last year
    !is.na(ApplicationDate),                               # Filters out rows with NA in ApplicationDate
    is.na(CompletionDate),                                 # Filters cases with NA in CompletionDate
    !DPRStatus %in% ExitCon,                               # Filters out cases with DPRStatus in ExitCon
    !Division %in% "Unknown"                               # Filters out cases with Division "Unknown"
  )

# Add a new column JourneyType based on OfferIssuedDate
B$JourneyType <- ifelse(is.na(B$OfferIssuedDate), "A2C", "O2C")

# Calculate TimeinPipeline based on JourneyType
B$TimeinPipeline <- ifelse(B$JourneyType == "A2C",
                           pmax(round(as.double(bizdays(B[,"ApplicationDate"], today(), 'Rmetrics/LONDON')), 0), 0),
                           pmax(round(as.double(bizdays(B[, "OfferIssuedDate"], today(), 'Rmetrics/LONDON')), 0), 0))

# Summarize Loan Quantity, Loan Value, and CaseSize
B_Summary <- B %>%
  group_by(Division, ProductNameClean, LTVBand, SolicitorType, LoanPurpose, KeyTier1, JourneyType, TimeinPipeline) %>%
  summarise(
    LoanAmount = sum(LoanAmount),
    LoanCount = n(),
    CaseSize = LoanAmount / LoanCount
  )

# Calculate ConditionalProbability using sapply
B_Summary$ConditionalProbability <- sapply(1:nrow(B_Summary), function(x) {
  lp <- B_Summary$LoanPurpose[x]
  st <- B_Summary$SolicitorType[x]
  dv <- B_Summary$Division[x]
  tp <- B_Summary$TimeinPipeline[x] + 1
  jc <- paste0(substr(B_Summary$JourneyType[x], 1, 2), "C")
  jf <- paste0(substr(B_Summary$JourneyType[x], 1, 2), "F")
  
  alpha <- (1 - SubCDF[[lp]][[jc]][[st]][[tp]]) * Conversion[["All"]][[jc]][[dv]]
  beta <- (1 - Conversion[["All"]][[jc]][[dv]]) * (1 - SubCDF[[lp]][[jf]][[st]][[tp]])
  
  if (length(alpha) == 0 || length(beta) == 0) {
    result <- 0  # Assign a default value when there are missing results
  } else {
    result <- alpha / (alpha + beta)
  }
  
  result
})

# Calculate ExpectedLoanAmount
B_Summary$ExpectedLoanAmount <- B_Summary$ConditionalProbability * B_Summary$LoanAmount

#Calculate the and define the deterministic probability of completion for pipeline cases
Projection <- 500
StartDate <- today()
EndDate <- offset(today(), (Projection - 1), 'Rmetrics/LONDON')
Pipeline_Sequence <- bizseq(StartDate, EndDate, 'Rmetrics/LONDON')
Pipeline_Sequence_Floor <- format(floor_date(Pipeline_Sequence, unit = "month"), format = "%Y/%m/%d")
Pipeline_Colnames <- format(seq(
  from = floor_date(today(), unit = 'year'),
  by = 'month',
  length.out = (interval(StartDate, EndDate) %/% months(1) + 1),
  format = "%Y/%m/%d"
))

results_matrix <- data.frame(matrix(0, nrow(B_Summary), Projection))

foreach(i = 1:nrow(B_Summary)) %do% {
  
  lp <- B_Summary$LoanPurpose[i]
  st <- B_Summary$SolicitorType[i]
  dv <- B_Summary$Division[i]
  tp <- B_Summary$TimeinPipeline[i] + 1
  jc <- B_Summary$JourneyType[i]
  el <- B_Summary$ExpectedLoanAmount[i]
  
  fit <- SubDistribution[[lp]][[jc]][[st]]
  Parameters <- lapply(fit$parameters, function(x) fit[[x]])
  p <- do.call(paste0("p", fit$family[1]), c(tp, Parameters))
  
  #P(X < t | X > a) = (F(X) - F(a)) / (1 - F(a))
  #F(a) is presented here as p, also can be understood as P(X < a)
  #F(X) is presented here as x, also can be understood as P(X < t)
  #P(X < t | X > a) is denoted as B_PDF here.
  x <- numeric(Projection)
  for (j in 1:Projection){
    # Calculate the B_PDF values for each day and store them in B_ExpectedCompletions
    x[j] <- do.call(paste0("p", fit$family[1]), c((tp + j), Parameters))
  }
  
  #x <- round(x, digits = 10)
  B_CDF <- (x - p) / (1 - p)
  #Calculating the probability of completion for each business day from the present
  B_interval_CDF <- c(B_CDF[1], diff(B_CDF))
  B_ExpectedCompletions <- B_interval_CDF * el
  B_ExpectedCompletions <- round(B_ExpectedCompletions, digits = 2)
  
  results_matrix[i,] <- B_ExpectedCompletions
  
}

#Naming the column according to the day we expect the completion to occur
colnames(results_matrix) <- as.Date(Pipeline_Sequence)
B_Summary <- cbind.data.frame(B_Summary, results_matrix)

# Add results_matrix to B_Summary
#B_Aggregate <- cbind.data.frame(B_Summary, results_matrix)

# Create B_Aggregate by summing results_matrix by Division
B_Aggregate <- cbind.data.frame(B_Summary$Division, results_matrix)
names(B_Aggregate)[1] <- "Division"
B_Aggregate <- rowsum(B_Aggregate[, -1], B_Aggregate[, 1], na.rm = TRUE)
B_Aggregate <- cbind.data.frame(rownames(B_Aggregate), B_Aggregate)
names(B_Aggregate)[1] <- "Division"






