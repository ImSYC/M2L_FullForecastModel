#Setting Work Directory
setwd("F:/shiyuan/M2L_FullForecastModel/M2L_FullForecastModel/")

#Start timer
STime <- Sys.time()

#Calling the load packages script
source("load_package.R")
source("process_data.R")

#Setting Output path
outputpath <- file.path("./R_Output", format(Sys.Date(), "%Y-%m"), floor_date(today(), unit= "week"))
dir.create(outputpath, recursive = TRUE)




################################################################################
#Creating a function to fit the distribution of the data
#Will fit a separate distribution for cases under different loan purposes and different solicitor types
#Data frame used to fit this will be 

################################################################################
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

#Excluding 
All <- All %>% filter(!(
  (Division == "TAMI1" & CompletionDate > floor_date(today(), unit = "year")) | 
    (Division == "TAMI1" & ApplicationDate > floor_date(today() - 365, unit = "year") & StatusChangeDate > floor_date(today(), unit = "year") & StatusChangeDate < floor_date(today() - 31, unit = "month"))
))

#Fitting the distribution for all applications over the past 12 months
Fit_Distribution(All)


################################################################################

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

################################################################################


#Looking at  year to date data
#Part A - Completions YTD
#Part B - Pipeline expected completions
#Part C - Expected New Business
#


################################################################################
# Part A: YTD Completions
A <- AllCases %>%
  filter(CompletionDate >= floor_date(today(), unit = "year"),  # Filters for cases completed in the current year or later
         DPRStatus %in% "Completed")  # Filters for cases with "Completed" DPRStatus

# Summarize Loan Quantity and Loan Value
A_Summary <- A %>%
  group_by(Division, ProductNameClean, LTVBand, SolicitorType, LoanPurpose, KeyTier1, CompletionDate, CompMonth) %>%
  summarize(LoanAmount = sum(LoanAmount),
            LoanCount = n()) 

# Store a copy of A_Summary
A_Summary_print <- A_Summary

# Aggregate LoanAmount by Division and CompletionDate
A_Aggregate <- A_Summary %>%
  group_by(Division, CompletionDate) %>%
  summarize(LoanAmount = sum(LoanAmount))

# Melt the data frame
A_Aggregate <- reshape2::melt(A_Aggregate, id.vars = c("Division", "CompletionDate"), value.name = "LoanAmount")

# Cast the data into a wide format
A_Aggregate <- reshape2::dcast(A_Aggregate, Division ~ CompletionDate, value.var = "LoanAmount")


################################################################################
# Part B: Current Application+ Pipeline
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
  
  alpha <- (1 - SubCDF[[lp]][[jc]][[st]][[tp]]) * Conversion[[jc]][[dv]]
  beta <- (1 - Conversion[[jc]][[dv]]) * (1 - SubCDF[[lp]][[jf]][[st]][[tp]])
  
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

################################################################################
#Part c
#
Past_Year_Cases <- All %>% filter(DPRStatus %in% "Completed")
Past_Year_Fit <- fitDist(Past_Year_Cases$A2C, k = 2, type = "realplus")
Past_Year_Param <- lapply(Past_Year_Fit$parameters, function(x) Past_Year_Fit[[x]])
Past_Year_CDF <- lapply(T, function(x) do.call(paste0("p", Past_Year_Fit$family[1]), c(x, Past_Year_Param)))
CompletionCDF <- cbind(Past_Year_CDF)

################################################################################
#Part D
#Other Data Pieces

AppL <- 8
Past_Mondays <- seq(floor_date(today(), unit = "week") + 1 - 7*AppL, floor_date(today(), unit = "week") + 1, by = "week")
AppRates <- sapply(Divisions, function(div) {
  Apps1 <- AllCases %>% filter(Division == div)
  sapply(seq_along(Past_Mondays[1:AppL]), function(j) {
    Apps2 <- Apps1 %>% filter(ApplicationDate >= Past_Mondays[j] & ApplicationDate <  Past_Mondays[j + 1])
    sum(Apps2$LoanAmount)
  })
})
colnames(AppRates) <- Divisions
rownames(AppRates) <- format(Past_Mondays[1:AppL], "%Y/%m/%d")
AppRates <- as.data.frame(AppRates[rev(rownames(AppRates)), ])

AB_Aggregate <- merge(A_Aggregate, B_Aggregate, by = "Division", all = TRUE)
AB_Aggregate_t <- t(AB_Aggregate)
colnames(AB_Aggregate_t) <- AB_Aggregate_t[1, ]
AB_Aggregate_t <- AB_Aggregate_t[-1, ]
AB_Aggregate_cs <- as.data.frame(apply(AB_Aggregate_t, 2, function(x) {
  cumsum(ifelse(is.na(x), 0, x))
}))



#Summarising Loan Quantity and Loan Value, by Division, ProductName, SolicitorType, LoanPurpose, CompMonth, KeyTier1
Product_Composition_Summary <- AllCases %>%
  filter(ApplicationDate >= (today() - 45)) %>%
  group_by(Division, LTVBand, ProductNameClean) %>%
  summarise(LoanAmount = sum(LoanAmount)) %>%
  #summarise(TotalLoanAmount = sum(LoanAmount), .groups = "drop") %>%
  group_by(Division) %>%
  mutate(Composition = LoanAmount / sum(LoanAmount))

#Case size by funder
CaseSize <- All %>%
  filter(ApplicationDate >= floor_date((today() - months(3)), unit = "month")) %>%
  group_by(Division) %>%
  summarise(LoanAmount = sum(LoanAmount),
            LoanCount = n()) %>%
  mutate(CaseSize = LoanAmount / LoanCount)

################################################################################


################################################################################
#Part F
#Outputting
#Setup forecast model template
inputfile <-  file.path( "./R_Input/Template_Forecast_v7.1.xlsm")

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

# Change the column name from "X777" to "777"
#names(inputRunRates)[names(inputRunRates) == "X777"] <- "777"


outputfilename <- paste(today(),"Forecast_v7.2_temp.xlsm", sep = "_")
outputfilepath <- file.path(outputpath, outputfilename)

wb <- loadWorkbook(inputfile)
deleteData(wb, "R_Data1", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data2", 1:1000, 2:10000, TRUE)
deleteData(wb, "R_Data3", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data4", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data5", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data6", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data7", 1:1000, 1:10000, TRUE)

options("openxlsx.dateFormat" = "dd/mm/yyyy")

for(i in 1:ncol(inputRunRates)){
  if(!names(inputRunRates[i]) %in% c("PhoenixM2L", "PhoenixSLHF")){
    writeData(wb, names(inputRunRates[i]), inputRunRates[,i], startRow = 18, startCol = 2, colNames = FALSE, rowNames = FALSE, withFilter = FALSE)
  }
}

writeData(wb, "R_Data1", A_Summary, startRow = 1, colNames = TRUE, rowNames = FALSE, withFilter = FALSE)
writeData(wb, "R_Data2", B_Summary, startRow = 2, colNames = TRUE, rowNames = FALSE, withFilter = FALSE)
#writeDataTable(wb,"AB_Aggregate", AB_Aggregate_cs, colNames = TRUE, rowNames = TRUE, withFilter = FALSE)
writeData(wb, "R_Data3", CompletionCDF, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, "R_Data4", Pipeline_Sequence, startRow = 1, colNames = TRUE, rowNames = TRUE)
writeData(wb, "R_Data5", Conversion_Mat, startRow = 1, colNames = TRUE, rowNames = TRUE)
writeData(wb, "R_Data6", AppRates, startRow = 1, colNames = TRUE, rowNames = TRUE)
writeData(wb, "R_Data7", Product_Composition_Summary, startRow = 1, colNames = TRUE, rowNames = FALSE)
saveWorkbook(wb, outputfilepath, overwrite = TRUE)
saveWorkbook(wb, file.path("F:\\06. Data & BI - M2L Analytics\\M2L_Analytics\\nb\\M2L0018 - 777 Pricing Committee Pack\\Handover\\WeeklyForecast", outputfilename), overwrite = TRUE)

ETime <- Sys.time()

ETime - STime





