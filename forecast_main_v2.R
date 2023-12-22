#Setting Work Directory
setwd("F:/shiyuan/M2L_FullForecastModel/M2L_FullForecastModel/")

#Start timer
STime <- Sys.time()
Production_Year <- year(today())

################################################################################
#1. Calling the load packages script
source("load_package.R")
source("process_data.R")

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

All_Sample_Exits <- All %>%
  filter (DPRStatus %in% ExitCon) %>%
  mutate(PurposeType = ifelse(LoanPurpose == "Drawdown", "Drawdown", "Other")) %>%
  filter (!Division == "Unknown")

Divisions <- unique(All_Sample_Exits$Division)

LoanPurpose <- c("All", unique(All_Sample_Exits$PurposeType))
JourneyType <- c("A2C", "O2C")

Conversion <- list()
for(div in Divisions){

  Sample_Exits_1 <- All_Sample_Exits %>% filter (Division == div)
  
  for (lp in LoanPurpose){
    
    Sample_Exits <- Sample_Exits_1
    
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
Conversion_Mat[is.na(Conversion_Mat)] <- 0

Conversion_Mat_Full <- data.frame(A2C_DD = unlist(Conversion[["Drawdown"]][["A2C"]]), 
                                  O2C_DD = unlist(Conversion[["Drawdown"]][["O2C"]]),
                                  A2C_Other = unlist(Conversion[["Other"]][["A2C"]]), 
                                  O2C_Other = unlist(Conversion[["Other"]][["O2C"]]),
                                  row.names = Divisions)
Conversion_Mat_Full[is.na(Conversion_Mat_Full)] <- 0

################################################################################
#Historic App Rates
#Observation period = 8 weeks
AppL <- 8
Past_Mondays <- seq(floor_date(today(), unit = "week") + 1 - 7 * AppL, floor_date(today(), unit = "week") + 1, by = "week")

#Define dataframe Recent_KFIs and Recent_Apps
# Define a function to filter rows within 30 days of each other
filter_duplicated_kfi <- function(data) {
  data %>%
    arrange(ApplicationCreatedDate) %>%
    mutate(diff_days = c(NA, diff(ApplicationCreatedDate))) %>%
    filter(is.na(diff_days) | diff_days > 30)
}

# Modify the dataframe as described
Recent_KFIs <- AllCases %>% 
  filter(ApplicationCreatedDate >= (today() %m+% months(-3))) %>%
  group_by(App1Forename, App1Surname, App1DOB, LoanPurpose, BrokerForename, BrokerSurname) %>%
  do(filter_duplicated_kfi(.)) %>%
  ungroup()

# Remove the diff_days column using indexing
Recent_KFIs <- Recent_KFIs[, -which(names(Recent_KFIs) == "diff_days")]

Recent_Apps <- AllCases


#Finding the KFIs and Apps of DD and Other over the past period.

KFIRates_Other <- sapply(Divisions, function(div) {
  KFI1 <- Recent_KFIs %>% filter(Division == div & !LoanPurpose == "Drawdown")
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    KFI2 <- KFI1 %>% filter(KFIDate >= Past_Mondays[j] & KFIDate <  Past_Mondays[j + 1])
    sum(KFI2$LoanAmount)
  })
})
rownames(KFIRates_Other) <- format(Past_Mondays[1:AppL], "%Y/%m/%d")
KFIRates_Other <- as.data.frame(KFIRates_Other[rev(rownames(KFIRates_Other)), ])

KFIRates_DD <- sapply(Divisions, function(div) {
  KFI1 <- Recent_KFIs %>% filter(Division == div & LoanPurpose == "Drawdown")
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    KFI2 <- KFI1 %>% filter(KFIDate >= Past_Mondays[j] & KFIDate <  Past_Mondays[j + 1])
    sum(KFI2$LoanAmount)
  })
})
rownames(KFIRates_DD) <- format(Past_Mondays[1:AppL], "%Y/%m/%d")
KFIRates_DD <- as.data.frame(KFIRates_DD[rev(rownames(KFIRates_DD)), ])

AppRates_Other <- sapply(Divisions, function(div) {
  App1 <- Recent_Apps %>% filter(Division == div & !LoanPurpose == "Drawdown")
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    App2 <- App1 %>% filter(ApplicationDate >= Past_Mondays[j] & ApplicationDate <  Past_Mondays[j + 1])
    sum(App2$LoanAmount)
  })
})
rownames(AppRates_Other) <- format(Past_Mondays[1:AppL], "%Y/%m/%d")
AppRates_Other <- as.data.frame(AppRates_Other[rev(rownames(AppRates_Other)), ])

AppRates_DD <- sapply(Divisions, function(div) {
  App1 <- Recent_Apps %>% filter(Division == div & LoanPurpose == "Drawdown")
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    App2 <- App1 %>% filter(ApplicationDate >= Past_Mondays[j] & ApplicationDate <  Past_Mondays[j + 1])
    sum(App2$LoanAmount)
  })
})
rownames(AppRates_DD) <- format(Past_Mondays[1:AppL], "%Y/%m/%d")
AppRates_DD <- as.data.frame(AppRates_DD[rev(rownames(AppRates_DD)), ])


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
A_Summary <- A %>%
  mutate(CompMonth = as.numeric(CompMonth)) %>%
  group_by(Division, ProductNameClean, LTVBand, LoanPurpose, CompMonth) %>%
  summarize(
    LoanAmount = sum(LoanAmount),
    LoanCount = n()
  )

A_LoanAmount <- A_Summary %>%
  spread(key = CompMonth, value = c("LoanAmount"), fill = 0)
A_LoanCount <- A_Summary %>% 
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

################################################################################
#Specify out the CDF of DD and non DD
# Assuming NB_Other_CDF and NB_DD_CDF have more than 500 values
NB_Other_CDF <- round(c(diff(unlist(MasterCDF[["Remortgage"]][["A2C"]])), rep(0, 500 - 365)), 6)
NB_DD_CDF <- round(c(diff(unlist(MasterCDF[["Drawdown"]][["A2C"]][1:(365*3)])), rep(0, 500 - 365)), 6)

# Assuming NB_Other_CDF and NB_DD_CDF have more than 500 values
NB_Matrix_Other <- matrix(0, nrow = 500, ncol = 500)
NB_Matrix_Drawdown <- matrix(0, nrow = 500, ncol = 500)

for (i in 0:499) {
  NB_Matrix_Other[i, ] <- c(rep(0, i), NB_Other_CDF[1:(500 - i)])
  NB_Matrix_Drawdown[i, ] <- c(rep(0, i), NB_DD_CDF[1:(500 - i)])
}

# Create a sequence of business days
start_date <- today() + 1
NB_Labels <- bizseq(start_date, start_date + 1000, 'Rmetrics/LONDON')[1:500]

# Name the rows and columns of the matrices
colnames(NB_Matrix_Other) <- rownames(NB_Matrix_Other) <- NB_Labels
colnames(NB_Matrix_Drawdown) <- rownames(NB_Matrix_Drawdown) <- NB_Labels

#the rows are the assumed days from which the apps originate
#the columns are the days which we assume the apps complete.
#By multiplying and summing the assumed number of apps each day against each column vector we can calculate the number of completions expected each day.


################################################################################
#Number of working days in each month
#To be used for pipeline calculations
WorkingDays <- data.frame(matrix(0, nrow = 12, ncol = 2))
rownames(WorkingDays) <- seq(1,12)
colnames(WorkingDays) <- c(year(today()), year(today() + 365))
Current_Month <- month(today())
Current_Year <- year(today())

for(i in 1:2){
  for(j in 1:12){
    if(colnames(WorkingDays)[i] == Current_Year & j == Current_Month){
      WorkingDays[j,i] <- bizdays(today(), ceiling_date(today(), "month") - 1, 'Rmetrics/LONDON')
    } else if(colnames(WorkingDays)[i] > Current_Year | (colnames(WorkingDays)[i] == Current_Year & j > Current_Month)) {
      start_date <- ymd(paste(colnames(WorkingDays)[i], j, 1))
      end_date <- ceiling_date(start_date, unit = "month") - days(1)
      WorkingDays[j,i] <- bizdays(start_date, end_date, 'Rmetrics/LONDON') + 1
    }
  }
}
################################################################################
NB_Timeline <- data.frame(matrix(1, nrow = 500, ncol = 2))
colnames(NB_Timeline) <- c("Week", "Date")

NB_Timeline$Date <- Pipeline_Sequence
for(i in 2:500){
  if(as.numeric(NB_Timeline$Date[i] - NB_Timeline$Date[i-1]) >= 2){
    NB_Timeline$Week[i] <- pmin(NB_Timeline$Week[i-1] + 1 , 9)
  } else{
    NB_Timeline$Week[i] <- NB_Timeline$Week[i-1]
  }
}

################################################################################
PipelineSize_Other <- B_Summary %>%
  group_by(Division) %>%
  summarise(SumLoanAmount = sum(LoanAmount[LoanPurpose != "Drawdown"], na.rm = TRUE))
PipelineSize_Other <- as.data.frame(PipelineSize_Other)

PipelineSize_Drawdown <- B_Summary %>%
  group_by(Division) %>%
  summarise(SumLoanAmount = sum(LoanAmount[LoanPurpose == "Drawdown"], na.rm = TRUE))
PipelineSize_Drawdown <- as.data.frame((PipelineSize_Drawdown))


################################################################################
#Setting Output path
outputpath <- file.path("./R_Forecasts")
dir.create(outputpath, recursive = TRUE)

# Get the list of files in the folder
files <- list.files(path = "./R_Forecasts", pattern = "\\.xlsx$", full.names = TRUE)

# Check if there are any files
if (length(files) > 0) {
  # Find the latest file based on modification time
  latest_file <- files[which.max(file.info(files)$mtime)]
  
  # Load the workbook
  wb <- loadWorkbook(latest_file)
  
  # Now, 'wb' is the workbook for the latest file
} else {
  # No files in the folder
  cat("No files found in the folder.\n")
}

#Setting up output name
outputfilename <- paste(format(Sys.time(), "%Y%m%d_%H%M%S"), "Forecast_v8.xlsx", sep = "_")
outputfilepath <- file.path(outputpath, outputfilename)


# A function to handle repeated operations in your loop
writeDataToSheet <- function(wb, sheet, data, startCol, startRow) {
  for (i in seq_along(startRow)) {
    writeData(wb, sheet = sheet, data, startCol = startCol[i], startRow = startRow[i], colNames = FALSE, rowNames = FALSE)
  }
}

for (div in Divisions) {
  # Delete data
  deleteData(wb, sheet = div, cols = 2:3, rows = c(4, 15:17, 20:27, 30:37, 55:100), gridExpand = TRUE)
  
  # Write data
  writeDataToSheet(wb, div, year(today()), startCol = 2, startRow = 4)
  writeDataToSheet(wb, div, data = year(today()) + 1, startCol = 3, startRow = 4)
  #writeDataToSheet(wb, div, data = PipelineSize_Other[PipelineSize_Other$Division == div, 2], startCol = 2, startRow = 15)
  #writeDataToSheet(wb, div, data = PipelineSize_Drawdown[PipelineSize_Other$Division == div, 2], startCol = 2, startRow = 15)
  writeDataToSheet(wb, div, data = t(Conversion_Mat_Full[div, 3:4]), startCol = 2, startRow = 16)
  writeDataToSheet(wb, div, data = t(Conversion_Mat_Full[div, 1:2]), startCol = 3, startRow = 16)
  writeDataToSheet(wb, div, data = rev(format(Past_Mondays[1:AppL], "%Y/%m/%d")), startCol = 1, startRow = 20)
  writeDataToSheet(wb, div, data = rev(format(Past_Mondays[1:AppL], "%Y/%m/%d")), startCol = 1, startRow = 30)
  writeDataToSheet(wb, div, data = c(KFIRates_Other[div], AppRates_Other[div]), startCol = c(2, 3), startRow = 20)
  writeDataToSheet(wb, div, data = c(KFIRates_DD[div], AppRates_DD[div]), startCol = c(2, 3), startRow = 30)
  
  # Write Product_Composition_Summary
  Productlist_nb <- as.data.frame(Product_Composition_Summary %>% filter(Division == div))
  writeDataToSheet(wb, div, data = Productlist_nb[c("ProductNameClean", "Composition")], startCol = 1, startRow = 55)
  
  # Write unique product lists
  Productlist_A <- unique(A$ProductNameClean[A$Division == div])
  Productlist_B <- unique(B$ProductNameClean[B$Division == div])
  Productlist_YTD <- union(Productlist_A, Productlist_B)
  deleteData(wb, sheet = div, cols = 6:7, rows = 31:500, gridExpand =  TRUE)
  writeDataToSheet(wb, div, data = rep(Productlist_YTD, each = 3), startCol = 7, startRow = 31)
  writeDataToSheet(wb, div, data = rep(c(1,2,3), length(Productlist_YTD)), startCol = 6, startRow = 31)
  
  #Print working days
  deleteData(wb, sheet = div, cols = 40, rows = 7:506, gridExpand =  TRUE)
  writeDataToSheet(wb, div, data = NB_Timeline, startCol = 40, startRow = 7)
}

deleteData(wb, "R_Data1", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data2", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data3", 1:1000, 1:10000, TRUE)
deleteData(wb, "R_Data4", 1:1000, 1:10000, TRUE)
deleteData(wb, "Control", 2:3, 3:16, TRUE)

writeData(wb, "R_Data1", A_Summary, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
writeData(wb, "R_Data2", B_Summary, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
writeData(wb, "R_Data3", NB_Matrix_Other, startCol = 1, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, "R_Data4", NB_Matrix_Drawdown, startCol = 1, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, "Control", WorkingDays, startCol = 2, startRow = 3, colNames = FALSE, rowNames = FALSE)

saveWorkbook(wb, outputfilepath, overwrite = FALSE)



