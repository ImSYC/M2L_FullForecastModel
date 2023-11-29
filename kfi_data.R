#Setting Work Directory
setwd("F:/shiyuan/M2L_FullForecastModel/M2L_FullForecastModel/")

#Calling the load packages script
source("load_package.R")
source("process_data.R")

#Setting Output path
outputpath <- file.path("./R_Output", format(Sys.Date(), "%Y-%m"), floor_date(today(), unit= "week"))
dir.create(outputpath, recursive = TRUE)



################################################################################
#Removing Duplicate KFIs
################################################################################
#Finding the rate of KFI's and Apps recently

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
  filter(!is.na(ApplicationDate) | !DPRStatus %in% c("App Decline", "Cancelled / Declined", "Expired", "Decision Decline")) %>%
  group_by(App1Forename, App1Surname, App1DOB, LoanPurpose, BrokerForename, BrokerSurname) %>%
  do(filter_duplicated_kfi(.)) %>%
  ungroup()

# Remove the diff_days column using indexing
Recent_KFIs <- Recent_KFIs[, -which(names(Recent_KFIs) == "diff_days")]


################################################################################
AppL <- 8
Past_Mondays <- seq(floor_date(today(), unit = "week") + 1 - 7*AppL, floor_date(today(), unit = "week") + 8, by = "week")
KFIRates <- sapply(Divisions, function(div) {
  KFI1 <- Recent_KFIs %>% filter(Division == div)
  sapply(seq_along(Past_Mondays[1:(length(Past_Mondays)-1)]), function(j) {
    KFI2 <- KFI1 %>% filter(KFIDate >= Past_Mondays[j] & KFIDate <  Past_Mondays[j + 1])
    sum(KFI2$LoanAmount)
  })
})

colnames(KFIRates) <- Divisions
rownames(KFIRates) <- format(Past_Mondays[1:(length(Past_Mondays)-1)], "%Y/%m/%d")
KFIRates <- as.data.frame(KFIRates[rev(rownames(KFIRates)), ])

################################################################################
################################################################################
AppL <- 8
Past_Mondays <- seq(floor_date(today(), unit = "week") + 1 - 7 * AppL, floor_date(today(), unit = "week") + 1, by = "week")

# Initialize a list to store the results
AppRates_LoanPurpose <- list()

# Loop over Divisions
for (i in 1:length(Divisions)){
  div <- Divisions[i]
  
  # Create an empty dataframe to store the results for this Division
  division_df <- data.frame(Week = rev(Past_Mondays[1:AppL]))
  
  for(j in 1:length(LoanPurpose)){
    lp <- LoanPurpose[j]
    Apps1 <- AllCases %>% 
      filter(Division == div & LoanPurpose == lp)
    
    # Calculate loan sums for each week
    loan_sums <- sapply(seq_along(Past_Mondays[1:AppL]), function(k) {
      Apps2 <- Apps1 %>% filter(ApplicationDate >= Past_Mondays[k] & ApplicationDate <  Past_Mondays[k + 1])
      sum(Apps2$LoanAmount)
    })
    
    # Add a column to the dataframe for this loan purpose
    division_df[lp] <- rev(loan_sums)
  }
  
  # Add the dataframe for this Division to the list
  AppRates_LoanPurpose[[div]] <- division_df
}


################################################################################
#Monitor for Apex KFIs
#Creating a table specific for Apex
Apex_KFI <- Recent_KFIs %>% 
  filter(Division == "TAMI1" & KFIDate >= as.Date("2023-09-02"))%>%
  group_by(KFIDate, ProductName) %>%
  summarize(TotalLoanAmount = sum(LoanAmount, na.rm = TRUE))

# Melt the dataframe
Apex_KFI <- reshape2::melt(Apex_KFI, id.vars = c("ProductName", "KFIDate"), value.name = "LoanAmount")

# Cast (reshape) the dataframe
Apex_KFI <- reshape2::dcast(Apex_KFI, KFIDate ~ ProductName, value.var = "LoanAmount")

# Arrange the rows by KFIDate (largest at the top)
Apex_KFI <- Apex_KFI %>%
  arrange(desc(KFIDate))

################################################################################
#Outputting
outputfilename <- paste(today(),"KFI_Summary.xlsx", sep = "_")

wb3 <- createWorkbook()
addWorksheet(wb3, "KFIs")
addWorksheet(wb3, "Apex")
writeData(wb3, "KFIs", KFIRates, startRow = 1, colNames = TRUE, rowNames = TRUE)
writeData(wb3, "Apex", Apex_KFI, startRow = 1, colNames = TRUE, rowNames = TRUE)
addWorksheet(wb3, "Apps -> ")
# Loop through each Division and add a worksheet for it
for (div in Divisions) {
  division_df <- AppRates_LoanPurpose[[div]]
  addWorksheet(wb3, div)
  writeData(wb3, div, division_df, startRow = 1, colNames = TRUE, rowNames = FALSE)
}
saveWorkbook(wb3, file.path(outputpath, outputfilename), overwrite = TRUE)
