#Purpose:
#Pull data from DPRDailyExtract and DPRDailyExtract_DENCOM
#Trim the unnecessary columns
#Format the data into a usable table
#Add helper columns as needed

#Create connections to server
#con1 - Connected for daily extract (Raw case data)
#con2 - Connected for LTV Band definitions
con1 <- dbConnect(odbc::odbc(), "143", Database = "Z_M2L_ANALYTICS_REPO")
con2 <- dbConnect(odbc::odbc(), "143", Database = "key_insight_sandbox")

################################################################################
#Reading tables from connections
################################################################################
#Read M2L and SLHF table from DB
M2L <- dbReadTable(con1,"tbl_DPRDailyExtract")
SLHF <- dbReadTable(con1,"tbl_DPRDailyExtract_DENCOM")


#Read LTV grid info
LTVGrid <- dbReadTable(con2, "kj_LTV_Segments_master_formatted")
colnames(LTVGrid) <- sub("\\.", "", colnames(LTVGrid))


################################################################################
#Creating the original combined data tables
################################################################################
#Listing the columns to keep within the datatables
keep <- c(
  "ApplicationCaseNumber",
  "DPRStatus",
  "ApplicationCreatedDate",
  "KFIDate",
  "ApplicationDate",
  "OfferIssuedDate",
  "CompletionDate",
  "StatusChangeDate",
  "Division",
  "ProductRange",
  "ProductName",
  "InterestRateAER",
  "LoanPurpose",
  "LoanAmount",
  "CashAdvance",
  "InitialAdvance",
  "CommittedFacility",
  "RemainingCommitment",
  "LoanBalanceAsAtLastAnniversary",
  "AccruedInterest",
  "TotalOutstandingBalanceIncludingAllAccruedInterest",
  "ESTLTV",
  "CurrentLTV",
  "OriginalLTV",
  "EstimatedPropertyValue",
  "SingleorJointApplication",
  "App1Forename",
  "App1Surname",
  "App1DOB",
  "App2DOB",
  "App1Gender",
  "App2Gender",
  "Region",
  "PropertyType",
  "BrokerFirm",
  "ApplicantSolicitor",
  "LenderSolicitor",
  "BrokerForename",
  "BrokerSurname"
)


#Trimming the datatables previously pulled from the server
M2L <- M2L[keep]
SLHF <- SLHF[keep]
#Assigning division to SLHF, Phoenix Standard Life,
#Standard Life does not have a DIVISION column as it is stored separately.
SLHF[,'Division'] <- "PGSL" 

#Combining to create the full data table
AllCases <- rbind.data.frame(M2L,SLHF)


################################################################################
#Adding needed columns to the dataframe
################################################################################
#Creating the functions required
#

#Adding lag times to the data table
AllCases[["A2F"]] <- as.double(bizdays(AllCases[["ApplicationDate"]], AllCases[["StatusChangeDate"]], 'Rmetrics/LONDON'))
AllCases[["A2F"]][AllCases[["A2F"]] < 0] <- 0
AllCases[["O2F"]] <- as.double(bizdays(AllCases[["OfferIssuedDate"]], AllCases[["StatusChangeDate"]], 'Rmetrics/LONDON'))
AllCases[["O2F"]][AllCases[["O2F"]] < 0] <- 0
AllCases[["A2C"]] <- as.double(bizdays(AllCases[["ApplicationDate"]], AllCases[["CompletionDate"]], 'Rmetrics/LONDON'))
AllCases[["A2C"]][AllCases[["A2C"]] < 0] <- 0
AllCases[["O2C"]] <- as.double(bizdays(AllCases[["OfferIssuedDate"]], AllCases[["CompletionDate"]], 'Rmetrics/LONDON'))
AllCases[["O2C"]][AllCases[["O2C"]] < 0] <- 0
AllCases[["AppMonth"]]  <- month(AllCases[["ApplicationDate"]])
AllCases[["CompMonth"]] <- month(AllCases[["CompletionDate"]])


#Mapping KeyTier1Other
AllCases <- AllCases %>% mutate(KeyTier1 = 
                                  case_when(AllCases$BrokerFirm == "Age Partnership  Ltd"               ~ "Age Partnership",
                                            AllCases$BrokerFirm == "Age Partnership Limited"            ~ "Age Partnership",
                                            AllCases$BrokerFirm == "Bower"                              ~ "Bower",
                                            AllCases$BrokerFirm == "Equity Release Supermarket Limited" ~ "Equity Release Supermarket",
                                            AllCases$BrokerFirm == "Key Equity Release"                 ~ "Key Equity Release",
                                            AllCases$BrokerFirm == "Key Retirement Solutions Ltd"       ~ "Key Partnerships",
                                            AllCases$BrokerFirm == "Responsible Life Limited"           ~ "Responsible",
                                            AllCases$BrokerFirm == "RSUK"                               ~ "Responsible",
                                            AllCases$BrokerFirm == "The Equity Release Experts"         ~ "TERE",
                                            AllCases$BrokerFirm == "The Right Equity Release Limited"   ~ "TRER",
                                            AllCases$BrokerFirm == "Cavendish Equity Release"           ~ "Tier 1",
                                            AllCases$BrokerFirm == "Financial Experience Ltd"           ~ "Tier 1",
                                            AllCases$BrokerFirm == "HBFS Equity Release Limited"        ~ "Tier 1",
                                            AllCases$BrokerFirm == "Hub Financial Solutions Limited"    ~ "Tier 1",
                                            AllCases$BrokerFirm == "My Equity Release Expert"           ~ "Tier 1",
                                            AllCases$BrokerFirm == "Sixty Plus Ltd"                     ~ "Tier 1",
                                            AllCases$BrokerFirm == "Stepchange Financial Solutions"     ~ "Tier 1",
                                            AllCases$BrokerFirm == "Retirement Solutions (UK) Limited"  ~ "Tier 1",
                                            AllCases$BrokerFirm == "Fluent Lifetime Limited"            ~ "Tier 1",
                                            AllCases$BrokerFirm == "Age Lifetime"                       ~ "Sunlife",
                                            AllCases$BrokerFirm == "Age Partnership Limited"            ~ "Sunlife",
                                            TRUE                                                        ~ "Tail"))


#Mapping Broker Finance Groups
AllCases <- AllCases %>% mutate(BrokerFinance =
                                  case_when(AllCases$KeyTier1 == "Key Equity Release"         ~ "KeyOB",
                                            AllCases$KeyTier1 == "Key Partnerships"           ~ "KeyWOM",
                                            AllCases$KeyTier1 == "Age Partnership"            ~ "Age",
                                            AllCases$KeyTier1 == "Sunlife"                    ~ "Age",
                                            AllCases$KeyTier1 == "Bower"                      ~ "SpecialistnT1",
                                            AllCases$KeyTier1 == "Equity Release Supermarket" ~ "SpecialistnT1",
                                            AllCases$KeyTier1 == "Responsible"                ~ "SpecialistnT1",
                                            AllCases$KeyTier1 == "TERE"                       ~ "SpecialistnT1",
                                            AllCases$KeyTier1 == "TRER"                       ~ "SpecialistnT1",
                                            AllCases$KeyTier1 == "Tier 1"                     ~ "SpecialistnT1",
                                            TRUE                                              ~ "Tail"))


#Mapping Solicitor type
AllCases <- AllCases %>% mutate(SolicitorType =
                                  case_when(str_detect(str_to_lower(ApplicantSolicitor), "equilaw")       ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "ashford")       ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "gilroy")        ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "gw legal")      ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "gwlegal")       ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "g w legal")     ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "adlington")     ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "boyd")          ~ "Panel",
                                            ApplicantSolicitor == "Thomas Boyd Whyte"                     ~ "NonPanel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "forever")       ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "goldsmith w")   ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "poynton")       ~ "Panel",
                                            str_detect(str_to_lower(ApplicantSolicitor), "roger brooker") ~ "Panel",
                                            TRUE                                                          ~ "NonPanel"))


#Reformatting the LoanPurpose Column
AllCases$LoanPurpose <- gsub(" ", "", AllCases$LoanPurpose)


#Adding a clean product name column
AllCases$ProductNameClean <- AllCases$ProductName
AllCases$ProductNameClean <- gsub("\\d+\\.\\d+%", "", AllCases$ProductNameClean)
AllCases$ProductNameClean <- gsub("F/Adv| F/Advance| F\\\\Adv| Further Advance| Drawdown| Lump Sum| DD| FA| LS", "", AllCases$ProductNameClean)
AllCases$ProductNameClean <- gsub("(\\D)(\\d)", "\\1 \\2", AllCases$ProductNameClean)
AllCases$ProductNameClean <- gsub("\\s+$", "", AllCases$ProductNameClean)
AllCases$ProductNameClean <- gsub("  ", " ", AllCases$ProductNameClean)


#Reformatting the dates column
AllCases$ApplicationCreatedDate <- as.Date(AllCases$ApplicationCreatedDate, format = "%Y/%m/%d")
AllCases$KFIDate                <- as.Date(AllCases$KFIDate, format = "%Y/%m/%d")
AllCases$ApplicationDate        <- as.Date(AllCases$ApplicationDate, format = "%Y/%m/%d")
AllCases$OfferIssuedDate        <- as.Date(AllCases$OfferIssuedDate, format = "%Y/%m/%d")
AllCases$CompletionDate         <- as.Date(AllCases$CompletionDate, format = "%Y/%m/%d")
AllCases$StatusChangeDate       <- as.Date(AllCases$StatusChangeDate, format = "%Y/%m/%d")


#Adding a function of calculating age
age <- function(from, to) {
  from_lt <- as.POSIXlt(from, origin = "01-01-1900")
  to_lt <- as.POSIXlt(to, origin = "01-01-1900")
  
  age <- to_lt$year - from_lt$year
  
  ifelse(to_lt$mon < from_lt$mon |
           (to_lt$mon == from_lt$mon & to_lt$mday < from_lt$mday),
         age - 1, age)
}


#Calculating the age of Applicant 1 and 2, and calculated LTV
AllCases <- AllCases %>%
  mutate(AgeCalcDate = as.Date(ifelse(!is.na(CompletionDate), CompletionDate, ifelse(!is.na(ApplicationDate), ApplicationDate, ApplicationCreatedDate)), origin = "1970-01-01"),
         App1Age = ifelse(!is.na(App1DOB), 
                          age(App1DOB, AgeCalcDate),
                          NA),
         App2Age = ifelse(!is.na(App2DOB),
                          age(App2DOB, AgeCalcDate),
                          NA)) %>%
  mutate(YoungestAppAge = pmin(App1Age, App2Age, na.rm = TRUE)) %>%
  mutate(CalculatedLTV = ifelse(YoungestAppAge < 54 | CommittedFacility / EstimatedPropertyValue >= 0.7,
                                NA,
                                CommittedFacility / EstimatedPropertyValue))


#Adding a function to calculate the LTV band of the case
matchLTVBand <- function(age, appType, ltv) {
  # check for missing values
  if (is.na(ltv) || ltv == 0) {
    return(NA)
  } else {
    # look up the row in the LTV grid
    gridRow <- LTVGrid[LTVGrid$Age == age & LTVGrid$AppType == appType, -c(1,2)]
    # find the leftmost column in the row that is greater than ltv
    ltvBand <- if (length(gridRow) > 0) {
      ifelse(ltv > gridRow[15], names(gridRow)[16], names(gridRow)[which.max(ltv <= gridRow)])
    } else {
      NA
    }
    # return the LTV band
    return(ltvBand)
  }
}

# use mapply to apply the function to each row of AllCases
AllCases$LTVBand <- mapply(matchLTVBand, AllCases$YoungestAppAge, AllCases$SingleorJointApplication, AllCases$CalculatedLTV)


################################################################################
#Definitions
FailCon  <- c("App Decline", "Cancelled / Declined", "Expired", "Decision Decline")
CompCon  <- c("Completed", "Delayed Completion")
ExitCon  <- c(FailCon, CompCon)
T        <- seq(0, 365, 1)
T_length <- length(T)
LoanPurpose <- unique(AllCases$LoanPurpose)
LoanPurposeCount <- length(LoanPurpose)
Divisions <- unique(AllCases$Division[!AllCases$Division == "Unknown"])
SolicitorTypes <- unique(AllCases$SolicitorType)

################################################################################
