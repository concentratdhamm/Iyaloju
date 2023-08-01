library(REDCapR)
library(ggplot2)
library(gtsummary)
library(tidyverse)
library(lubridate)
library(kableExtra)
library(readxl)
library(writexl)
library(sjmisc)
library(stringr)
library(openxlsx)

Pregnant_register <- read_xlsx("pregant_register_-_all_versions.xlsx", guess_max = 20000)

Pregnant_register <- rename(Pregnant_register, consent = `Do you agree to participate in this study?`, name = `Name of Interviewer`, lga = `Local Government Area`, relati = `Relationship with the pregnant woman`, pname = `Name of the pregnant woman`, lgcd = `LG/LCDA CODE/PATIENT NO`, reg = `Registration Date`, linkf = `link facility`, hospid = `Hospital ID`, age = `Age of the respondent`, marital = `Marital Status`, phone = `Telephone number`, phone2 = `Husband Phone number`, phone2_001 = `Next of Kin Phone number`, nkin = `Name of Next of Kin`, add = `Address of the respondent`, edu = `What is your education level`, occu = `What is your occupation ?`, q101 = `q101. Last menstrual Period(MLP)`, q102 = `q102.Number of Pregnancies`, q103 = `q103.Number of children Alive`, q104 = `q104.Number of Miscarriages`, q105 = `q105. Contraceptive Use`, q106 = `q106. Registered for ANC?`, q106b = `q106b. If yes, where (Name of the Facility/Clinic)`, q107 = `q107.Current Complaints/Diagnosis`, q108 = `q108.Pre-existing Medical Conditions (If any, please mention)`, q109 = `q109.History of Drug Allergy (Are you allergic to any drug?)`, geopoint_lat = `_Record your Location_latitude`, geopoint_long = `_Record your Location_longitude`)

Pregnant_register <- Pregnant_register %>%
  dplyr::select(consent, name, lga, relati, pname, lgcd, reg, linkf, hospid, age, marital, phone, phone2, phone2_001, nkin, add, edu, occu, q101, q102, q103, q104, q105, q106, q106b, q107, q108, q109, geopoint_lat, geopoint_long) %>%
  dplyr::filter(!is.na(name) & (pname != "Test") & (pname != "B"))

#Convert Categorical to Raw
for (row in 1:nrow(Pregnant_register)) {
  if(Pregnant_register[row, "consent"] == "Yes"){
    Pregnant_register[row, "consent"] <- "1"
  }else{
    if(Pregnant_register[row, "consent"] == "No"){
      Pregnant_register[row, "consent"] <- "0"
    }
  }
}

for (row in 1:nrow(Pregnant_register)) {
  if(Pregnant_register[row, "name"] == "A"){
    Pregnant_register[row, "name"] <- "1"
  }else{
    if(Pregnant_register[row, "name"] == "B"){
      Pregnant_register[row, "name"] <- "2"
    }
  }
}

for (row in 1:nrow(Pregnant_register)) {
  Pregnant_register[row, "lga"] <- 
  case_when(
    as.character(Pregnant_register[row, "lga"]) == "Agege"  ~ "1"
    ,as.character(Pregnant_register[row, "lga"]) == "Ajeromi" ~ "2"
    ,as.character(Pregnant_register[row, "lga"]) == "Amuwo-Odofin"  ~ "3"
    ,as.character(Pregnant_register[row, "lga"]) == "Apapa" ~ "4"
    ,as.character(Pregnant_register[row, "lga"]) == "Badagry Central"  ~ "5"
    ,as.character(Pregnant_register[row, "lga"]) == "Epe"  ~ "6"
    ,as.character(Pregnant_register[row, "lga"]) == "Eti-Osa"  ~ "7"
    ,as.character(Pregnant_register[row, "lga"]) == "Ibeju-Lekki"  ~ "8"
    ,as.character(Pregnant_register[row, "lga"]) == "Ifako Ijaye"  ~ "9"
    ,as.character(Pregnant_register[row, "lga"]) == "Ikeja"  ~ "10"
    ,as.character(Pregnant_register[row, "lga"]) == "Kosofe"  ~ "11"
    ,as.character(Pregnant_register[row, "lga"]) == "Lagos Mainland"  ~ "12"
    ,as.character(Pregnant_register[row, "lga"]) == "Mushin"  ~ "13"
    ,as.character(Pregnant_register[row, "lga"]) == "Ojo"  ~ "14"
    ,as.character(Pregnant_register[row, "lga"]) == "Oshodi Isolo"  ~ "15"
    ,as.character(Pregnant_register[row, "lga"]) == "Shomolu"  ~ "16"
    ,as.character(Pregnant_register[row, "lga"]) == "Surulere"  ~ "17"
  )
}

for (row in 1:nrow(Pregnant_register)) {
  Pregnant_register[row, "relati"] <- 
    case_when(
      tolower(as.character(Pregnant_register[row, "relati"])) == "husband"  ~ "1"
      ,tolower(as.character(Pregnant_register[row, "relati"])) == "father" ~ "2"
      ,tolower(as.character(Pregnant_register[row, "relati"])) == "brother"  ~ "3"
      ,tolower(as.character(Pregnant_register[row, "relati"])) == "other" ~ "4"
      ,tolower(as.character(Pregnant_register[row, "relati"])) == "self"  ~ "5"
    )
}

for (row in 1:nrow(Pregnant_register)) {
  Pregnant_register[row, "marital"] <- 
    case_when(
      as.character(Pregnant_register[row, "marital"]) == "Single"  ~ "1"
      ,as.character(Pregnant_register[row, "marital"]) == "Married" ~ "2"
      ,as.character(Pregnant_register[row, "marital"]) == "Divorced"  ~ "3"
      ,as.character(Pregnant_register[row, "marital"]) == "Separated" ~ "4"
    )
}

for (row in 1:nrow(Pregnant_register)) {
  Pregnant_register[row, "edu"] <- 
    case_when(
      as.character(Pregnant_register[row, "edu"]) == "No formal education"  ~ "1"
      ,as.character(Pregnant_register[row, "edu"]) == "Completed Primary" ~ "2"
      ,as.character(Pregnant_register[row, "edu"]) == "Completed Secondary"  ~ "3"
      ,as.character(Pregnant_register[row, "edu"]) == "Completed Tertiary" ~ "4"
      ,as.character(Pregnant_register[row, "edu"]) == "Postgraduate" ~ "5"
    )
}

for (row in 1:nrow(Pregnant_register)) {
  Pregnant_register[row, "occu"] <- 
    case_when(
      as.character(Pregnant_register[row, "occu"]) == "Unemployed"  ~ "1"
      ,as.character(Pregnant_register[row, "occu"]) == "Self-Employed" ~ "2"
      ,as.character(Pregnant_register[row, "occu"]) == "Government Employment"  ~ "3"
      ,as.character(Pregnant_register[row, "occu"]) == "Private Employment" ~ "4"
    )
}

for (row in 1:nrow(Pregnant_register)) {
  if(Pregnant_register[row, "q105"] == "Yes"){
    Pregnant_register[row, "q105"] <- "1"
  }else{
    if(Pregnant_register[row, "q105"] == "No"){
      Pregnant_register[row, "q105"] <- "0"
    }
  }
}

for (row in 1:nrow(Pregnant_register)) {
  if(Pregnant_register[row, "q106"] == "Yes"){
    Pregnant_register[row, "q106"] <- "1"
  }else{
    if(Pregnant_register[row, "q106"] == "No"){
      Pregnant_register[row, "q106"] <- "0"
    }
  }
}

for (row in 1:nrow(Pregnant_register)) {
  if(Pregnant_register[row, "q109"] == "Yes"){
    Pregnant_register[row, "q109"] <- "1"
  }else{
    if(Pregnant_register[row, "q109"] == "No"){
      Pregnant_register[row, "q109"] <- "0"
    }
  }
}

Pregnant_register <- Pregnant_register %>%
  add_column(reg_2 = format(as.POSIXct(Pregnant_register$reg), format = "%Y-%m-%d")
             , .after = "reg")


drop_columns <- c('reg')
Pregnant_register <- Pregnant_register %>%
  select(-one_of(drop_columns))

Pregnant_register <- rename(Pregnant_register, reg = reg_2)

write_xlsx(Pregnant_register, "Pregnant_register.xlsx")