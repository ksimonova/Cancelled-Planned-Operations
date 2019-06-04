

library('tidyverse')
library('reshape2')
library('xlsx')
library('lubridate')
library('stringr')
library("RcppBDT")


#### 1. Read in database files----

# 1.1 Read in Specialty Level database

dat <- read_csv('//stats/WaitingTimes/Cancellations/Database/Specialty Level/Specialty Level Database.csv')


# 1.2 Read in All Specialty Level database

all <- read_csv('//stats/WaitingTimes/Cancellations/Database/All Specialties/All Specialties Database.csv')


# 1.3 Read in Discovery file

dis <- read.csv("//stats/WaitingTimes/Cancellations/Discovery/cancellations.csv", check.names = FALSE)


# 1.4 Read in Location Lookup
loclookup <- read.csv(paste0("Lookups/location.csv"),
                stringsAsFactors = F)


# 1.5 Read in Health Board Lookup
hblookup <- read.csv("Lookups/geography_codes_and_labels_hb2014.csv",
                        stringsAsFactors = F)

# 1.6 Read in Specialty descriptions
spec <- read.csv(paste0("Lookups/spec_lookup.csv"),
                 check.names = F, stringsAsFactors = F)

#1.7 Read in board codes file for Discovery
bcode <- read_csv("Lookups/Health Board Area 2014 Lookup.csv")



#### 2. Read in new month's data ----

# 2.1 Get current month's publication folder and read in submissions


#Get new publication month's folder
date <- Sys.Date() %m+% months(1)
month <- month(date)
year <- year(date)

folder <-getNthDayOfWeek(first, Tue, month, year)

folderyear <- year(folder)
foldermonth <- formatC(month(folder), width=2, flag="0")
folderday <- formatC(day(folder), width=2, flag="0")

#Set path of where submissions are stored
path <- paste0('//stats/WaitingTimes/Cancellations/Publications/',folderyear,foldermonth,folderday,'/FilesFromCustomer/')  

path <- paste0('//stats/WaitingTimes/Cancellations/Publications/20190604/FilesFromCustomer/')  




#Get all the file names with .xls or .xlsx extension in above folder
file.names <- dir(path, pattern =".xls") 


#Get current data's month
monthtext <- format(Sys.Date() %m-% months(1),"01/%m/%Y")




#Create Current month data frame
Currentmonth <- NULL

#Cycle through all excel files in submissions folder and append on to Current month data frame
for(i in 1:length(file.names)) {
      
      #Read in file
      file <- read.xlsx(paste0(path,file.names[i],sep=""),sheetName="Return")
  
      #Change names of columns if needed
      names(file)[names(file) == 'Hospital.Code'] <- 'Hospital'
      names(file)[names(file) == 'Specialty.Code'] <- 'Specialty'
      
      #Delete any rows where Hospital or Specialty is blank
      file <- file[!is.na(file[,1]),]  
      file <- file[!is.na(file[,2]),]  
      
      #Get name of board and populate vector with it
      loc <- str_locate(file.names[i], "-") 
      Board <-rep(substr(file.names[i],1,loc[1]-1),length(file$Hospital))
      
      #Fill vector with month
      Month <- rep(monthtext,length(file$Hospital))
      
      #Create file for current board's data        
      file <- cbind(Month,Board,file)
      
      #Select only the first 10 columns
      file <- file[,1:10]
      
      #Append board to current month data frame
      Currentmonth <- rbind(Currentmonth, file)
}


# 2.2 Clean current month's data


#Change Board column to character
Currentmonth$Board <- as.character(Currentmonth$Board)

#Rename select Boards
Currentmonth <- Currentmonth %>%
                  mutate(Board = ifelse(Board == "GoldenJubilee"| Board=="GJ" | Board=="GJNH",
                                              "Golden Jubilee",
                                        ifelse(Board == "A&A", "Ayrshire & Arran",
                                          ifelse(Board == "D&G", "Dumfries & Galloway",
                                            ifelse(Board == "ForthValley" | Board=="FV", "Forth Valley",
                                              ifelse(Board == "GG&C", "Greater Glasgow & Clyde",
                                                ifelse(Board == "WesternIsles" | Board=="WI", "Western Isles",
                                                  Board))))))) %>%
                   mutate(Board = ifelse(Board != "Golden Jubilee",
                                            paste0("NHS ",Board),Board))


#Replace NAs and .. with 0
Currentmonth[is.na(Currentmonth)] <-0
Currentmonth[Currentmonth==".."] <-0




#Rename Columns
Currentmonth <- Currentmonth %>% 
                  rename(
                      `Total Ops` = `Total.no.of.scheduled.elective.operations.in.theatre.systems`,   
                      `Total Cancelled` = `Total.no.of.scheduled.elective.cancellations.in..theatre.systems`,
                      `Clinical reason` = `Cancellation.based.on.clinical.reason.by.hospital`,
                      `Non-clinical/Capacity reason` = `Cancellation.based.on.capacity.or.non.clinical.reason.by.hospital`,
                      `Cancelled by patient` = `Cancelled.by.Patient`,
                      `Other` = `Other.reason`)


#Convert factors to character variables
Currentmonth <- Currentmonth %>%
                    mutate(Hospital = as.character(Hospital),
                           Specialty = as.character(Specialty),
                           Month = as.character(Month))


# Format specialty codes 
# Trim white space, remove full stops, aggregate to 2 digit specialty 
Currentmonth <- Currentmonth %>% 
                    mutate(Specialty = trimws(Specialty),
                           Specialty = gsub('.', '', Specialty, fixed = TRUE),
                           Specialty = substr(Specialty, 1, 2)) %>%
                    group_by(Month, Board, Hospital, Specialty) %>%
                    summarise_all(list(sum))  %>% # aggregate up to specialty          
                    ungroup()




# 2.3 Checks for submisssions

## Check 1 - Cancellation reasons should sum to total cancellations for each row

check1 <- sum(rowSums(Currentmonth[,c("Clinical reason","Cancelled by patient", "Non-clinical/Capacity reason","Other")]) - Currentmonth[,"Total Cancelled"])
if ( check1 > 0) stop("Cancellation reasons do not equal total cancellations")

## Open 'error1' in environment to see source of errors for check 1

error1 <- Currentmonth %>%
  filter(Currentmonth[,"Total Cancelled"]!=rowSums(Currentmonth[,c("Clinical reason","Cancelled by patient",
                                                                   "Non-clinical/Capacity reason","Other")]))
#Output to csv



## Check 2 - Total cancellations should be less than total operations for each row

check2 <- sum(Currentmonth[,"Total Ops"] < Currentmonth[,"Total Cancelled"])
if (check2 > 0) stop("Total cancellations exceed total operations")

## Open 'error2' in Enivronment to see source of errors for check 2

error2 <- Currentmonth %>%
  filter(Currentmonth[,"Total Ops"] < Currentmonth[,"Total Cancelled"])







#### 3. Output Specialty level database file ----

#Format month as a date
#Currentmonth$Month <- as.Date(Currentmonth$Month, format = "%d/%m/%Y")

#Bind current month's data to specialty level database
specdat <- rbind(dat, Currentmonth)

#write to spec database
write.csv(specdat,'//stats/WaitingTimes/Cancellations/Database/Specialty Level/Specialty Level Database.csv', row.names = FALSE)








#### 4. Output Publication data ----

#Aggregate to all specialty level
pub <- Currentmonth %>%
            select(-Specialty) %>%
            group_by(Month, Board, Hospital) %>%
            summarise_all(sum) %>%
            ungroup

#Set path for Output folder within latest publication folder
path <- paste0('//stats/WaitingTimes/Cancellations/Publications/',folderyear,foldermonth,folderday,'/Output/')  


#write to csv
write.csv(pub, paste0(path,'Publication Data.csv'), row.names = FALSE)









#### 5. Output All Specialties Level for summary ----

# 5.1 Update All Specialties database with latest data

#Convert Month to date 
all$Month <- as.Date(all$Month , format = "%d/%m/%Y")
pub$Month <- as.Date(pub$Month, format = "%d/%m/%Y")

#Append latest month to All Specialties database
all <- rbind(all, pub)


#Output to All Specialties Level Database
write.csv(all, '//stats/WaitingTimes/Cancellations/Database/All Specialties/All Specialties Database.csv', row.names = FALSE)





# 5.2 Output data for summary


#Aggergate to Board Level
summ <- all %>%
          select(-Hospital) %>%
          group_by(Month, Board) %>%
          summarise_all(sum) 

#Create a Scotland Level total
summ <- summ %>%
          ungroup %>%
          select(-Board) %>%
          group_by(Month) %>%
          summarise_all(sum) %>%
          mutate(Board = "NHS Scotland") %>%
          bind_rows(summ)

#Change to wide format
summ <- summ %>%
          gather(Indicator, Num, `Total Ops`:Other) %>%
          spread(Month, Num)


#write.csv
write.csv(summ, 'rmarkdown script/Data/Cancellation lookup.csv', row.names = FALSE)










#### 6. NHS Performs ----

#Create Board Level Data
board <- pub %>%
          ungroup %>%
          select(-Hospital) %>%
          group_by(Month,Board) %>%
          summarise_all(sum) %>%
          rename("Location" = "Board") %>%
          ungroup

#Creat Scotland Level Data
scot <- pub %>%
          ungroup %>%
          select(-Board, - Hospital) %>%
          group_by(Month) %>%
          summarise_all(sum) %>%
          mutate("Location" = "Scotland") %>%
          ungroup


          
pub <- pub %>%
        select(-Board) %>%
        rename("Location" = "Hospital") %>%
        left_join(loclookup) %>%
        select(Month:Locname, -Location) %>%
        rename("Location" = "Locname") %>%
        ungroup

pub <- rbind(pub, scot, board)


#Calculate percentages
pub <- pub %>%
          select(Month, Location, `Total Ops`, `Total Cancelled`, `Non-clinical/Capacity reason`) %>%
          mutate(Indicator1 = round(`Total Cancelled`/`Total Ops`, digits = 3),
                 Indicator2 = round(`Non-clinical/Capacity reason`/`Total Ops`, digits=3)) %>%
          mutate(Location = ifelse(Location == "New Dumfries & Galloway Royal Infirmary"
                                                , "Dumfries & Galloway Royal Infirmary",
                            ifelse(Location == "West Glasgow/Gartnavel General",
                                              "Gartnavel General Hospital", 
                            ifelse(Location == "Monklands District General Hospital",
                                                  "Monklands Hospital", 
                            ifelse(Location == "St John's Hospital",
                                                "St John's Hospital at Howden",
                            ifelse(Location == "Royal Infirmary of Edinburgh at Little France",
                                                  "Royal Infirmary of Edinburgh",
                            ifelse(Location == "Lorn & Islands Hospital",
                                                  "Lorn & Islands District General Hospital", 
                            ifelse(Location == "Stobhill Hospital",
                                              "New Stobhill Hospital",
                            ifelse(Location == "Victoria Infirmary",
                                                "New Victoria Hospital",
                            ifelse(Location == "Royal Hospital for Children",
                                               "Royal Hospital for Children Glasgow",
                            ifelse(Location == "Royal Hospital for Sick Children (Edinburgh)",
                                              "Royal Hospital for Sick Children Edinburgh",Location))))))))))) %>%
          filter(!(Location %in% c("Aberdeen Maternity Hospital",
                                   "Golden Jubilee",
                                   "Mountainhall Treatment Centre",
                                   "Princess Alexandra Eye Pavilion")))

 
ind1 <- pub %>%
          mutate(Topic = "Cancelled Operations",
                 Indicator = "Percentage of total number of cancellations",
                 Month = format(Month, format = "%d/%m/%Y")) %>%
          rename(Time_period = Month,
                 Data_value = Indicator1) %>%
          select(Topic, Indicator, Time_period, Location, Data_value)

ind2 <- pub %>%
            mutate(Topic = "Cancelled Operations",
                   Indicator = "Percentage of total number of cancellations",
                   Month = format(Month, format = "%d/%m/%Y")) %>%
            rename(Time_period = Month,
                   Data_value = Indicator2) %>%
            select(Topic, Indicator, Time_period, Location, Data_value)


#Output to latest publication folder
write.csv(ind1, paste0(path,'All cancellations ', month.abb[month], sprintf('%02d', year %% 100),'.csv'), row.names = FALSE, quote  = FALSE)
write.csv(ind2, paste0(path,'Cancellations for non-clinical reasons  ', month.abb[month], sprintf('%02d', year %% 100), '.csv'), row.names = FALSE, quote = FALSE)









#### 7. Discovery ----



#Match specialty codes to specialty descriptions
Currentmonth <- 
  Currentmonth %>%
  rename(`Specialty Code` = Specialty) %>%  # matching variables have same name
  left_join(spec)

rm(spec)

# Code missing specialty information as unknown 
Currentmonth <-
  Currentmonth %>%
  mutate(`Specialty Code` = ifelse(is.na(`Specialty Grouping`), 'Unknown', `Specialty Code`),
         `Specialty Grouping` = ifelse(is.na(`Specialty Grouping`), 'Unknown', `Specialty Grouping`),
         `Specialty Description` = ifelse(is.na(`Specialty Description`), 'Unknown', `Specialty Description`))


Currentmonth <- Currentmonth %>% 
  aggregate(.~`Month`+`Board`+`Hospital`+`Specialty Code`+`Specialty Grouping`+`Specialty Description`,data=.,sum)



# Add row for every location for every specialty 
Currentmonth <- 
  Currentmonth %>%
  ungroup() %>%
  complete(nesting(Month, Board, Hospital), 
           nesting(`Specialty Code`, `Specialty Grouping`, `Specialty Description`)) %>%
  mutate_at(vars(`Total Ops`:Other),
            funs(ifelse(is.na(.), 0, .)))



# Add All Specialties rows 
Currentmonth <- 
  Currentmonth %>%
  select(-`Specialty Grouping`, -`Specialty Description`, -`Specialty Code`) %>%
  group_by(Hospital, Board, Month) %>%
  summarise_all(funs(sum)) %>%
  mutate(`Specialty Grouping` = 'All Specialties',
         `Specialty Description` = 'All Specialties',
         `Specialty Code` = 'All') %>%
  bind_rows(Currentmonth)




# Match on location names 
Currentmonth <- 
  Currentmonth %>%
  rename(Location = Hospital) %>%    # matching variables have same name
  left_join(loclookup) %>%
  ungroup() %>%
  select(-c(Add1:filler))  # remove unnecessary variables

# Create Board code


Currentmonth <- Currentmonth %>%
                mutate(HealthBoardArea2014Name = ifelse(substr(Board,1,3)=="NHS",substr(Board,5,40),Board))

Currentmonth$HealthBoardArea2014Name <- gsub("&","and",Currentmonth$HealthBoardArea2014Name)

Currentmonth <- Currentmonth %>%
                left_join(bcode) %>%
                select(-HealthBoardArea2014Name,-NRSHealthBoardAreaName) %>%
                rename(`Board Code` = HealthBoardArea2014Code )


# Transpose Reasons columns to rows 
Currentmonth <- 
  Currentmonth %>%
  gather(Indicator, Count, `Clinical reason`:Other) %>%
  rename(`Location Code` = Location) %>%
  select(`Board Code`,Board,Month,`Total Ops`,`Total Cancelled`,`Specialty Grouping`,`Specialty Description`,
                `Specialty Code`,	`Location Code`,Locname,Indicator,Count)


Currentmonth$Month <- format(Currentmonth$Month, format = "%d/%m/%Y")


#Add Current months data to previous months
Database<- rbind(dis,Currentmonth)



#Write to database
write.csv(Database, "//stats/WaitingTimes/Cancellations/Discovery/cancellations.csv", row.names = FALSE)





#### 8. Create Summary ----

outpath <- paste0(path,folderyear,"-",foldermonth,"-",folderday,"-Cancellations-Summary.docx")

#outpath <- paste0(path,"2019-01-03-Cancellations-Summary2.docx")

rmarkdown::render('rmarkdown script/Cancellation_Summary.Rmd',
                  output_file = outpath)


#### 9. Open Data ----

#Amend lookup file to only include used columns
hblookup <- hblookup %>%
              select(HB2014,HB2014Name)

#Create base open data file, amending file names and formats
open <- all %>%
          mutate(Month = format(Month, "%Y%m"),
                 Board = gsub("&", "and", Board)) %>%
          left_join(hblookup, by = c("Board" = "HB2014Name")) %>%
          rename(TotalOperations = `Total Ops`,
                  TotalCancelled = `Total Cancelled`,
                    CancelledByPatientReason = `Cancelled by patient`,
                      ClinicalReason = `Clinical reason`,
                        NonClinicalCapacityReason = `Non-clinical/Capacity reason`,
                          OtherReason = `Other`,
                            HBT2014 = HB2014) %>%
          select(-Board) 
          

#Create Scotland level file
scot <- open %>%
          select(-Hospital, -HBT2014) %>%
          group_by(Month) %>%
          summarise_all(sum) %>%
          mutate(Country = "S92000003") %>%
          select(Month, Country, TotalOperations:OtherReason)


#Create Board level file
board <- open %>% 
            select(-Hospital) %>%
            group_by(Month, HBT2014) %>%
            summarise_all(sum) %>%
            select(Month, HBT2014, TotalOperations:OtherReason)

#Create hospital level file
hospital <- open %>%
              select(-HBT2014)
      
#Output files

#Calculate dates for file name
month2 <- month(Sys.Date() %m-% months(1))
filemonth <- month.name[month2]
year2 <- year(Sys.Date() %m-% months(1))

#Write files to csv
write.csv(scot, paste0("//stats/WaitingTimes/Cancellations/Open Data/Output/Cancellations_Scotland_",
                              filemonth, "_", year2, ".csv"), row.names= FALSE)
write.csv(board, paste0("//stats/WaitingTimes/Cancellations/Open Data/Output/Cancellations_by_Board_",
                            filemonth, "_", year2, ".csv"), row.names = FALSE)
write.csv(hospital, paste0("//stats/WaitingTimes/Cancellations/Open Data/Output/Cancellations_by_Hospital_",
                          filemonth, "_", year2, ".csv"), row.names = FALSE)
