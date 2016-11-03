##Descriptive statistics of University path...

##############
## Preparation
##############
rm(list = ls()) #Remove all objects in the environment
gc() ##Free up the memory

if(!exists("original_path"))
  original_path <- getwd()
setwd(file.path("CollegeStudentsStats"))

options(scipen=999)
dir.create(file.path("output", Sys.Date()), showWarnings = FALSE)

##libraries
options(java.parameters = "-Xmx2g")
library(readxl)
library(data.table)
library(dplyr)
library(stringdist)
library(XLConnect)

##
insertRow <- function(existingDF, newrow, r) {
  existingDF[seq(r+1,nrow(existingDF)+1),] <- existingDF[seq(r,nrow(existingDF)),]
  existingDF[r,] <- newrow
  existingDF
}

##Find the file...
list.files("input")
n              <- readline(prompt="Enter the folder's name: ") ##list.files("input")[1]
totalFormFName <- list.files(file.path("input", n), full=T)[grepl("科系職涯地圖", list.files(file.path("input", n)))]

##Standard data
MatchTable    <- read_excel(totalFormFName, sheet = "學系相關資料")
OldEducation  <- read_excel(totalFormFName, sheet = "升學藍圖")
OldEmployment <- read_excel(totalFormFName, sheet = "就業藍圖")

##Raw data
EducationFName      <- list.files(file.path("input", n), full=T)[grepl("升學", list.files(file.path("input", n)))]
EmploymentFName     <- list.files(file.path("input", n), full=T)[grepl("就業", list.files(file.path("input", n)))]
wb                  <- loadWorkbook(EducationFName)
RawEducation        <- readWorksheet(wb, sheet = 1, header = FALSE)
RawEducation[2, ]
names(RawEducation) <- RawEducation[2, ]
RawEducation        <- RawEducation[-c(1:2), ]
wb                  <- loadWorkbook(EmploymentFName)
RawEmployment       <- readWorksheet(wb, sheet = 1, header = TRUE)

IndustryRenew <- read.csv("產業新類別vs舊類別.csv", stringsAsFactors=F)
JobRenew      <- read.csv("職務小類轉換表.csv", stringsAsFactors=F)
##CollegeNameT  <- read.csv("學校名稱正規化表格.csv",stringsAsFactors=F)

##Remove redundant curriculum vitaes...
RawEmployment$履歷編號 <- NULL
RawEmployment <- unique(setDT(RawEmployment))

Employment <- RawEmployment[, c("會員編號","學校代碼", "學校名稱", "科系名稱", "科系類別代號", 
                           "科系類別名稱", "產業小類代碼", "產業小類名稱", 
                           "職務小類代碼", "職務小類名稱", "產業小類代碼1",
                           "產業小類名稱1", "職務小類代碼1", "職務小類名稱1",
                           "產業小類代碼2", "產業小類名稱2", "職務小類代碼2",
                           "職務小類名稱2", "產業小類代碼3", "產業小類名稱3", 
                           "職務小類代碼3", "職務小類名稱3"), with=FALSE ]

Employment <- Employment[!grepl("[0-9]", 學校名稱)]
Employment <- Employment[!grepl("學分班", 科系名稱)]

##College names correction...
Employment <- Employment[學校名稱 %in% MatchTable$學校名稱]
Employment$SchoolName <- Employment$學校名稱
#for(x in unique(Employment$學校名稱)){
#  tmp <- ifelse(toString(rev(sort(CollegeNameT$對應表[CollegeNameT$trim後原始 == x ]))[1])!="" & toString(rev(sort(CollegeNameT$對應表[CollegeNameT$trim後原始== x]))[1])!="NA"
#                , rev(sort(CollegeNameT$對應表[CollegeNameT$trim後原始== x]))[1]
#                , x)
#  if(x!=tmp)
#    Employment$SchoolName[which(Employment$學校名稱==x)] <- tmp
#}
#for(x in unique(Employment$SchoolName)){
#  ##Fuzzy matching : find the Min dist word
#  Employment$SchoolName[which(Employment$SchoolName==x)] <- unique(MatchTable$學校名稱)[which.min(stringdist(x, unique(MatchTable$學校名稱) ,method='jw'))][1]
#}

##Check
#Employment[學校名稱!=SchoolName, .(學校名稱, SchoolName)] %>% unique

##Department names corrections...
Employment$Department <- Employment$科系名稱

UniMatchT        <- MatchTable[, c("學校名稱", names(MatchTable)[11])] %>% unique
names(UniMatchT) <- c("學校名稱", "科系名稱")
UniMatchT        <- UniMatchT[which(UniMatchT$科系名稱!="NULL"), ]
EmploymentMatchT <- Employment[, .(SchoolName, Department)] %>% unique

#check <- data.frame("Original"=character(), "Match"=character(), stringsAsFactors = FALSE)
for(i in 1:nrow(EmploymentMatchT)){
  ##Fuzzy matching : find the Min dist word
  tmp <- UniMatchT$科系名稱[UniMatchT$學校名稱==Employment$SchoolName[i]][which.min(stringdist(EmploymentMatchT$Department[i], UniMatchT$科系名稱[UniMatchT$學校名稱==Employment$SchoolName[i]] ,method='jw'))][1]
  
  if(EmploymentMatchT$Department[i]!=tmp){
    Employment$Department[which(Employment$SchoolName==EmploymentMatchT$SchoolName[i] & Employment$Department==EmploymentMatchT$Department[i])] <- tmp
    #check <- insertRow(check, c(EmploymentMatchT$Department[i], tmp), 1)
  } 
  cat("\r Department Correction : ", format(round(i/nrow(EmploymentMatchT)*100, 2), nsmall=2), " %")
}




##########
## Module
##########
DStatModule <- function(DT, school, department, target){
  ##Try new method...
  ##Not using eval..., just rename the col in the function...
  ##More readable?
  setDF(DT)
  names(DT)[which(names(DT)==school)]     <- "School"
  names(DT)[which(names(DT)==department)] <- "Department"
  names(DT)[which(names(DT)==target)]     <- "Target"

  DStatDT <- DT[ , .N, by= .(School, Department, Target)]
  DStatDT <- DStatDT[order(School, Department, -N)]
  #DStatDT[, Percentage:=N/sum(N), by = .(School, Department)]
  
  return(DStatDT)
}
#cbnName="名稱(一)"
UpdateModule <- function(ODT, NDT, cbnName){
  setDT(ODT)
  #school, department, index of cbnName and the next one
}


####################
## Start : Education
####################

#####################
## Start : Employment
#####################