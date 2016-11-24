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
options(java.parameters = "-Xmx4g")
library(readxl)
library(data.table)
library(dplyr)
library(stringdist)
library(XLConnect)
library(magrittr)
library(stringi)

## Check encoding
EncodingCheck <- function(tmp){
  if(is.list(tmp)){
    for(i in 1:ncol(tmp)){
      tryCatch({
        print(Encoding(tmp[[i]]) %>% unique)
      }, error = function(e) {
        ##conditionMessage(e)
      })
    }
  }else{
    for(i in 1:ncol(tmp)){
      tryCatch({
        print(Encoding(tmp[,i]) %>% unique)
      }, error = function(e) {
        ##conditionMessage(e)
      })
    }
  }
  
}
##
#insertRow <- function(existingDF, newrow, r) {
#  existingDF[seq(r+1,nrow(existingDF)+1),] <- existingDF[seq(r,nrow(existingDF)),]
#  existingDF[r,] <- newrow
#  existingDF
#}

##Find the file...
list.files("input")
n              <- readline(prompt="Enter the folder's name: ") ##list.files("input")[1]
totalFormFName <- list.files(file.path("input", n), full=T)[grepl("科系職涯地圖", list.files(file.path("input", n))) & !grepl("~$", list.files(file.path("input", n)), fixed = TRUE)]

##Standard data
wb            <- loadWorkbook(totalFormFName)
MatchTable    <- readWorksheet(wb, sheet = "學系相關資料", header = TRUE)
OldEducation  <- readWorksheet(wb, sheet = "升學藍圖", header = TRUE)
OldEmployment <- readWorksheet(wb, sheet = "就業藍圖", header = TRUE)
#MatchTable    <- read_excel(totalFormFName, sheet = "學系相關資料")
#OldEducation  <- read_excel(totalFormFName, sheet = "升學藍圖")
#OldEmployment <- read_excel(totalFormFName, sheet = "就業藍圖")

## Encoding problem
EncodingCheck(OldEmployment)
setDT(OldEmployment)
OldEmployment[, names(OldEmployment) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
for(i in 1:ncol(OldEmployment)){
  OldEmployment[[i]] <- OldEmployment[[i]] %>% iconv("UTF-8")
}
EncodingCheck(OldEducation)
setDT(OldEducation)
OldEducation[, names(OldEducation) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
for(i in 1:ncol(OldEducation)){
  OldEducation[[i]] <- OldEducation[[i]] %>% iconv("UTF-8")
}

##Raw data
gc()
EducationFName      <- list.files(file.path("input", n), full=T)[grepl("升學", list.files(file.path("input", n)))]
EmploymentFName     <- list.files(file.path("input", n), full=T)[grepl("就業", list.files(file.path("input", n)))]
wb                  <- loadWorkbook(EducationFName)
RawEducation        <- readWorksheet(wb, sheet = 1, header = FALSE)
RawEducation[2, ]
names(RawEducation) <- RawEducation[2, ]
RawEducation        <- RawEducation[-c(1:2), ]
wb                  <- loadWorkbook(EmploymentFName)
RawEmployment       <- readWorksheet(wb, sheet = 1, header = TRUE)

EncodingCheck(RawEmployment)
setDT(RawEmployment)
RawEmployment[, names(RawEmployment) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
for(i in 1:ncol(RawEmployment)){
  RawEmployment[[i]] <- RawEmployment[[i]] %>% iconv("UTF-8")
}
EncodingCheck(RawEducation)
setDT(RawEducation)
RawEducation[, names(RawEducation) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
for(i in 1:ncol(RawEducation)){
  RawEducation[[i]] <- RawEducation[[i]] %>% iconv("UTF-8")
}

IndustryRenew <- read.csv("產業新類別vs舊類別.csv", stringsAsFactors=F)
JobRenew      <- read.csv("職務小類轉換表.csv", stringsAsFactors=F)
##CollegeNameT  <- read.csv("學校名稱正規化表格.csv",stringsAsFactors=F)
EncodingCheck(JobRenew)
EncodingCheck(IndustryRenew)
##
IndustryRenew  %>% names
for(i in 1:nrow(IndustryRenew)){
  #RawEmployment %>% names
  #OldEmployment %>% names
  
  RawEmployment$產業小類名稱[which(RawEmployment$產業小類名稱==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  RawEmployment$產業小類名稱[which(RawEmployment$產業小類名稱1==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  RawEmployment$產業小類名稱[which(RawEmployment$產業小類名稱2==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  RawEmployment$產業小類名稱[which(RawEmployment$產業小類名稱3==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  RawEmployment$產業小類名稱[which(RawEmployment$產業小類名稱4==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  
  OldEmployment$名稱.一.[which(OldEmployment$名稱.一.==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  OldEmployment$名稱.二.[which(OldEmployment$名稱.二.==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  OldEmployment$名稱.三.[which(OldEmployment$名稱.三.==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  OldEmployment$名稱.四.[which(OldEmployment$名稱.四.==IndustryRenew$X1111產業小類名稱[i])] <- IndustryRenew$X1111產業小類名稱New[i]
  cat("\r Industry Renew : " , format(round(i/nrow(IndustryRenew)*100), nsmall=2), "%")
}
JobRenew %>% names
for(i in 1:nrow(JobRenew)){
  #RawEmployment %>% names
  #OldEmployment %>% names
  
  RawEmployment$職務小類名稱[which(RawEmployment$職務小類名稱==JobRenew$old[i])] <- JobRenew$new[i]
  RawEmployment$職務小類名稱1[which(RawEmployment$職務小類名稱1==JobRenew$old[i])] <- JobRenew$new[i]
  RawEmployment$職務小類名稱2[which(RawEmployment$職務小類名稱2==JobRenew$old[i])] <- JobRenew$new[i]
  RawEmployment$職務小類名稱3[which(RawEmployment$職務小類名稱3==JobRenew$old[i])] <- JobRenew$new[i]
  RawEmployment$職務小類名稱4[which(RawEmployment$職務小類名稱4==JobRenew$old[i])] <- JobRenew$new[i]
  
  OldEmployment$名稱.一.[which(OldEmployment$名稱.一.==JobRenew$old[i])] <- JobRenew$new[i]
  OldEmployment$名稱.二.[which(OldEmployment$名稱.二.==JobRenew$old[i])] <- JobRenew$new[i]
  OldEmployment$名稱.三.[which(OldEmployment$名稱.三.==JobRenew$old[i])] <- JobRenew$new[i]
  OldEmployment$名稱.四.[which(OldEmployment$名稱.四.==JobRenew$old[i])] <- JobRenew$new[i]
  cat("\r Job Renew : " , format(round(i/nrow(JobRenew)*100), nsmall=2), "%")
}
##氣象學類 -> 大氣科學學類
OldEducation$名稱[OldEducation$名稱=="氣象學類"] <- "大氣科學學類"

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
Employment <- Employment[學校名稱 %in% (MatchTable$學校名稱 %>% iconv("UTF-8") %>% unique)]
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
Employment <- Employment[!is.na(科系名稱)]

UniMatchT <- MatchTable[, c("學校名稱", names(MatchTable)[11])] %>% unique
EncodingCheck(UniMatchT)
setDT(UniMatchT)
UniMatchT[, names(UniMatchT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
for(i in 1:ncol(UniMatchT)){
  UniMatchT[[i]] <- UniMatchT[[i]] %>% iconv("UTF-8")
}

names(UniMatchT) <- c("學校名稱", "科系名稱")
UniMatchT        <- UniMatchT[which(UniMatchT$科系名稱!="NULL"), ]
EmploymentMatchT <- Employment[, .(SchoolName, Department)] %>% unique

names(RawEducation)[6] <- "SchoolName"
names(RawEducation)[9] <- "Department"
Education <- RawEducation[就學地區2!=""] 
EducationMatchT <- Education[, .(SchoolName, Department)] %>% unique

##Employment
#check <- data.frame("Original"=character(), "Match"=character(), stringsAsFactors = FALSE)
for(i in 1:nrow(EmploymentMatchT)){
  ##Fuzzy matching : find the Min dist word
  tmp <- UniMatchT$科系名稱[UniMatchT$學校名稱==EmploymentMatchT$SchoolName[i]][which.min(stringdist(EmploymentMatchT$Department[i], UniMatchT$科系名稱[UniMatchT$學校名稱==EmploymentMatchT$SchoolName[i]] ,method='jw'))][1]
  
  if(EmploymentMatchT$Department[i]!=tmp){
    Employment$Department[which(Employment$SchoolName==EmploymentMatchT$SchoolName[i] & Employment$Department==EmploymentMatchT$Department[i])] <- tmp
    #check <- insertRow(check, c(EmploymentMatchT$Department[i], tmp), 1)
  } 
  cat("\r Department Correction Part.1: ", format(round(i/nrow(EmploymentMatchT)*100, 2), nsmall=2), " %")
}
## School
SchoolNameT       <- EducationMatchT$SchoolName %>% unique
CorrectSchoolName <- UniMatchT$學校名稱 %>% unique 
for(i in 1:length(SchoolNameT)){
  tmp <- CorrectSchoolName[which.min(stringdist(gsub("大學", "", gsub("專科", "", gsub("管理", "", SchoolNameT[i]))), UniMatchT$學校名稱 %>% unique ,method='jw'))][1]
  
  if(SchoolNameT[i]!=tmp){
    Education$SchoolName[which(Education$SchoolName==SchoolNameT[i])] <- tmp
    #check <- insertRow(check, c(EmploymentMatchT$Department[i], tmp), 1)
    EducationMatchT$SchoolName[which(EducationMatchT$SchoolName==SchoolNameT[i])] <- tmp
  } 
}
## Department
for(i in 1:nrow(EducationMatchT)){
  ##Fuzzy matching : find the Min dist word
  tmp <- UniMatchT$科系名稱[UniMatchT$學校名稱==EducationMatchT$SchoolName[i]][which.min(stringdist(EducationMatchT$Department[i], UniMatchT$科系名稱[UniMatchT$學校名稱==EducationMatchT$SchoolName[i]] ,method='jw'))][1]
  
  if(EducationMatchT$Department[i]!=tmp){
    Education$Department[which(Education$SchoolName==EducationMatchT$SchoolName[i] & Education$Department==EducationMatchT$Department[i])] <- tmp
    #check <- insertRow(check, c(EducationMatchT$Department[i], tmp), 1)
  } 
  cat("\r Department Correction Part.2: ", format(round(i/nrow(EducationMatchT)*100, 2), nsmall=2), " %")
}

names(Education)[17] <- "晉升學校名稱"
names(Education)[18] <- "待除錯晉升學校名稱"
Education$晉升學校名稱[which(Education$就學地區=="" & Education$晉升學校名稱=="")] <- "國外地區大學"
Education$晉升學校名稱[which(Education$晉升學校名稱=="")] <- Education$待除錯晉升學校名稱[which(Education$晉升學校名稱=="")]
nrow(Education[is.na(Education$晉升學校名稱)])
Education <- Education[!is.na(Education$晉升學校名稱)]
## School
SchoolNameT <- Education$晉升學校名稱 %>% unique 
for(i in 1:length(SchoolNameT)){
  tmp <- CorrectSchoolName[which.min(stringdist(gsub("大學", "", gsub("專科", "", gsub("管理", "", SchoolNameT[i]))), UniMatchT$學校名稱 %>% unique ,method='jw'))][1]
  
  if(SchoolNameT[i]!=tmp){
    Education$晉升學校名稱[which(Education$晉升學校名稱==SchoolNameT[i])] <- tmp
  } 
}


##########
## Module
##########
##DT=Employment
##school="SchoolName"
##department="Department"
##target="職務小類名稱"
DStatModule <- function(DT, school, department, target){
  ##Try new method...
  ##Not using eval..., just rename the col in the function...
  ##More readable?
  setDT(DT)
  names(DT)[which(names(DT)==school)]     <- "School"
  names(DT)[which(names(DT)==department)] <- "Department"
  names(DT)[which(names(DT)==target)]     <- "Target"
  
  DStatDT <- DT[ , .N, by= .(School, Department, Target)]
  DStatDT <- DStatDT[order(School, Department, -N)]
  #DStatDT[, Percentage:=N/sum(N), by = .(School, Department)]
  
  ##Weired.. 
  names(DT)[which(names(DT)=="School")]     <- school
  names(DT)[which(names(DT)=="Department")] <- department
  names(DT)[which(names(DT)=="Target")]     <- target
  
  return(DStatDT)
}

#target="名稱(一)"
#ODT=OldEmployment
#NDT=Job1
#switchV="職務"   "學校" "產業" 
UpdateModule <- function(NDT, ODT, target, switchV=""){
  print(paste0(switchV, ": ", target))
  setDT(ODT)
  setDT(NDT)
  
  ## Encoding problem
  #ODT[, names(ODT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
  #ODT[, names(ODT) := lapply(.SD, function(x) {if (is.character(x)) x %>% iconv("UTF-8")})]
  #for(i in 1:ncol(ODT)){
  #  ODT[[i]] <- ODT[[i]] %>% iconv("UTF-8")
  #}
  #NDT[, names(NDT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
  #NDT[, names(NDT) := lapply(.SD, function(x) {if (is.character(x)) x %>% iconv("UTF-8")})]
  #for(i in 1:ncol(NDT)){
  #  NDT[[i]] <- NDT[[i]] %>% iconv("UTF-8", "UTF-8")
  #}
  
  names(ODT) <- gsub("[()（）]", ".", names(ODT))
  target     <- gsub("[()（）]", ".", target)
  
  names(ODT)[which(grepl("類別", names(ODT)) & grepl("0", names(ODT)))] <- "類別"
  if(switchV=="產業" | switchV=="學校"){
    switchV <- 1
  }else{
    switchV <- 0
  }
  
  SODT <- ODT[類別==switchV, c("學校名稱", "科系名稱", target, names(ODT) %>% .[which(.==target)+1]), with=F]
  NDT
  names(NDT) <- c("學校名稱", "科系名稱", target, names(ODT) %>% .[which(.==target)+1])
  
  print(" Merging New and old data...")
  TDT      <- rbind(NDT, SODT)
  #tmp_eval <- paste0(names(ODT) %>% .[which(.==target)+1], ":=sum(as.numeric(", names(ODT) %>% .[which(.==target)+1],"))")
  ##Can't use original weird name... 
  names(TDT)[4] <- "Freq"
  names(TDT)[3] <- "Item"
  
  #library(magrittr)
  #TDT$Freq  %>% extract(9997)
  TDT$Freq[which(TDT$Freq=="NULL")] <- 0
  TDT$Freq <- as.numeric(TDT$Freq)
  TDT$Item[which(TDT$Item=="NULL")] <- "No Data"
  TDT[, Freq:=sum(as.numeric(Freq), na.rm=T), by=c("學校名稱", "科系名稱", "Item")]
  
  TDT$UniKey <- 0
  TDT$UniKey[which(TDT$Item=="No Data")] <- which(TDT$Item=="No Data")
  
  TDT <- TDT %>% unique
  TDT <- TDT[Item!=""]
  TDT <- TDT[Item!="工讀生"]
  TDT$UniKey <- NULL
  
  ##
  TDT[, Sum:=sum(as.numeric(Freq),na.rm=T), by = c("學校名稱", "科系名稱")]
  TDT <- TDT[order(學校名稱, 科系名稱, -Freq),]
  TDT <- TDT[Item!="其他"]
  TDT
  TDT <- TDT[, head(.SD, 10), by = c("學校名稱", "科系名稱")]
  
  ##
  print(" Patch")
  lack <- table(paste0(TDT$學校名稱, TDT$科系名稱)) %>% as.data.frame(stringsAsFactors=F) %>% .[which(.$Freq!=10), 1]
  lack <- TDT[paste0(學校名稱, 科系名稱) %in% lack, .(學校名稱, 科系名稱)] %>% unique
  if(nrow(lack)>0){
    for(i in 1:nrow(lack)){
      TDT <- rbindlist(list(TDT,
                            as.list(c(lack[i] %>% c %>% unlist, 
                                      c("行政人員", "國內業務人員", "門市／店員／專櫃人員", "客服人員", "行政助理", "業務助理", "國貿助理", "行銷企劃人員", "採購助理", "餐飲服務人員")[!(c("行政人員", "國內業務人員", "門市／店員／專櫃人員", "客服人員", "行政助理", "業務助理", "國貿助理", "行銷企劃人員", "採購助理", "餐飲服務人員") %in% TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Item])][1], 
                                      ifelse(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Freq][length(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Item])]>0,TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Freq][length(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Item])],1),
                                      TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Sum][1] %>% as.numeric + ifelse(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Freq][length(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Item])]>0,TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Freq][length(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Item])],1) %>% as.numeric
                            ))
      ))
      TDT$Sum[TDT$學校名稱==lack$學校名稱[i] & TDT$科系名稱==lack$科系名稱[i]] <- TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Sum][length(TDT[學校名稱==lack$學校名稱[i] & 科系名稱==lack$科系名稱[i], Sum])]
      cat("\r Caculating...", format(round(i/nrow(lack)*100, 3), nsmall=3), "%")
    }
  }
  
  
  #is.list(TDT)
  #table(Encoding(TDT[[3]]))
  #tmp <- TDT$Item %>% iconv("UTF8", "UTF8")
  #stri_enc_mark(TDT$Item)
  #all(stri_enc_isutf8(TDT$Item))
  #TDT$Item <- stri_encode(TDT$Item, "", "UTF-8")
  #tmp <- stri_trans_general(tmp, "Latin-ASCII")
  #tmp <-tmp %>% iconv(from="UTF-8", to="UTF8")
  TDT[, names(TDT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
  TDT <- rbind(TDT[Item!="No Data",], TDT[Item=="No Data",])
  #TDT[, names(TDT) := lapply(.SD, function(x) {if (is.character(x)) x %>% iconv("UTF-8"); x})]
  #apply(TDT, 2, class)
  #for(i in 1:ncol(TDT)){
  #  TDT[[i]] <- TDT[[i]] %>% iconv("UTF-8")
  #}
  
  ##Use this logic?
  cat("\n")
  TDT[, SumWithoutOthers:=sum(as.numeric(Freq),na.rm=T), by = c("學校名稱", "科系名稱")]
  tmp <- TDT[ , .(學校名稱, 科系名稱, Sum, SumWithoutOthers)] %>% unique
  tmp[,Freq:=as.numeric(Sum)-as.numeric(SumWithoutOthers)]
  tmp <- cbind(tmp, "其他")
  names(tmp)[length(names(tmp))] <- "Item"
  TDT <- rbind(TDT, tmp)
  
  ## Percentage
  TDT[, Percentage:=as.numeric(Freq)/as.numeric(Sum)]
  TDT$Sum <- NULL
  TDT$SumWithoutOthers <- NULL
  TDT$Percentage[which(TDT$Freq==0)] <- 0
  TDT <- TDT[order(學校名稱, 科系名稱),]
  
  ##table(paste0(TDT$學校名稱, TDT$科系名稱)) %>% as.data.frame(stringsAsFactors=F) %>% .[which(.$Freq!=10),]
  ##TDT[學校名稱=="XXX" & 科系名稱=="XXX"]
  #for(i in 1:ncol(TDT)){
  #  TDT[[i]] <- TDT[[i]] %>% iconv("UTF-8")
  #}
  #which(is.na(TDT$Item)) %% 11 %>% unique
  #TDT$Item[which(is.na(TDT$Item))] <- "其他"
  TDT$排名 <- 1:11
  TDT$類別 <- switchV
  
  names(TDT)[3:5] <- names(ODT)[which(names(ODT)==target):(which(names(ODT)==target)+2)]
  
  
  print("Done.")
  return(TDT)
}

## Sample balance
#colName="樣本數.一."
#DF1=totalJ
#DF2=totalI
BalanceModule <- function(DF1, DF2, colName){
  
  stopifnot(nrow(DF1)==nrow(DF2))
  stopifnot(paste0(DF1$學校名稱, DF1$科系名稱)==paste0(DF2$學校名稱, DF2$科系名稱))
  
  colNameParse <- parse(text=colName)
  
  DF1sum <- DF1[, list(Sum=sum(as.numeric(eval(colNameParse)))), by=.(學校名稱, 科系名稱)]
  DF2sum <- DF2[, list(Sum=sum(as.numeric(eval(colNameParse)))), by=.(學校名稱, 科系名稱)]
  
  # test
  stopifnot(paste0(DF1sum$學校名稱, DF1sum$科系名稱)==paste0(DF2sum$學校名稱, DF2sum$科系名稱))
  
  ## Difference between both DF
  DF1sum[,Diff:=Sum-DF2sum$Sum]
  DF1sum <- DF1sum[Diff!=0]
  
  # Balance
  if(nrow(DF1sum)>0){
    for(i in 1:nrow(DF1sum)){
      if(DF1sum$Diff[i] > 0){
        #DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], c("學校名稱", "科系名稱", colName), with=F]
        
        ##Have Item : Others...
        if(DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][11] %>% c %>% as.numeric > 0){
          
          DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==11, 
                  eval(colName):=
                    as.character(DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][11] %>% c %>% as.numeric + abs(DF1sum$Diff[i])),
                  with=F]
          ##Add to the first one..
        }else if(DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][1] %>% c %>% as.numeric > 0){
          DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==1, 
                  eval(colName):=
                    as.character(DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][1] %>% c %>% as.numeric + abs(DF1sum$Diff[i])),
                  with=F]
        }else{
          DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], eval(names(DF2)[which(names(DF2)==colName)-1]):="Error", with=F]
          DF2[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==1, 
                  eval(colName):=
                    as.character(abs(DF1sum$Diff[i])),
                  with=F]
        }
      }else{
        #DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], c("學校名稱", "科系名稱", colName), with=F]
        
        ##Have Item : Others...
        if(DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][11] %>% c %>% as.numeric > 0){
          
          DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==11, 
                  eval(colName):=
                    as.character(DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][11] %>% c %>% as.numeric + abs(DF1sum$Diff[i])),
                  with=F]
          ##Add to the first one..
        }else if(DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][1] %>% c %>% as.numeric > 0){
          DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==1, 
                  eval(colName):=
                    as.character(DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], colName, with=F][1] %>% c %>% as.numeric + abs(DF1sum$Diff[i])),
                  with=F]
        }else{
          DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i], eval(names(DF1)[which(names(DF1)==colName)-1]):="Error", with=F]
          DF1[學校名稱==DF1sum$學校名稱[i] & 科系名稱==DF1sum$科系名稱[i] & 排名==1, 
                  eval(colName):=
                    as.character(abs(DF1sum$Diff[i])),
                  with=F]
        }
      }
    }
    print("Balance complete.")
    
    DF1[,eval(parse(text=names(DF1)[which(names(DF1)==colName)+1])):=as.character(as.numeric(eval(parse(text=names(DF1)[which(names(DF1)==colName)])))/sum(as.numeric(eval(parse(text=names(DF1)[which(names(DF1)==colName)]))))), by=.(學校名稱, 科系名稱)]
    DF2[,eval(parse(text=names(DF2)[which(names(DF2)==colName)+1])):=as.character(as.numeric(eval(parse(text=names(DF2)[which(names(DF2)==colName)])))/sum(as.numeric(eval(parse(text=names(DF2)[which(names(DF2)==colName)]))))), by=.(學校名稱, 科系名稱)]
  }else{
    print("Need not to balance.")
  }
  
  return(list(DF1, DF2))
}

########################
## Start 1-1: Education
########################
School1 <- DStatModule(Education, "SchoolName", "Department", "晉升學校名稱")
M1   <- UpdateModule(School1, OldEducation, "名稱", "學校")
#data.frame(table(paste0(M1$學校名稱, M1$科系名稱))) %$% .[which(Freq!=11),]
setkey(M1, 學校名稱, 科系名稱, 排名, 類別)

names(OldEducation)[which(grepl("類別", names(OldEducation)))] <- "類別"

tmp <- OldEducation[, names(OldEducation)[which(!grepl("^[名稱 | 百分比 | 樣本數]",names(OldEducation)))], with=F]
tmp <- tmp[類別=="1"]
setkey(tmp, 學校名稱, 科系名稱, 排名, 類別)

M1 <- M1[, lapply(.SD, as.character)]
setkey(M1, 學校名稱, 科系名稱, 排名, 類別)

totalS <- tmp[M1]
#merge(totalJ, tmp, all=T)
#apply(totalJ, 2, class) %>% unique
#apply(tmp, 2, class) %>% unique

totalS <- totalS[,c(names(totalS)[which(names(totalS)!="類別" & names(totalS)!="排名")], "類別", "排名"), with=F]
totalS <- totalS[order(學校名稱, 科系名稱, as.numeric(排名))]
#write.csv(totalS, "測試學校.csv", row.names=F)

########################
## Start 1-2: Education
########################
Department1 <- DStatModule(Education, "SchoolName", "Department", "科系類別名稱")
M1   <- UpdateModule(Department1, OldEducation, "名稱", "學類")
setkey(M1, 學校名稱, 科系名稱, 排名, 類別)

names(OldEducation)[which(grepl("類別", names(OldEducation)))] <- "類別"

tmp <- OldEducation[, names(OldEducation)[which(!grepl("^[名稱 | 百分比 | 樣本數]",names(OldEducation)))], with=F]
tmp <- tmp[類別=="0"]
setkey(tmp, 學校名稱, 科系名稱, 排名, 類別)

M1 <- M1[, lapply(.SD, as.character)]
setkey(M1, 學校名稱, 科系名稱, 排名, 類別)

totalD <- tmp[M1]
#merge(totalJ, tmp, all=T)
#apply(totalJ, 2, class) %>% unique
#apply(tmp, 2, class) %>% unique

totalD <- totalD[,c(names(totalD)[which(names(totalD)!="類別" & names(totalD)!="排名")], "類別", "排名"), with=F]
totalD <- totalD[order(學校名稱, 科系名稱, as.numeric(排名))]

#############################
# Merge : School & Department
#############################

listJI <- BalanceModule(totalS, totalD, "樣本數")
totalM <- rbind(listJI[[1]], listJI[[2]])
write.csv(totalM, paste0("output\\", Sys.Date(),"\\升學藍圖.csv"), row.names=F)

##########################################################################
##########################################################################
## End
##########################################################################
##########################################################################

########################
## Start 2-1: Employment
########################

## Encoding problem
#ODT[, names(ODT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
#ODT[, names(ODT) := lapply(.SD, function(x) {if (is.character(x)) x %>% iconv("UTF-8")})]
#for(i in 1:ncol(ODT)){
#  ODT[[i]] <- ODT[[i]] %>% iconv("UTF-8")
#}
#NDT[, names(NDT) := lapply(.SD, function(x) {if (is.character(x)) Encoding(x) <- "unknown"; x})]
#NDT[, names(NDT) := lapply(.SD, function(x) {if (is.character(x)) x %>% iconv("UTF-8")})]
#for(i in 1:ncol(NDT)){
#  NDT[[i]] <- NDT[[i]] %>% iconv("UTF-8", "UTF-8")
#}

Job1 <- DStatModule(Employment, "SchoolName", "Department", "職務小類名稱")
M1   <- UpdateModule(Job1, OldEmployment, "名稱(一)", "職務")
Job2 <- DStatModule(Employment, "SchoolName", "Department", "職務小類名稱1")
M2   <- UpdateModule(Job2, OldEmployment, "名稱(二)", "職務")
Job3 <- DStatModule(Employment, "SchoolName", "Department", "職務小類名稱2")
M3   <- UpdateModule(Job3, OldEmployment, "名稱(三)", "職務")
Job4 <- DStatModule(Employment, "SchoolName", "Department", "職務小類名稱3")
M4   <- UpdateModule(Job4, OldEmployment, "名稱(四)", "職務")

setkey(M1, 學校名稱, 科系名稱, 排名, 類別)
setkey(M2, 學校名稱, 科系名稱, 排名, 類別)
setkey(M3, 學校名稱, 科系名稱, 排名, 類別)
setkey(M4, 學校名稱, 科系名稱, 排名, 類別)

names(OldEmployment)[which(grepl("類別", names(OldEmployment)))] <- "類別"

tmp <- OldEmployment[, names(OldEmployment)[which(!grepl("^[名稱 | 百分比 | 樣本數]",names(OldEmployment)))], with=F]
tmp <- tmp[類別=="0"]
setkey(tmp, 學校名稱, 科系名稱, 排名, 類別)

totalJ <- Reduce(function(...) merge(..., all = T), list(M1, M2, M3, M4))

names(totalJ)
totalJ <- totalJ[, lapply(.SD, as.character)]
setkey(totalJ, 學校名稱, 科系名稱, 排名, 類別)

totalJ <- tmp[totalJ]
#merge(totalJ, tmp, all=T)
#apply(totalJ, 2, class) %>% unique
#apply(tmp, 2, class) %>% unique

totalJ <- totalJ[,c(names(totalJ)[which(names(totalJ)!="類別" & names(totalJ)!="排名")], "類別", "排名"), with=F]
totalJ <- totalJ[order(學校名稱, 科系名稱, as.numeric(排名))]

#write.csv(totalJ, "職務測試.csv", row.names=F)

########################
## Start 2-2: Industry
########################
industry1 <- DStatModule(Employment, "SchoolName", "Department", "產業小類名稱")
M1   <- UpdateModule(industry1, OldEmployment, "名稱(一)", "產業")
industry2 <- DStatModule(Employment, "SchoolName", "Department", "產業小類名稱1")
M2   <- UpdateModule(industry2, OldEmployment, "名稱(二)", "產業")
industry3 <- DStatModule(Employment, "SchoolName", "Department", "產業小類名稱2")
M3   <- UpdateModule(industry3, OldEmployment, "名稱(三)", "產業")
industry4 <- DStatModule(Employment, "SchoolName", "Department", "產業小類名稱3")
M4   <- UpdateModule(industry4, OldEmployment, "名稱(四)", "產業")

setkey(M1, 學校名稱, 科系名稱, 排名, 類別)
setkey(M2, 學校名稱, 科系名稱, 排名, 類別)
setkey(M3, 學校名稱, 科系名稱, 排名, 類別)
setkey(M4, 學校名稱, 科系名稱, 排名, 類別)

names(OldEmployment)[which(grepl("類別", names(OldEmployment)))] <- "類別"

tmp <- OldEmployment[, names(OldEmployment)[which(!grepl("^[名稱 | 百分比 | 樣本數]",names(OldEmployment)))], with=F]
tmp <- tmp[類別=="1"]
setkey(tmp, 學校名稱, 科系名稱, 排名, 類別)

totalI <- Reduce(function(...) merge(..., all = T), list(M1, M2, M3, M4))

names(totalI)
totalI <- totalI[, lapply(.SD, as.character)]
setkey(totalI, 學校名稱, 科系名稱, 排名, 類別)

totalI <- tmp[totalI]
#merge(totalI, tmp, all=T)
#apply(totalI, 2, class) %>% unique
#apply(tmp, 2, class) %>% unique

totalI <- totalI[,c(names(totalI)[which(names(totalI)!="類別" & names(totalI)!="排名")], "類別", "排名"), with=F]
totalI <- totalI[order(學校名稱, 科系名稱, as.numeric(排名))]

########################
# Merge : Job & Industry
########################

listJI <- BalanceModule(totalJ, totalI, "樣本數.一.")
listJI <- BalanceModule(listJI[[1]], listJI[[2]], "樣本數.二.")
listJI <- BalanceModule(listJI[[1]], listJI[[2]], "樣本數.三.")
listJI <- BalanceModule(listJI[[1]], listJI[[2]], "樣本數.四.")
totalM <- rbind(listJI[[1]], listJI[[2]])
write.csv(totalM, paste0("output\\", Sys.Date(),"\\就業藍圖.csv"), row.names=F)
