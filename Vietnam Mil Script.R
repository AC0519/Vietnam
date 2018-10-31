library(dplyr)
library(ggplot2)
library(tidyr)
library(readxl)
library(caTools)
library(randomForest)
library(rpart)
library(plotly)


#######
#Load and Clean Data
######

VietnamDebriefInventory071818 <- read_excel("C://VietnamDebriefInventory071818.xlsx", sheet = "Military")

MilPOWViet <- VietnamDebriefInventory071818

MilPOWViet$X__1 <- NULL
MilPOWViet$`Major Camps (not in order)` <- NULL

MilPOWViet <- rename(MilPOWViet, Last = 'Last Name', First = 'First Name', TypeDebrief = 'Type of Debrief', DebriefingAgency = 'Debriefing Agency', DateDebrief = 'Date of Debrief', NumPages = '# of Pages')


###Seperate Rank/Branch into two columns

#Branch Column
MilPOWViet$Branch <- MilPOWViet$`Rank/Branch`

MilPOWViet$Branch <- gsub(".*,","",MilPOWViet$Branch)

#get rid of rows where a branch is not listed
MilPOWViet <- subset(MilPOWViet,!is.na(MilPOWViet$Branch))

#Clean up Branch so there are only 4 branches of service
service <- function(Branch){
  if (Branch == ' USAF' | Branch == 'USAF'){
    return('USAF')
  }else if(Branch == ' USA' | Branch == 'USA'){
    return('USA')
  }else if(Branch == " USN"){
    return('USN')
  }else if(Branch == " USMC"){
    return('USMC')
  }else{
    return(Branch)
  }
}

MilPOWViet$Branch <- sapply(MilPOWViet$Branch, service)

#Rank Column
MilPOWViet <- rename(MilPOWViet, Rank = "Rank/Branch")

MilPOWViet$Rank <- gsub(",.*","",MilPOWViet$Rank)


###rewrite Rank column as E/O format

rankfunc <- function(Rank){
  if (Rank=='1 LT' | Rank=='1LT' | Rank=='LTJG' | Rank=='1stLT' | Rank=='1 Lt'| Rank=='LTJR'){
    return('O-2')
  }else if (Rank =='LT' | Rank == 'Capt'| Rank == 'CPT'| Rank == 'Cpt'){
    return('O-3')
  }else if (Rank == 'Major'| Rank == 'LCDR'| Rank == 'Maj'| Rank == 'MAJ'){
    return('O-4')
  }else if (Rank == 'Lt. Col'| Rank == 'Lt Col'| Rank == 'LT Col'| Rank == 'LtCol'| Rank == 'CDR'| Rank == 'LTC'){
    return('O-5')
  }else if (Rank == 'Col'){
    return('O-6')
  }else if (Rank == 'W01'){
    return('W1')
  }else if (Rank == 'CW2'){
    return('W2')
  }else if (Rank == 'CW3'){
    return('W3')
  }else if (Rank == 'Pvt'){
    return('E-2')
  }else if (Rank == 'LCpl' | Rank == 'PFC'){
    return('E-3')
  }else if (Rank == 'CPL'| Rank == 'SP4'| Rank == 'SPC'| Rank == 'SN'| Rank == 'Cpl'| Rank == 'USA'){
    return('E-4')
  }else if (Rank == 'Sgt'| Rank == 'SP5'| Rank == 'SGT'| Rank == 'SA'| Rank == 'PO3'| Rank == 'Ssgt'){
    return('E-5')
  }else if (Rank == 'TSgt'| Rank == 'SSGT'| Rank == 'SP6'| Rank == 'SSG'| Rank == 'SSgt'){
    return('E-6')
  }else if (Rank == 'SFC'| Rank == 'Msgt'| Rank == 'MSgt'){
    return('E-7')
  }else if (Rank == 'MSG'| Rank == 'MSGT'){
    return('E-8')
  }else{
    return(Rank)
  }
}

MilPOWViet$Rank <- sapply(MilPOWViet$Rank, rankfunc)


###Seperate Captivity period into two columns and turn into timeseries data


#Endcap 
MilPOWViet$Endcap <- MilPOWViet$`Captivity Period`  

MilPOWViet$Endcap <- gsub(".*-","",MilPOWViet$Endcap)

MilPOWViet$Endcap <- strptime(MilPOWViet$Endcap, "%d %b %Y")


#Startcap
MilPOWViet <- rename(MilPOWViet, Startcap = "Captivity Period")

MilPOWViet$Startcap <- gsub("-.*","",MilPOWViet$Startcap)

MilPOWViet$Startcap <- strptime(MilPOWViet$Startcap, "%d %b %Y")


#create time in captivity column

MilPOWViet$TimeinCap <- abs(ceiling(((MilPOWViet$Endcap - MilPOWViet$Startcap) / 86400)))


#Remove cases with NA time in captivity
MilPOWViet <- subset(MilPOWViet,!is.na(MilPOWViet$TimeinCap))

#Mean, Min, and Max time of all cases in captivity
MilPOWViet$TimeinCap <- as.numeric(MilPOWViet$TimeinCap)

mean(MilPOWViet$TimeinCap)
min(MilPOWViet$TimeinCap)
max(MilPOWViet$TimeinCap)


##############
#EDA
#############

#Cases by branch of service
ggplot(MilPOWViet) +geom_bar(aes(Branch))

#Cases by Rank
ggplot(MilPOWViet) +geom_bar(aes(Rank))

#cases by branch and rank
BranchColors <- c("USA" = "dark green","USAF"="#009ACD", "USMC"="#698B22", "USN"="#000080")

pl <- ggplot(MilPOWViet)+
geom_bar(aes(Rank, fill = Branch), position = "dodge")+
theme(panel.background = element_blank(), panel.grid.major.y = element_line(color = "grey"),plot.title = element_text(hjust = .5, size = rel(2), face = "bold"))+
ggtitle("POWs by Rank and Branch of Service")+
scale_fill_manual(values = BranchColors, name = "Branch of Service")+
scale_y_continuous(name= "Number of POWs")

ggplotly(pl)

#box plot of time in captivity
ggplot(MilPOWViet, aes(x= Branch, y= TimeinCap))+
geom_boxplot(outlier.size = 2, outlier.color = "red")

#box plot of time in captivity
ggplot(MilPOWViet, aes(x= Rank, y= TimeinCap))+
geom_boxplot(outlier.size = 2, outlier.color = "red")


#type of debrief
ggplot(MilPOWViet) + geom_bar(aes(MilPOWViet$TypeDebrief))


#########
#Random Forest Experiment
#########

MilPOWViet$Startcap <- as.POSIXct(MilPOWViet$Startcap)
MilPOWViet$Endcap <- as.POSIXct(MilPOWViet$Endcap)
MilPOWViet$Branch <- as.factor(MilPOWViet$Branch)
MilPOWViet$Rank <- as.factor(MilPOWViet$Rank)

VietLR <- select(MilPOWViet, Rank, Branch, Startcap, Endcap, TimeinCap)

sample <- sample.split(VietLR$Branch, SplitRatio = .7)
train <- subset(VietLR, sample==T)
test <- subset(VietLR, sample==F)

rf.model <- randomForest(Branch ~ ., data = train, importance=T)
rf.model$confusion
rf.model$importance

rf.predictions <- predict(rf.model, test)
table(rf.predictions, test$Branch)

test$Branch.Predict <- rf.predictions










