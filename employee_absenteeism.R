rm(list=ls())

## Set Directory
setwd("C:/Users/Shivam/Desktop/Data Science edWisor/PROJECT/R")

getwd()

## load libraries
x =c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
       "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees')

#install.packages(x)
lapply(x, require,character.only = TRUE)
rm(x)

## Read Data
library(readxl)
data = read_excel("Absenteeism_at_work_Project.xls")
data = as.data.frame(data)
### Explore the data
dim(data)

str(data)

colnames(data)

## Remove space between columns names
names(data) = gsub(" ","_",names(data))
names(data) = gsub("/","_",names(data))

############################ Missing Value Analysis ##################################
missing_val = data.frame(apply(data,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(data)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]

 ggplot(data = missing_val[1:18,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
   geom_bar(stat = "identity",fill = "red")+xlab("Parameter")+
   ggtitle("Missing data percentage") + theme_bw()
 

## Remove all the 22 observation having values "NA" in target variable
 data = DropNA(data,Var="Absenteeism_time_in_hours")

## KNN Imputation
 data = knnImputation(data,k=3)
data = trunc(data)

## Convert into appropriate  categorical variable
data$ID = as.factor(as.numeric(data$ID))
data$Reason_for_absence = as.factor(as.numeric(data$Reason_for_absence))
data$Month_of_absence = as.factor(as.numeric(data$Month_of_absence))
data$Day_of_the_week = as.factor(as.numeric(data$Day_of_the_week))
data$Seasons = as.factor(as.numeric(data$Seasons))
data$Social_smoker = as.factor(as.numeric(data$Social_smoker))
data$Social_drinker = as.factor(as.numeric(data$Social_drinker))
data$Education = as.factor(as.numeric(data$Education))
data$Disciplinary_failure = as.factor(as.numeric(data$Disciplinary_failure))

####################### Data Visualization ####################


ggplot(data=data, aes(x=Reason_for_absence, y=Absenteeism_time_in_hours)) + geom_bar(stat="Identity")
ggplot(data = data,aes(x=ID , y=Absenteeism_time_in_hours))+geom_bar(stat = "Identity")
ggplot(data =data,aes(x=Month_of_absence , y=Absenteeism_time_in_hours)) + geom_bar(stat = "Identity")
ggplot(data=data,aes(x=Day_of_the_week , y=Absenteeism_time_in_hours)) + geom_bar(stat = "Identity")
ggplot(data=data,aes(x=Seasons,y=Absenteeism_time_in_hours))+geom_bar(stat = "Identity")
ggplot(data=data,aes(x=Disciplinary_failure  , y= Absenteeism_time_in_hours))+geom_bar(stat = "Identity")

plot(data$Body_mass_index,data$Absenteeism_time_in_hours)
plot(data$Transportation_expense,data$Absenteeism_time_in_hours)
plot(data$Distance_from_Residence_to_Work,data$Absenteeism_time_in_hours)
plot(data$Service_time,data$Absenteeism_time_in_hours)
plot(data$Age,data$Absenteeism_time_in_hours)
plot(data$Work_load_Average_day,data$Absenteeism_time_in_hours)

########################## Outlier Analysis ################################
numeric_index = sapply(data,is.numeric)
numeric_data = data[,numeric_index]
cnames = colnames(numeric_data)

 for (i in 1:length(cnames))
 {
   assign(paste0("gn",i), ggplot(aes_string(y = (cnames[i]), x = "Absenteeism_time_in_hours"), data = subset(data))+ 
            stat_boxplot(geom = "errorbar", width = 0.5,group=1) +
            geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                         outlier.size=1, notch=TRUE) +
            theme(legend.position="bottom")+
            labs(y=cnames[i],x="Absenteeism_time_in_hours")+
            ggtitle(paste("Box plot of ",cnames[i])))
 }
 
 ## Plotting plots together
 gridExtra::grid.arrange(gn1,gn2,gn3,ncol=3)
 gridExtra::grid.arrange(gn4,gn5,ncol=2)
 gridExtra::grid.arrange(gn6,gn7,ncol=2)
 gridExtra::grid.arrange(gn8,gn9,ncol=2)
 gridExtra::grid.arrange(gn10,gn11,ncol=2)
 #gridExtra::grid.arrange(gn12,ncol=1)
 
for (i in cnames){
  print(i)
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  print(length(val))
  data[,i][data[,i] %in% val] = NA
}

 data = knnImputation(data, k = 3)
 
##################################### Feature Selection #################################
 ## Correlation Plot
 corrgram(data[,numeric_index], order = F,
          upper.panel=panel.pie, text.panel=panel.txt, cor.method = "pearson",main = "Correlation Plot")


 #Anova test
 anova_test=aov(Absenteeism_time_in_hours ~ID+Reason_for_absence+Month_of_absence+Day_of_the_week+Seasons+
                  Education+Disciplinary_failure+Social_drinker+Social_smoker,data = data)
 
 
 summary(anova_test)
 
 anova_test_1 = aov(Absenteeism_time_in_hours ~Disciplinary_failure+Social_drinker+Social_smoker,
                    data=data)
 summary(anova_test_1)
 

 
 
 ## Dimension Reduction
 data = subset(data,select=-c(Weight,Month_of_absence,Day_of_the_week,Seasons,Education,Social_smoker))

 ################################# Feature Scaling ########################################
 ## Normality Check
 
 hist(data$Service_time)
 hist(data$Distance_from_Residence_to_Work)
hist(data$Transportation_expense) 
hist(data$Age)
hist(data$Work_load_Average_day)
hist(data$Hit_target)
hist(data$Son)
hist(data$Pet)
hist(data$Height)
hist(data$Body_mass_index)

## Normalization
 
num_names = c("Transportation_expense","Service_time","Distance_from_Residence_to_Work","Age","Work_load_Average_day",
              "Hit_target","Son","Pet","Height","Body_mass_index","Absenteeism_time_in_hours")
for ( i in num_names){
  print(i)
  data[,i] = (data[,i] - min(data[,i]))/
    (max(data[,i] - min(data[,i])))
}

########################################## Sampling #########################################
## Divide the data into train and test using stratified sampling method

set.seed(1234)
data_index = createDataPartition(data$Absenteeism_time_in_hours , p=0.80 , list = F)
train = data[data_index,]
test = data[-data_index,]


####################################### Model Development ###################################

## Decision Tree Regression
fit = rpart(formula = Absenteeism_time_in_hours ~. ,data=train , method = "anova")
summary(fit)
predictions_DT = predict(fit, test[,-15])

regr.eval(test[,15],predictions_DT,stats = c("mae","mape","rmse"))
### rmse = 0.1719

#### Random Forest ####
library(randomForest)
 
rf_model= randomForest(Absenteeism_time_in_hours ~.,train,importance=TRUE,ntree=200)

summary(rf_model)

rf_predict=predict(rf_model,test[,-15])


regr.eval(test[,15],rf_predict,stats = c("mae","mape","rmse"))
#### rmse = 0.1627(when ntree=100) , rmse = 0.16214(when ntree=500) , rmse=0.1621(when ntree=1000)

#############
# From the above two models Random Forest is more appropriate .



## Q2. How much losses every month can we project in 2011 if same trend of absenteeism continues?

df = data.frame(read_excel("Absenteeism_at_work_Project.xls"))

df = subset(df,select=c("Month.of.absence","Service.time","Work.load.Average.day","Absenteeism.time.in.hours"))

#Work loss = ((Work load per day/ service time)* Absenteeism hours)
df = DropNA(df,Var="Absenteeism.time.in.hours")
df = knnImputation(df,k=3)
df$loss=with(df,((df$Work.load.Average.day / df$Service.time)*df$Absenteeism.time.in.hours))

No_absent = sum(df$loss[which(df$Month.of.absence %in% 0)])
Jan = sum(df$loss[which(df$Month.of.absence %in% 1)])
Feb = sum(df$loss[which(df$Month.of.absence %in% 2)])
Mar = sum(df$loss[which(df$Month.of.absence %in% 3)])
Apr = sum(df$loss[which(df$Month.of.absence %in% 4)])
May = sum(df$loss[which(df$Month.of.absence %in% 5)])
June = sum(df$loss[which(df$Month.of.absence %in% 6)])
July = sum(df$loss[which(df$Month.of.absence %in% 7)])
Aug = sum(df$loss[which(df$Month.of.absence %in% 8)])
Sept = sum(df$loss[which(df$Month.of.absence %in% 9)])
Oct = sum(df$loss[which(df$Month.of.absence %in% 10)])
Nov = sum(df$loss[which(df$Month.of.absence %in% 11)])
Dec = sum(df$loss[which(df$Month.of.absence %in% 12)])

df1 = data.frame(Months = c("0","Jan","Feb","Mar","Apr","May","June","July","Aug","Sep","Oct","Nov","Dec"), 
                 Loss= c(No_absent,Jan,Feb,Mar,Apr,May,June,July,Aug,Sept,Oct,Nov,Dec))
View(df1)
