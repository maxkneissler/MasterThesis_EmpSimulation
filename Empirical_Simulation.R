################################################################################
###               Empirical Simulation in the Master Thesis                  ###
### The Quest for a Modern and Dynamic Budgeting and Forecasting Approach    ###
###               By Max Kneißler (University of Tübingen)                   ###
################################################################################



### Packages for the analysis

library(data.table)
library(extrafont)
library(ggplot2)
library(ggpubr)
library(gridExtra)
library(matlib)
library(stringr)
library(readxl)
library(stargazer)
library(tidyverse)
library(xts)



### Import of the different sheets of the database

Income_Statement <- read_excel("3M_Data.xlsx", sheet = "Income_Statement")
Sales_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "Sales_per_Division")
OpIncome_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "OpIncome_per_Division")
Sales_per_Region <- read_excel("3M_Data.xlsx", sheet = "Sales_per_Region")
Division_Region_Matrix_Sales <- 
            read_excel("3M_Data.xlsx", sheet = "Division_Region_Matrix_Sales")
Ext_Parameters <- read_excel("3M_Data.xlsx", sheet = "External Parameters")
Ext_Parameters_Chg <- 
            read_excel("3M_Data.xlsx", sheet = "External Parameters Changes")




################################################################################
###             Chapter 5.2: Graphical Illustration of the Data              ###
################################################################################



### Plot 1: Net Sales vs. Operating Income vs. Net Income over time

Fig_One <- Income_Statement[c(1,7,15),] %>% t() %>% as.data.frame()
names(Fig_One) <- Fig_One[1,]
Fig_One <- Fig_One[-1,]
names(Fig_One)[3] <- "Net income"

Fig_One <- cbind(Quarter = rownames(Fig_One), Fig_One)
Fig_One <- Fig_One[rev(rownames(Fig_One)),]
rownames(Fig_One) <- 1:nrow(Fig_One)
Fig_One <- gather(Fig_One, "Net sales", "Operating income", "Net income", 
                  key = "Income_Position", value = "Income_Figure") %>% 
                  arrange(Income_Position)

quarters <- substr(Fig_One$Quarter, 1, 2)
years <- substr(Fig_One$Quarter, 4, 7)
Fig_One$Quarter <- paste(years, quarters, sep = " - ")

windowsFonts(Times = windowsFont("Times New Roman"))
Fig_One$Income_Position <- 
  factor(Fig_One$Income_Position, 
         levels = c("Net sales", "Operating income", "Net income"))

ggplot(Fig_One, aes(x=Quarter, y=as.numeric(Income_Figure), 
                    colour=Income_Position, group=Income_Position)) + 
                    geom_point() + geom_line() + theme_bw() +
                    xlab("Quarters") + ylab("Figures in Million $") +
                    theme(axis.text.x = element_text(angle = 45, hjust = 1),
                          legend.title = element_blank(),
                          text = element_text(family="Times", 
                                              face = "bold", size = 12),
                          axis.title.y = element_text(margin = margin(r = 7)),
                          axis.title.x = element_text(margin = margin(t = 5))) + 
                    scale_y_continuous(limits = c(0, 9000))



### Plot 2: Net Sales vs. Operating Income in the different divisions over time

DivSalesFig3 <- Sales_per_Division[c(1,2,3,4,7),] %>% t() %>% as.data.frame()
names(DivSalesFig3) <- DivSalesFig3[1,]
DivSalesFig3 <- DivSalesFig3[-1,]
DivSalesFig3 <- cbind(Quarter = rownames(DivSalesFig3), DivSalesFig3)
DivSalesFig3$Quarter <- paste(years, quarters, sep = " - ")
DivSalesFig3 <- DivSalesFig3[rev(rownames(DivSalesFig3)),]
rownames(DivSalesFig3) <- 1:nrow(DivSalesFig3)
names(DivSalesFig3)[2] <- "Safety and Industrial Segment"
names(DivSalesFig3)[3] <- "Transportation and Electronics Segment"
names(DivSalesFig3)[4] <- "Health Care Segment"
names(DivSalesFig3)[5] <- "Consumer Segment"

DivSalesFig3 <- gather(DivSalesFig3, 2,3,4,5, 
                  key = "Division", value = "Net_Sales") %>% arrange(Division)
DivSalesFig3 <- 
  subset(DivSalesFig3, select = c("Quarter", "Division", "Net_Sales"))

DivSalesFig3$Division <- 
  factor(DivSalesFig3$Division, levels = c("Safety and Industrial Segment", 
                                  "Transportation and Electronics Segment", 
                                  "Health Care Segment",
                                  "Consumer Segment"))

SalesDiv <- ggplot(DivSalesFig3, aes(x=Quarter, y=as.numeric(Net_Sales), 
                    colour=Division, group=Division)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("") + ylab("Net Sales in Million $") +
  theme(legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_text(margin = margin(r = 7)),
        axis.title.x = element_text(size = 0.5),
        axis.text.x = element_blank(),
        legend.position = "none") + 
  scale_y_continuous(limits = c(500, 3500))



DivOpFig3 <- OpIncome_per_Division[c(1,2,3,4,7),] %>% t() %>% as.data.frame()
names(DivOpFig3) <- DivOpFig3[1,]
DivOpFig3 <- DivOpFig3[-1,]
DivOpFig3 <- cbind(Quarter = rownames(DivOpFig3), DivOpFig3)
DivOpFig3$Quarter <- paste(years, quarters, sep = " - ")
DivOpFig3 <- DivOpFig3[rev(rownames(DivOpFig3)),]
rownames(DivOpFig3) <- 1:nrow(DivOpFig3)
DivOpFig3 <- DivOpFig3 %>% select(-6)
names(DivOpFig3)[2] <- "Safety and Industrial Segment"
names(DivOpFig3)[3] <- "Transportation and Electronics Segment"
names(DivOpFig3)[4] <- "Health Care Segment"
names(DivOpFig3)[5] <- "Consumer Segment"

DivOpFig3 <- gather(DivOpFig3, 2,3,4,5, 
            key = "Division", value = "Operating_Income") %>% arrange(Division)
DivOpFig3 <- 
  subset(DivOpFig3, select = c("Quarter", "Division", "Operating_Income"))

DivOpFig3$Division <- 
  factor(DivOpFig3$Division, levels = c("Safety and Industrial Segment", 
                                        "Transportation and Electronics Segment", 
                                        "Health Care Segment",
                                        "Consumer Segment"))

OpIncDiv <- ggplot(DivOpFig3, aes(x=Quarter, y=as.numeric(Operating_Income), 
                         colour=Division, group=Division)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Operating Income in Million $") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_text(margin = margin(r = 7)),
        axis.title.x = element_text(margin = margin(t = 5)),
        legend.margin = margin(0,0,0,0),
        legend.justification = "top") + 
  scale_y_continuous(limits = c(-800, 1200))

## Combine both plots rowwise
grid.arrange(SalesDiv, OpIncDiv, nrow=5, ncol=1, 
             layout_matrix = rbind(c(1,1,1,NA),
                                   c(1,1,1,NA),
                                   c(1,1,1,NA),
                                   c(1,1,1,NA),
                                   c(3,3,3,3),
                                   c(3,3,3,3),
                                   c(3,3,3,3),
                                   c(3,3,3,3),
                                   c(3,3,3,3)))



### Plot 3: Net Sales across the different regions over time

SalesRegionPlot <- Sales_per_Region[1:3,] %>% t() %>% as.data.frame()
names(SalesRegionPlot) <- SalesRegionPlot[1,]
SalesRegionPlot <- SalesRegionPlot[-1,]
SalesRegionPlot <- cbind(Quarter = rownames(SalesRegionPlot), SalesRegionPlot)
SalesRegionPlot$Quarter <- paste(years, quarters, sep = " - ")
SalesRegionPlot <- SalesRegionPlot[rev(rownames(SalesRegionPlot)),]
rownames(SalesRegionPlot) <- 1:nrow(SalesRegionPlot)

SalesRegionPlot <- gather(SalesRegionPlot, 2,3,4, 
                    key = "Region", value = "Net_Sales") %>% arrange(Region)

ggplot(SalesRegionPlot, aes(x=Quarter, y=as.numeric(Net_Sales), 
                      colour=Region, group=Region)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Net Sales in Million $") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_text(margin = margin(r = 7)),
        axis.title.x = element_text(margin = margin(t = 5))) + 
  scale_y_continuous(limits = c(0, 5000))



### Plot 4: Net Sales across the different regions and divisions over time

DivRegionPlot <- Division_Region_Matrix_Sales[,1:5]
names(DivRegionPlot)[1] <- "Division"
names(DivRegionPlot)[2] <- "Quarter"
DivRegionPlot <- DivRegionPlot %>% 
  filter(!grepl("Total Company", Division))
DivRegionPlot <- DivRegionPlot %>% 
  filter(!grepl("Elimination of Dual Credit", Division))
DivRegionPlot <- DivRegionPlot %>% 
  filter(!grepl("Corporate and Unallocated", Division))

quarters <- substr(DivRegionPlot$Quarter, 1, 2)
years <- substr(DivRegionPlot$Quarter, 4, 7)
DivRegionPlot$Quarter <- paste(years, quarters, sep = " - ")
DivRegionPlot <- DivRegionPlot[rev(rownames(DivRegionPlot)),]

DivRegionPlot <- gather(DivRegionPlot, 3,4,5,
              key = "Region", value = "Net_Sales") %>% arrange(Region) %>%
              subset(select = c("Quarter", "Division", "Region", "Net_Sales"))


## Production the four different plots separately
SafeIndu <- ggplot(DivRegionPlot %>% filter(Division == "Safety & Industrial"), 
       aes(x=Quarter, y=as.numeric(Net_Sales), 
                            colour=Region, group=Region)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Net Sales in Million $") +
  theme(axis.text.x = element_blank(),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_text(margin = margin(r = 7)),
        axis.title.x = element_blank(),
        legend.position = "none") + 
  scale_y_continuous(limits = c(0, 2000)) + ggtitle("Safety & Industrial")


TransElec <- ggplot(DivRegionPlot %>% 
                      filter(Division == "Transportation and Electronics"), 
                   aes(x=Quarter, y=as.numeric(Net_Sales), 
                       colour=Region, group=Region)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Net Sales in Million $") +
  theme(axis.text.x = element_blank(),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_blank(),
        axis.title.x = element_blank(),
        legend.position="none",
        axis.text.y=element_blank()) + 
  scale_y_continuous(limits = c(0, 2000)) + 
  ggtitle("Transportation & Electronics")


Health <- ggplot(DivRegionPlot %>% filter(Division == "Health Care"), 
                   aes(x=Quarter, y=as.numeric(Net_Sales), 
                       colour=Region, group=Region)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Net Sales in Million $") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_text(margin = margin(r = 7)),
        axis.title.x = element_text(margin = margin(t = 5)),
        legend.position = "none") + 
  scale_y_continuous(limits = c(0, 2000)) + ggtitle("Health Care")


Cons <- ggplot(DivRegionPlot %>% filter(Division == "Consumer"), 
                   aes(x=Quarter, y=as.numeric(Net_Sales), 
                       colour=Region, group=Region)) + 
  geom_point() + geom_line() + theme_bw() +
  xlab("Quarters") + ylab("Net Sales in Million $") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank(),
        text = element_text(family="Times", 
                            face = "bold", size = 12),
        axis.title.y = element_blank(),
        axis.title.x = element_text(margin = margin(t = 5)),
        legend.margin = margin(0, 0, 0,0),
        legend.justification = "top",
        axis.text.y=element_blank()) + 
  scale_y_continuous(limits = c(0, 2000)) + ggtitle("Consumer")

## Combine the four plots with equal sizes of the plot (reason for many figures)
grid.arrange(SafeIndu, TransElec, Health, Cons, ncol=2, nrow = 2,
             layout_matrix = 
               rbind(c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
                       1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,
                       2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,NA,NA,NA,NA,NA,NA,NA,NA,
                       NA,NA,NA,NA,NA,NA,NA,NA,NA),
                     c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
                       1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,
                       2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,NA,NA,NA,NA,NA,NA,NA,NA,
                       NA,NA,NA,NA,NA,NA,NA,NA,NA),
                     c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
                       1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,
                       2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,NA,NA,NA,NA,NA,NA,NA,NA,
                       NA,NA,NA,NA,NA,NA,NA,NA,NA),
                     c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
                       1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,
                       2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,NA,NA,NA,NA,NA,NA,NA,NA,
                       NA,NA,NA,NA,NA,NA,NA,NA,NA),
                     c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                       3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4),
                     c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                       3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4),
                     c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                       3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4),
                     c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                       3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4),
                     c(3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,
                       3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,
                       4,4,4,4,4)))




################################################################################
###                   Chapter 5.3: Training of the Algorithm                 ###
################################################################################


### Training and test sets

Division_Region_Matrix_Sales <-
  Division_Region_Matrix_Sales[rev(rownames(Division_Region_Matrix_Sales)),]
DivRegion_Train <- Division_Region_Matrix_Sales[1:112,]
DivRegion_Val <- Division_Region_Matrix_Sales[113:140,]

SalesDiv_Train <- Sales_per_Division %>% subset(select=c(1,8:39))
SalesDiv_Val <- Sales_per_Division %>% subset(select=c(1:7))

SalesReg_Train <- Sales_per_Region %>% subset(select=c(1,8:39))
SalesReg_Val <- Sales_per_Region %>% subset(select=c(1:7))

Ext_Parameters <- Ext_Parameters[rev(rownames(Ext_Parameters)),]
Ext_Parameters_Chg <- Ext_Parameters_Chg[rev(rownames(Ext_Parameters_Chg)),]
Ext_Parameters_Train <- Ext_Parameters[1:32,]
Ext_Parameters_Val <- Ext_Parameters[33:38,]



### Significance test of the external parameters

## 1. Training with respect to the regions

# Sales projection in the Americans
SalesAmerTrain <- SalesReg_Train[1,] %>% t() %>% as.data.frame()
SalesAmerTrain <- cbind(Quarter = rownames(SalesAmerTrain), SalesAmerTrain)
SalesAmerTrain <- SalesAmerTrain[-1,]
SalesAmerTrain <- SalesAmerTrain[rev(rownames(SalesAmerTrain)),]
rownames(SalesAmerTrain) <- 1:nrow(SalesAmerTrain)
names(SalesAmerTrain)[2] <- "Net_Sales"

SalesAmerTrain <- 
  cbind(SalesAmerTrain, Ext_Parameters_Train[,c(2,5,8,11,14,17,20,23)])

# PPP and Exchange Rate not considered as set to 1 as base
ModelAmerSales <- 
  lm(Net_Sales ~ GDP_USA_Perc_Change + Unemp_USA_Perc + IntRate_USA_Perc +
       CPI_USA + Avg_Wage_USA + ConsBaro_USA, SalesAmerTrain)

# Summary Statistics
# sink("BaseAmer.txt")
summary(ModelAmerSales)
# sink()


# Sales projection in Asia Pacific
SalesAsiaTrain <- SalesReg_Train[2,] %>% t() %>% as.data.frame()
SalesAsiaTrain <- cbind(Quarter = rownames(SalesAsiaTrain), SalesAsiaTrain)
SalesAsiaTrain <- SalesAsiaTrain[-1,]
SalesAsiaTrain <- SalesAsiaTrain[rev(rownames(SalesAsiaTrain)),]
rownames(SalesAsiaTrain) <- 1:nrow(SalesAsiaTrain)
names(SalesAsiaTrain)[2] <- "Net_Sales"

SalesAsiaTrain <- 
  cbind(SalesAsiaTrain, Ext_Parameters_Train[,c(4,7,10,13,16,19,22,25)])

ModelAsiaSales <- 
  lm(Net_Sales ~ GDP_CHN_Perc_Change + Umemp_CHN_Perc + IntRate_CHN_Perc + 
PPP_JPN + CPI_CHN + ExchRate_CHN + Avg_Wage_JPN + ConsBaro_CHN, SalesAsiaTrain)

# Summary Statistics
# sink("BaseAsia.txt")
summary(ModelAsiaSales)
# sink()


# Sales projection in Europa / Middle East / Africa
SalesEurTrain <- SalesReg_Train[3,] %>% t() %>% as.data.frame()
SalesEurTrain <- cbind(Quarter = rownames(SalesEurTrain), SalesEurTrain)
SalesEurTrain <- SalesEurTrain[-1,]
SalesEurTrain <- SalesEurTrain[rev(rownames(SalesEurTrain)),]
rownames(SalesEurTrain) <- 1:nrow(SalesEurTrain)
names(SalesEurTrain)[2] <- "Net_Sales"

SalesEurTrain <- 
  cbind(SalesEurTrain, Ext_Parameters_Train[,c(3,6,9,12,15,18,21,24)])

ModelEurSales <- 
  lm(Net_Sales ~ GDP_GER_Perc_Change + Unemp_GER_Perc + IntRate_GER_Perc + PPP_GER + 
       CPI_GER + ExchRate_GER + Avg_Wage_GER + ConsBaro_GER, SalesEurTrain)

# Summary Statisttics
# sink("BaseEur.txt)")
summary(ModelEurSales)
# sink()



## 2. Training with respect to the divisions

# Safety & Industrial
SalesSafeInduTrain <- SalesDiv_Train[1,] %>% t() %>% as.data.frame()
SalesSafeInduTrain <- 
  cbind(Quarter = rownames(SalesSafeInduTrain), SalesSafeInduTrain)
SalesSafeInduTrain <- SalesSafeInduTrain[-1,]
SalesSafeInduTrain <- SalesSafeInduTrain[rev(rownames(SalesSafeInduTrain)),]
rownames(SalesSafeInduTrain) <- 1:nrow(SalesSafeInduTrain)
names(SalesSafeInduTrain)[2] <- "Net_Sales"

SalesSafeInduTrain <- cbind(SalesSafeInduTrain, Ext_Parameters_Train[,2:25])
ModelSafeIndu <- lm(Net_Sales ~ . - Quarter, SalesSafeInduTrain)

# Summary Statistics
# sink("BaseSafeIndu.txt)")
summary(ModelSafeIndu)
# sink()


# Transportation and Electronics
SalesTransElecTrain <- SalesDiv_Train[2,] %>% t() %>% as.data.frame()
SalesTransElecTrain <- 
  cbind(Quarter = rownames(SalesTransElecTrain), SalesTransElecTrain)
SalesTransElecTrain <- SalesTransElecTrain[-1,]
SalesTransElecTrain <- SalesTransElecTrain[rev(rownames(SalesTransElecTrain)),]
rownames(SalesTransElecTrain) <- 1:nrow(SalesTransElecTrain)
names(SalesTransElecTrain)[2] <- "Net_Sales"

SalesTransElecTrain <- cbind(SalesTransElecTrain, Ext_Parameters_Train[,2:25])
ModelTransElec <- lm(Net_Sales ~ . - Quarter, SalesTransElecTrain)

# Summary Statistics
# sink("BaseTransElec.txt)")
summary(ModelTransElec)
# sink()


# Health Care
SalesHealthTrain <- SalesDiv_Train[3,] %>% t() %>% as.data.frame()
SalesHealthTrain <- 
  cbind(Quarter = rownames(SalesHealthTrain), SalesHealthTrain)
SalesHealthTrain <- SalesHealthTrain[-1,]
SalesHealthTrain <- SalesHealthTrain[rev(rownames(SalesHealthTrain)),]
rownames(SalesHealthTrain) <- 1:nrow(SalesHealthTrain)
names(SalesHealthTrain)[2] <- "Net_Sales"

SalesHealthTrain <- cbind(SalesHealthTrain, Ext_Parameters_Train[,2:25])
ModelHealth <- lm(Net_Sales ~ . - Quarter, SalesHealthTrain)

# Summary Statistics
# sink("BaseHealth.txt)")
summary(ModelHealth)
# sink()


# Consumer
SalesConsTrain <- SalesDiv_Train[4,] %>% t() %>% as.data.frame()
SalesConsTrain <- cbind(Quarter = rownames(SalesConsTrain), SalesConsTrain)
SalesConsTrain <- SalesConsTrain[-1,]
SalesConsTrain <- SalesConsTrain[rev(rownames(SalesConsTrain)),]
rownames(SalesConsTrain) <- 1:nrow(SalesConsTrain)
names(SalesConsTrain)[2] <- "Net_Sales"

SalesConsTrain <- cbind(SalesConsTrain, Ext_Parameters_Train[,2:25])
ModelCons <- lm(Net_Sales ~ . - Quarter, SalesConsTrain)

# Summary Statistics
# sink("BaseCons.txt)")
summary(ModelCons)
# sink()



## 3. Training for matrix element of division and region


# Eliminate all unallocated positions
names(DivRegion_Train)[1] <- "Division"
DivRegion_Train <- DivRegion_Train[,1:5] %>% 
            filter(!grepl("Corporate and Unallocated", Division)) %>%
            filter(!grepl("Total Company", Division)) %>%
            filter(!grepl("Elimination of Dual Credit", Division))


# External Parameters only for the considered time horizon
DivRegExt_Paramter <- Ext_Parameters_Train[17:32,]



## Division: Safety & Industrial, look at all regions separately
SafeInduRegion_Train <- 
            DivRegion_Train %>% filter(Division == "Safety & Industrial")
SafeInduRegion_Train <- cbind(SafeInduRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
SI_Amer <- lm(Americas ~ GDP_USA_Perc_Change + Unemp_USA_Perc + IntRate_USA_Perc + 
                CPI_USA + Avg_Wage_USA, SafeInduRegion_Train)
# sink("SI_Amer.txt)")
summary(SI_Amer)
# sink()

# Asia Pacific
SI_Asia <- lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change + 
                ExchRate_CHN + CPI_CHN, SafeInduRegion_Train)
# sink("SI_Asia.txt)")
summary(SI_Asia)
# sink()

# Europe / Middle East / Africa
SI_Eur <- lm(`Europe, Middle East, Africa` ~ ExchRate_GER, SafeInduRegion_Train)
# sink("SI_Eur.txt)")
summary(SI_Eur)
# sink()


## Division: Transportation & Electronics, look at all regions separately

TranElecRegion_Train <- 
      DivRegion_Train %>% filter(Division == "Transportation and Electronics")
TranElecRegion_Train <- cbind(TranElecRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
TE_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA, TranElecRegion_Train)
# sink("TE_Amer.txt)")
summary(TE_Amer)
# sink()

# Asia Pacific
TE_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + CPI_CHN, 
                TranElecRegion_Train)
# sink("TE_Asia.txt)")
summary(TE_Asia)
# sink()

# Europe / Middle East / Africa
TE_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER, TranElecRegion_Train)
# sink("TE_Eur.txt)")
summary(TE_Eur)
# sink()


## Division: Health Care, look at all regions separately

HealthRegion_Train <- DivRegion_Train %>% filter(Division == "Health Care")
HealthRegion_Train <- cbind(HealthRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
HC_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA + Unemp_USA_Perc + CPI_USA, 
     HealthRegion_Train)
# sink("HC_Amer.txt)")
summary(HC_Amer)
# sink()

# Asia Pacific
HC_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change, 
     HealthRegion_Train)
# sink("HC_Asia.txt)")
summary(HC_Asia)
# sink()

# Europe / Middle East / Africa
HC_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER + GDP_GER_Perc_Change, 
        HealthRegion_Train)
# sink("HC_Eur.txt)")
summary(HC_Eur)
# sink()



## Division: Consumer, look at all regions separately 

ConsRegion_Train <- DivRegion_Train %>% filter(Division == "Consumer")
ConsRegion_Train <- cbind(ConsRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
C_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA, ConsRegion_Train)
# sink("C_Amer.txt)")
summary(C_Amer)
# sink()

# Asia Pacific
C_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change + 
       ExchRate_CHN, ConsRegion_Train)
# sink("C_Asia.txt)")
summary(C_Asia)
# sink()

# Europe / Middle East / Africa
C_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER + Unemp_GER_Perc + ConsBaro_GER, 
     ConsRegion_Train)
# sink("C_Eur.txt)")
summary(C_Eur)
# sink()



### 4. Training Procedure to forecast the matrix element sales


## Interpolation from 2013 to 2016

# Avg. percentage of total sales of each regions per  division
names(Division_Region_Matrix_Sales)[1] <- "Division"
DivRegionInterpol <- Division_Region_Matrix_Sales %>% 
  filter(!grepl("Total Company", Division))
DivRegionInterpol <- DivRegionInterpol %>% 
  filter(!grepl("Elimination of Dual Credit", Division))
DivRegionInterpol <- DivRegionInterpol %>% 
  filter(!grepl("Corporate and Unallocated", Division))
DivRegionInterpol <- DivRegionInterpol[,-6]

DivRegionInterpol$Perc_Amer <- 
          DivRegionInterpol$Americas/DivRegionInterpol$Worldwide
DivRegionInterpol$Perc_Asia <- 
          DivRegionInterpol$`Asia Pacific`/DivRegionInterpol$Worldwide
DivRegionInterpol$Perc_Eur <- 
    DivRegionInterpol$`Europe, Middle East, Africa`/DivRegionInterpol$Worldwide

SIAmer_Perc <- DivRegionInterpol[,c(1,7)] %>% 
                  filter(Division == "Safety & Industrial") %>%
                  summarise(mean = mean(Perc_Amer))

SIAsia_Perc <- DivRegionInterpol[,c(1,8)] %>% 
  filter(Division == "Safety & Industrial") %>%
  summarise(mean = mean(Perc_Asia))

SIEur_Perc <- DivRegionInterpol[,c(1,9)] %>% 
  filter(Division == "Safety & Industrial") %>%
  summarise(mean = mean(Perc_Eur))

TEAmer_Perc <- DivRegionInterpol[,c(1,7)] %>% 
  filter(Division == "Transportation and Electronics") %>%
  summarise(mean = mean(Perc_Amer))

TEAsia_Perc <- DivRegionInterpol[,c(1,8)] %>% 
  filter(Division == "Transportation and Electronics") %>%
  summarise(mean = mean(Perc_Asia))

TEEur_Perc <- DivRegionInterpol[,c(1,9)] %>% 
  filter(Division == "Transportation and Electronics") %>%
  summarise(mean = mean(Perc_Eur))

HCAmer_Perc <- DivRegionInterpol[,c(1,7)] %>% 
  filter(Division == "Health Care") %>%
  summarise(mean = mean(Perc_Amer))

HCAsia_Perc <- DivRegionInterpol[,c(1,8)] %>% 
  filter(Division == "Health Care") %>%
  summarise(mean = mean(Perc_Asia))

HCEur_Perc <- DivRegionInterpol[,c(1,9)] %>% 
  filter(Division == "Health Care") %>%
  summarise(mean = mean(Perc_Eur))

CAmer_Perc <- DivRegionInterpol[,c(1,7)] %>% 
  filter(Division == "Consumer") %>%
  summarise(mean = mean(Perc_Amer))

CAsia_Perc <- DivRegionInterpol[,c(1,8)] %>% 
  filter(Division == "Consumer") %>%
  summarise(mean = mean(Perc_Asia))

CEur_Perc <- DivRegionInterpol[,c(1,9)] %>% 
  filter(Division == "Consumer") %>%
  summarise(mean = mean(Perc_Eur))

DivRegionInterpol <- DivRegionInterpol[,1:6]


# Fill the table with the divisional names and total sales
Temp <- Sales_per_Division[c(1:4), c(1,24:39)]
Temp[5:64,] <- NA
j=0
x=5
y=8
for (i in 1:15){
    
    Temp[x:y,1] <- Temp[1:4,1]
    Temp[x:y,2] <- Temp[1:4,2+i]
    j=j+4
    
    x=5+j
    y=8+j
}

Temp <- Temp[,1:2]
names(Temp)[1] <- "Division"
names(Temp)[2] <- "Worldwide"

# Columns for the regional sales per division
Temp$Period <- NA
Temp$Americas <- NA
Temp$`Asia Pacific` <- NA
Temp$`Europe, Middle East, Africa` <- NA


# Compute the divisional and regional sales with the interpolated mean
j=0
for (i in 1:16){
  
  Temp[1+j,4] <- Temp[1+j,2]*SIAmer_Perc[1,1]
  Temp[2+j,4] <- Temp[2+j,2]*TEAmer_Perc[1,1]
  Temp[3+j,4] <- Temp[3+j,2]*HCAmer_Perc[1,1]
  Temp[4+j,4] <- Temp[4+j,2]*CAmer_Perc[1,1]
  
  Temp[1+j,5] <- Temp[1+j,2]*SIAsia_Perc[1,1]
  Temp[2+j,5] <- Temp[2+j,2]*TEAsia_Perc[1,1]
  Temp[3+j,5] <- Temp[3+j,2]*HCAsia_Perc[1,1]
  Temp[4+j,5] <- Temp[4+j,2]*CAsia_Perc[1,1]
  
  Temp[1+j,6] <- Temp[1+j,2]*SIEur_Perc[1,1]
  Temp[2+j,6] <- Temp[2+j,2]*TEEur_Perc[1,1]
  Temp[3+j,6] <- Temp[3+j,2]*HCEur_Perc[1,1]
  Temp[4+j,6] <- Temp[4+j,2]*CEur_Perc[1,1]
  j=j+4
}

# Insert the correct quarter
Temp[1:4,3] <- paste("Q4", 2016, sep = "_")
Temp[5:8,3] <- paste("Q3", 2016, sep = "_")
Temp[9:12,3] <- paste("Q2", 2016, sep = "_")
Temp[13:16,3] <- paste("Q1", 2016, sep = "_")
Temp[17:20,3] <- paste("Q4", 2015, sep = "_")
Temp[21:24,3] <- paste("Q3", 2015, sep = "_")
Temp[25:28,3] <- paste("Q2", 2015, sep = "_")
Temp[29:32,3] <- paste("Q1", 2015, sep = "_")
Temp[33:36,3] <- paste("Q4", 2014, sep = "_")
Temp[37:40,3] <- paste("Q3", 2014, sep = "_")
Temp[41:44,3] <- paste("Q2", 2014, sep = "_")
Temp[45:48,3] <- paste("Q1", 2014, sep = "_")
Temp[49:52,3] <- paste("Q4", 2013, sep = "_")
Temp[53:56,3] <- paste("Q3", 2013, sep = "_")
Temp[57:60,3] <- paste("Q2", 2013, sep = "_")
Temp[61:64,3] <- paste("Q1", 2013, sep = "_")

Temp <- Temp[rev(rownames(Temp)),]
Temp <- Temp %>% subset(select = c(1,3,4,5,6,2))


# Match the divisional names with the original table
Temp$Division[Temp$Division == "Total Consumer Business Group"] <- 
            "Consumer"
Temp$Division[Temp$Division == "Total Health Care Business Group"] <- 
            "Health Care"
Temp$Division[Temp$Division == "Total Transportation and Electronics Business Segment"] <- 
            "Transportation and Electronics"
Temp$Division[Temp$Division == "Total Safety and Industrial Business Segment"] <- 
            "Safety & Industrial"

# Connect the original and the interpolated data set
DivRegionInterpol <- rbind(Temp, DivRegionInterpol)


## Data frame separation into the interpolated time series

Interpol_SIAmer <- 
  DivRegionInterpol[,1:3] %>% filter(Division == "Safety & Industrial")
names(Interpol_SIAmer)[3] <- "Region"
Interpol_TEAmer <- 
  DivRegionInterpol[,1:3] %>% 
  filter(Division == "Transportation and Electronics")
names(Interpol_TEAmer)[3] <- "Region"
Interpol_HCAmer <- 
  DivRegionInterpol[,1:3] %>% filter(Division == "Health Care")
names(Interpol_HCAmer)[3] <- "Region"
Interpol_CAmer <- 
  DivRegionInterpol[,1:3] %>% filter(Division == "Consumer")
names(Interpol_CAmer)[3] <- "Region"
Interpol_SIAsia <- 
  DivRegionInterpol[,c(1:2,4)] %>% filter(Division == "Safety & Industrial")
names(Interpol_SIAsia)[3] <- "Region"
Interpol_TEAsia <- 
  DivRegionInterpol[,c(1:2,4)] %>% 
  filter(Division == "Transportation and Electronics")
names(Interpol_TEAsia)[3] <- "Region"
Interpol_HCAsia <- 
  DivRegionInterpol[,c(1:2,4)] %>% filter(Division == "Health Care")
names(Interpol_HCAsia)[3] <- "Region"
Interpol_CAsia <- 
  DivRegionInterpol[,c(1:2,4)] %>% filter(Division == "Consumer")
names(Interpol_CAsia)[3] <- "Region"
Interpol_SIEur <- 
  DivRegionInterpol[,c(1:2,5)] %>% filter(Division == "Safety & Industrial")
names(Interpol_SIEur)[3] <- "Region"
Interpol_TEEur <- 
  DivRegionInterpol[,c(1:2,5)] %>% 
  filter(Division == "Transportation and Electronics")
names(Interpol_TEEur)[3] <- "Region"
Interpol_HCEur <- 
  DivRegionInterpol[,c(1:2,5)] %>% filter(Division == "Health Care")
names(Interpol_HCEur)[3] <- "Region"
Interpol_CEur <- 
  DivRegionInterpol[,c(1:2,5)] %>% filter(Division == "Consumer")
names(Interpol_CEur)[3] <- "Region"



### Estimation of the underlying process

## AIC / SBC to determine the reject the null of a unit root

lags_criterion_adj <- function(Data_orig, lags, div, region){
  
  # Help parameters
  z <- 1+lags
  y <- 2+lags
  Ext_Par <- Ext_Parameters_Chg[y:nrow(Data_orig),]
  last_lag <- nrow(Data_orig)-1
  
  if(region == "Americans"){
      reg_aid <- 0
    } else if (region == "European"){
      reg_aid <- 1
    } else if (region == "Asia Pacific"){
      reg_aid <- 2
    }
  
  # Different setups for matrices with or without lags
  if (lags == 0){
    X <- as.data.frame(matrix(1,last_lag,3))
    # linear trend
    # X[,2] <- 1:nrow(X)
    
    # quadratic trend
    X[,2] <- (1:nrow(X))^2
    
    X[,3] <- Data_orig[1:last_lag,1]
    X <- as.matrix(X)
  } else {
    # Function to lag the time series
    Data <- setDT(Data_orig)[, paste0('lagged', 1:last_lag):= shift(Region,1:last_lag)][]
    Data <- Data[lags+2:nrow(Data),]
    
    X <- data.frame(matrix(ncol=3, nrow=last_lag-lags))
    naming <- c("Constant", "Trend", "Lastval")
    colnames(X) <- naming
    
    # set up the X matrix for the OLS estimation
    X$Constant <- 1
    # linear trend
    # X[,2] <- 1:nrow(X)
    
    # quadratic trend
    X$Trend <- (1:nrow(X))^2
    X$Lastval <- Data_orig[c(z:last_lag),1]
    
    Data_lagged <- matrix(0, last_lag-lags, lags) %>% as.data.frame()
    Data_lagged <- Data[,2:z]-Data[,3:y]
    Data_lagged <- Data_lagged[1:nrow(X),]
    
    X <- cbind(X, Data_lagged) %>% as.matrix()
  }
  
  
  # Include different external parameters dependent on the time series
  if (div == "S&I"){
    X <- cbind(X, Ext_Par[,2+reg_aid], Ext_Par[,5+reg_aid]) # GDP + Unemployment
  } else if (div == "T&E"){
    X <- cbind(X, Ext_Par[,2+reg_aid], Ext_Par[,14+reg_aid]) # GDP + CPI
  } else if (div == "HC"){
    X <- cbind(X, Ext_Par[,2+reg_aid], Ext_Par[,20+reg_aid]) # GDP + Wage
  } else if (div == "C"){
    X <- cbind(X, Ext_Par[,2+reg_aid], Ext_Par[,23+reg_aid]) # GDP + Cons. Barometer
  } 
  
  # Compute the OLS parameters
  X_trans <- t(X) %>% as.matrix()
  First_Part <- as.matrix(inv(X_trans%*%as.matrix(X)))
  Second_Part <- as.matrix(X_trans%*%as.matrix(Data_orig[y:nrow(Data_orig),1]))
  OLS_solution  <- First_Part %*% Second_Part
  
  # Compute the parts of the information criterion
  rho_hat <- OLS_solution[3,1]
  Temp <- rep(1,ncol(X_trans))
  s_squared <- 
    1/(ncol(X_trans)-nrow(First_Part))%*%(Temp%*%as.matrix((Data_orig[y:nrow(Data_orig),1]-as.matrix(X)%*%OLS_solution)))^2
  
  std_error_OLS <- sqrt(First_Part*s_squared[1,1])
  rho_hat_std_error <- std_error_OLS[3,3]
  t_stat <- (rho_hat-1)/rho_hat_std_error
  
  # Compute both criteria
  SBC <- log(s_squared)+nrow(First_Part)/ncol(X_trans)*log(ncol(X_trans))
  AIC <- log(s_squared)+2*nrow(First_Part)/ncol(X_trans)
  
  Output_matrix <- rbind(t_stat, rho_hat,SBC,AIC)
  
  return(Output_matrix)
  }


# Collect the different values and criteria
IC <- data.frame(matrix(ncol=16, nrow=48))

for (i in 0:15){
  IC[1:4,i+1] <- lags_criterion_adj(Interpol_SIAmer[,3], i, "S&I","Americans")
  IC[5:8,i+1] <- lags_criterion_adj(Interpol_SIAsia[,3], i, "S&I", "Asia Pacific")
  IC[9:12,i+1] <- lags_criterion_adj(Interpol_SIEur[,3], i, "S&I", "European")
  IC[13:16,i+1] <- lags_criterion_adj(Interpol_TEAmer[,3], i, "T&E", "Americans")
  IC[17:20,i+1] <- lags_criterion_adj(Interpol_TEAsia[,3], i, "T&E", "Asia Pacific")
  IC[21:24,i+1] <- lags_criterion_adj(Interpol_TEEur[,3], i, "T&E", "European")
  IC[25:28,i+1] <- lags_criterion_adj(Interpol_HCAmer[,3], i, "HC", "Americans")
  IC[29:32,i+1] <- lags_criterion_adj(Interpol_HCAsia[,3], i, "HC", "Asia Pacific")
  IC[33:36,i+1] <- lags_criterion_adj(Interpol_HCEur[,3], i, "HC", "European")
  IC[37:40,i+1] <- lags_criterion_adj(Interpol_CAmer[,3], i, "C", "Americans")
  IC[41:44,i+1] <- lags_criterion_adj(Interpol_CAsia[,3], i, "C", "Asia Pacific")
  IC[45:48,i+1] <- lags_criterion_adj(Interpol_CEur[,3], i, "C", "European")
}

# Find the lowest value in each row to fulfill the criterion
# Besides the C Europe, the criterion always decide for the same lags
# There, the SBC (because simpler) is chosen!

IC_min <- as.matrix(apply(IC,1,which.min))
IC_SBC <- as.data.frame(IC_min[c(3,7,11,15,19,23,27,31,35,39,43,47)])


# T-Stats and unit root proof
j=1
stats_check <- data.frame(matrix(ncol=2, nrow=12))

for (i in 1:12){
  stats_check[i,1] <- IC[j,IC_SBC[i,]]
  stats_check[i,2] <- IC[j+1,IC_SBC[i,]]
  j=j+4
  
}


## Coefficient estimation with changing coefficients for the constant and 
## the second external factor

coeff_estimate <- function(SI, TE, HC, C, lagged_vals, region, SBC){
  
  z <- lagged_vals
  y <- 1+lagged_vals
  length_TS <- nrow(SI)
  end_ts <- length_TS - lagged_vals
  
  Ext_Par <- Ext_Parameters_Chg[y:length_TS,]
  
  if(region == "Americans"){
    reg_aid <- 0
  } else if (region == "European"){
    reg_aid <- 1
  } else if (region == "Asia Pacific"){
    reg_aid <- 2
  }
  
  # Different setups for matrices with or without lags
  if (lagged_vals == 0){
    X <- matrix(1,length_TS*4,5) %>% as.data.frame()
    
    # Plug in the auxiliary variables for the different constants
    X[1:end_ts,2] <- 0
    X[(end_ts*2+1):nrow(X),2]<- 0
    
    X[1:(end_ts*2),3] <- 0
    X[(end_ts*3+1):nrow(X),3]<- 0
     
    X[1:(end_ts*3),4] <- 0 
      
    # linear trend
    # X[,5] <- 1:length_TS
    
    # quadratic trend
    X[,5] <- (1:length_TS)^2
    X <- as.matrix(X)
    
    y_exp <- rbind(SI, TE, HC, C)
    
  } else {
    # Function to lag the four time series
    Data1 <- setDT(SI)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data1 <- Data1[lagged_vals+1:nrow(Data1),] %>% as.data.frame()
    
    Data2 <- setDT(TE)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data2 <- Data2[lagged_vals+1:nrow(Data2),] %>% as.data.frame()
    
    Data3 <- setDT(HC)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data3 <- Data3[lagged_vals+1:nrow(Data3),] %>% as.data.frame()
    
    Data4 <- setDT(C)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data4 <- Data4[lagged_vals+1:nrow(Data4),] %>% as.data.frame()
    
    # Y-vector for OLS
    y_exp <- cbind(t(Data1[1:(length_TS-z),1]), t(Data2[1:(length_TS-z),1]), t(Data3[1:(length_TS-z),1]), 
                   t(Data4[1:(length_TS-z),1])) %>% t()
    
    X <- as.data.frame(matrix(1,end_ts*4,5+lagged_vals))
    naming <- c("Constant1", "Aux_Constant2", "Aux_Constant3", "Aux_Constant4", "Trend")
    colnames(X) <- naming
    
    # Plug in the auxiliary variables for the different constants
    X[1:end_ts,2] <- 0
    X[(end_ts*2+1):nrow(X),2]<- 0
    
    X[1:(end_ts*2),3] <- 0
    X[(end_ts*3+1):nrow(X),3]<- 0
    
    X[1:(end_ts*3),4] <- 0 
    # linear trend
    # X[,5] <- 1:nrow(X)
    
    # quadratic trend
    X[,5] <- (1:end_ts)^2
    
    for (i in 1:lagged_vals){
      X[1:end_ts,5+i] <- Data1[1:end_ts,1+i]
      X[(end_ts+1):(2*end_ts),5+i] <- Data2[1:end_ts,1+i]
      X[(2*end_ts+1):(3*end_ts),5+i] <- Data3[1:end_ts,1+i]
      X[(3*end_ts+1):(4*end_ts),5+i] <- Data4[1:end_ts,1+i]
    }
  }
  
  # Include GDP for every time series
  X <- cbind(X, Ext_Par[,2+reg_aid])
  
  # Row-bind the other external parameters for the time series
  Ext_Matrix <- data.frame(matrix(0,ncol=4, nrow=end_ts*4))
  Ext_Matrix[1:end_ts,1] <- Ext_Par[,5+reg_aid]
  Ext_Matrix[(end_ts+1):(end_ts*2),2]<- Ext_Par[,14+reg_aid]
  Ext_Matrix[(end_ts*2+1):(end_ts*3),3] <- Ext_Par[,20+reg_aid]
  Ext_Matrix[(end_ts*3+1):nrow(Ext_Matrix),4] <- Ext_Par[,23+reg_aid]
    
  X <- cbind(X, Ext_Matrix)
  
  
  # OLS coefficients
  coefficients_OLS <- data.frame(matrix(ncol=1, nrow=10+lagged_vals))
  
  # Collect the residuals for the SBC
  residuals <- data.frame(matrix(ncol=1, nrow=end_ts*4))
  
  X_trans <- t(X) %>% as.matrix()
  First_Part <- as.matrix(inv(X_trans%*%as.matrix(X)))
  Second_Part <- as.matrix(X_trans%*%as.matrix(y_exp))
  coefficients_OLS  <- First_Part %*% Second_Part
    
  residuals <- y_exp - as.matrix(X)%*%coefficients_OLS
  
  # SBC for the optimal model specification
  SBC_est <- 
    log((sd(residuals[,1]))^2)+lagged_vals*log(end_ts)*(1/end_ts)
  
  
  if (SBC == FALSE){
    return(coefficients_OLS)
  } else {
    return(SBC_est)
  }
}


## Estimate the AR-model specification for the different regions

SBC <- data.frame(matrix(ncol=3, nrow=19))
names(SBC)[1] <- "Americans"
names(SBC)[2] <- "European"
names(SBC)[3] <- "Asia Pacific"

for (i in 0:18){
  SBC[i+1,1] <- coeff_estimate(Interpol_SIAmer[,3], Interpol_TEAmer[,3], 
                 Interpol_HCAmer[,3], Interpol_CAmer[,3], i, "Americans", TRUE)
  SBC[i+1,2] <- coeff_estimate(Interpol_SIEur[,3], Interpol_TEEur[,3], 
                 Interpol_HCEur[,3], Interpol_CEur[,3], i, "European", TRUE)
  SBC[i+1,3] <- coeff_estimate(Interpol_SIAsia[,3], Interpol_TEAsia[,3], 
                 Interpol_HCAsia[,3], Interpol_CAsia[,3], i, "Asia Pacific", TRUE)
}


## Collect the coefficients of the AR-models
pred_coeffs <- data.frame(matrix(ncol=3, nrow=11))

pred_coeffs[1:10,1] <- coeff_estimate(Interpol_SIAmer[,3], Interpol_TEAmer[,3], 
                Interpol_HCAmer[,3], Interpol_CAmer[,3], 0, "Americans", FALSE)
pred_coeffs[1:10,2] <- coeff_estimate(Interpol_SIEur[,3], Interpol_TEEur[,3], 
                Interpol_HCEur[,3], Interpol_CEur[,3], 0, "European", FALSE)
pred_coeffs[1:11,3] <- coeff_estimate(Interpol_SIAsia[,3], Interpol_TEAsia[,3], 
                Interpol_HCAsia[,3], Interpol_CAsia[,3], 1, "Asia Pacific", FALSE)



### Predict the upcoming sales of all divisional and regional matrix elements

Sales_Hist <- 
  cbind(Interpol_SIAmer[,3], Interpol_SIAsia[,3], Interpol_SIEur[,3], 
        Interpol_TEAmer[,3], Interpol_TEAsia[,3], Interpol_TEEur[,3],
        Interpol_HCAmer[,3], Interpol_HCAsia[,3], Interpol_HCEur[,3],
        Interpol_CAmer[,3], Interpol_CAsia[,3], Interpol_CEur[,3])

names(Sales_Hist)[1] <- "S&I Americans"
names(Sales_Hist)[2] <- "S&I Asia Pacific"
names(Sales_Hist)[3] <- "S&I European"
names(Sales_Hist)[4] <- "T&E Americans"
names(Sales_Hist)[5] <- "T&E Asia Pacific"
names(Sales_Hist)[6] <- "T&E European"
names(Sales_Hist)[7] <- "HC Americans"
names(Sales_Hist)[8] <- "HC Asia Pacific"
names(Sales_Hist)[9] <- "HC European"
names(Sales_Hist)[10] <- "C Americans"
names(Sales_Hist)[11] <- "C Asia Pacific"
names(Sales_Hist)[12] <- "C European"


## Set-up of function to predict the sales

Sales_Pred <- function(Coeff, Prev_val, Externals, Lags, Pred_div){
  
    if(Pred_div =="SI"){
      if(Lags == 0){
      SI_pred <- 
        as.numeric(Coeff[1]+Coeff[5]+Coeff[6]*Externals[1,1]+Coeff[7]*Externals[1,2])
      } else if (Lags == 1){
      SI_pred <- 
        as.numeric(Coeff[1]+Coeff[5]+Coeff[6]*Prev_val+Coeff[7]*Externals[1,1]+Coeff[8]*Externals[1,2])
      }
      
      return(SI_pred)
      
    } else if (Pred_div == "TE"){
      if (Lags==0){
      TE_pred <- 
        as.numeric(Coeff[1]+Coeff[2]+Coeff[5]+Coeff[6]*Externals[1,1]+Coeff[8]*Externals[1,2])
      } else if (Lags == 1){
      TE_pred <- 
        as.numeric(Coeff[1]+Coeff[2]+Coeff[5]+Coeff[6]*Prev_val+Coeff[7]*Externals[1,1]+Coeff[9]*Externals[1,2])
      }
      
      return(TE_pred)
      
    } else if (Pred_div == "HC"){
      if (Lags ==0){
        HC_pred <- 
          as.numeric(Coeff[1]+Coeff[3]+Coeff[5]+Coeff[6]*Externals[1,1]+Coeff[9]*Externals[1,2])
      } else if (Lags ==1){
        HC_pred <- 
          as.numeric(Coeff[1]+Coeff[3]+Coeff[5]+Coeff[6]*Prev_val+Coeff[7]*Externals[1,1]+Coeff[10]*Externals[1,2])
      }
      
      return(HC_pred)
      
    } else if (Pred_div == "C"){
      if (Lags ==0){
       C_pred <- 
          as.numeric(Coeff[1]+Coeff[4]+Coeff[5]+Coeff[6]*Externals[1,1]+Coeff[10]*Externals[1,2])
      } else if (Lags ==1){
      C_pred <- 
        as.numeric(Coeff[1]+Coeff[4]+Coeff[5]+Coeff[6]*Prev_val+Coeff[7]*Externals[1,1]+Coeff[11]*Externals[1,2])
      }
      return(C_pred)
    }
}


## Predict the sales for the next quarters

Temp <- data.frame(matrix(ncol=12, nrow=1))
names(Temp) <- names(Sales_Hist)

# Loop over the next four quarters
for (i in 1:4){
  
    # Only the external parameters for the first and second quarter given
    # Second quarter parameters are assumed to be stable for third/fourth quarter
    if (i==1){
      y <- Ext_Parameters_Chg[37,]
    } else {
      y <- Ext_Parameters_Chg[38,]
    }
  
  # Predict the sales for each time series
  Temp[,1] <- 
    Sales_Pred(pred_coeffs[1:10,1], Sales_Hist[nrow(Sales_Hist),1], y[,c(2,5)],0,"SI")
  
  Temp[,2] <- 
    Sales_Pred(pred_coeffs[1:11,3], Sales_Hist[nrow(Sales_Hist),2],y[,c(4,7)],1,"SI")

  Temp[,3] <- 
    Sales_Pred(pred_coeffs[1:10,2], Sales_Hist[nrow(Sales_Hist),3],y[,c(3,6)],0,"SI")
  
  Temp[,4] <- 
    Sales_Pred(pred_coeffs[1:10,1], Sales_Hist[nrow(Sales_Hist),4],y[,c(2,14)],0,"TE")
  
  Temp[,5] <- 
    Sales_Pred(pred_coeffs[1:11,3], Sales_Hist[nrow(Sales_Hist),5], y[,c(4,16)],1,"TE")
  
  Temp[,6] <- 
    Sales_Pred(pred_coeffs[1:10,2], Sales_Hist[nrow(Sales_Hist),6],y[,c(3,15)],0,"TE")
  
  Temp[,7] <- 
    Sales_Pred(pred_coeffs[1:10,1], Sales_Hist[nrow(Sales_Hist),7],y[,c(2,20)],0,"HC")
  
  Temp[,8] <- 
    Sales_Pred(pred_coeffs[1:11,3], Sales_Hist[nrow(Sales_Hist),8],y[,c(4,22)],1,"HC")
  
  Temp[,9] <- 
    Sales_Pred(pred_coeffs[1:10,2], Sales_Hist[nrow(Sales_Hist),9], y[,c(3,21)],0,"HC")
  
  Temp[,10] <- 
    Sales_Pred(pred_coeffs[1:10,1], Sales_Hist[nrow(Sales_Hist),10],y[,c(2,23)],0,"C")
  
  Temp[,11] <- 
    Sales_Pred(pred_coeffs[1:11,3], Sales_Hist[nrow(Sales_Hist),11],y[,c(4,25)],1,"C")
  
  Temp[,12] <-
    Sales_Pred(pred_coeffs[1:10,2], Sales_Hist[nrow(Sales_Hist),12],y[,c(3,24)],0,"C")
  
  
  # Add the projected sales in order to estimate the next quarter
  Sales_Hist <- rbind(Sales_Hist, Temp)
}



################################################################################
###                       Final Rolling Forecast Function                    ###
################################################################################


### Assumptions about the future development of external parameters

Exp_Ext_Parameter <-
  data.frame(Q3_22 = c(0.003, 0.006, 0.002, -0.0015, 0.001, -0.006, 0.015, 0.012, 
                       0.028, 0.011, 0.007, 0.002, -0.005, -0.001, -0.009),
             Q4_22 = c(0.005, 0.009, 0.003, -0.001, -0.003, 0.001, 0.013, 0.011,
                       0.023, 0.010, 0.004, 0.0115, -0.002, 0, -0.0065),
             Q1_23 = c(0.003, 0.007, 0.003, -0.0007, 0, 0.001, 0.011, 0.007,
                       0.015, 0.008, 0, 0.0135, 0.001, 0.0015, -0.004),
             Q2_23 = c(0.004, 0.009, 0.004, -0.0005, -0.003, -0.001, 0.009,
                       0.004, 0.012, 0.006, 0, -0.002, 0.003, 0.003, -0.0025))


### Interpolated Sales History

Sales_History_Q2 <- 
  cbind(Interpol_SIAmer[,3], Interpol_TEAmer[,3], Interpol_HCAmer[,3], Interpol_CAmer[,3],
        Interpol_SIEur[,3], Interpol_TEEur[,3], Interpol_HCEur[,3], Interpol_CEur[,3],
        Interpol_SIAsia[,3], Interpol_TEAsia[,3], Interpol_HCAsia[,3], Interpol_CAsia[,3])


## Row-bind the interpolated series of Q1 and Q2 2022

Data22 <- data.frame(c(1515,1452), c(761,737), c(1221,1252), c(919,931),
                     c(690, 661), c(372,360), c(495,508), c(143,145),
                     c(847,812), c(1208,1171), c(408,419), c(251,254))

names(Data22) <- names(Sales_History_Q2)
Sales_History_Q2 <- rbind(Sales_History_Q2, Data22)
names(Sales_History_Q2)[1:12] <- "Region"

### Rolling Forecast Function

rolling_forecast <- function(Sales_History, region, External_Expectations){
  
  length_TS <- nrow(Sales_History)
  
  if(region == "Americans"){
    reg_aid <- 0
    reg_ext <- 0
    reg_forecast <- 0
    SI <- Sales_History[,1] %>% as.data.frame()
    TE <- Sales_History[,2] %>% as.data.frame()
    HC <- Sales_History[,3] %>% as.data.frame()
    C <- Sales_History[,4] %>% as.data.frame()
    names(SI) <- "Region"
  } else if (region == "European"){
    reg_aid <- 1
    reg_ext <- 2
    reg_forecast <- 4
    SI <- Sales_History[,5] %>% as.data.frame()
    TE <- Sales_History[,6] %>% as.data.frame()
    HC <- Sales_History[,7] %>% as.data.frame()
    C <- Sales_History[,8] %>% as.data.frame()
  } else if (region == "Asia Pacific"){
    reg_aid <- 2
    reg_ext <- 1
    reg_forecast <- 8
    SI <- Sales_History[,9] %>% as.data.frame()
    TE <- Sales_History[,10] %>% as.data.frame()
    HC <- Sales_History[,11] %>% as.data.frame()
    C <- Sales_History[,12] %>% as.data.frame()
  }
  
  names(SI) <- "Region"
  names(TE) <- "Region"
  names(HC) <- "Region"
  names(C) <- "Region"
  
  for (i in 0:10){
    
  z <- i
  y <- i+1
  end_ts <- length_TS - i
    
  Ext_Par <- Ext_Parameters_Chg[y:length_TS,]  
  
  # Different setups for matrices with or without lags
  if (i == 0){
    X <- matrix(1,length_TS*4,5) %>% as.data.frame()
    
    # Plug in the auxiliary variables for the different constants
    X[1:end_ts,2] <- 0
    X[(end_ts*2+1):nrow(X),2]<- 0
    
    X[1:(end_ts*2),3] <- 0
    X[(end_ts*3+1):nrow(X),3]<- 0
    
    X[1:(end_ts*3),4] <- 0 
    
    # quadratic trend
    X[,5] <- (1:length_TS)^2
    X <- as.matrix(X)
    
    # Y-vector for OLS
    if(region == "Americans"){
          y_exp <- c(Sales_History[,1], Sales_History[,2], Sales_History[,3], Sales_History[,4])
    } else if (region == "European"){
          y_exp <- c(Sales_History[,5], Sales_History[,6], Sales_History[,7], Sales_History[,8])
    } else if (region == "Asia Pacific"){
          y_exp <- c(Sales_History[,9], Sales_History[,10], Sales_History[,11], Sales_History[,12])
    }
    
  } else {
    # Function to lag the four time series
    Data1 <- setDT(SI)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data1 <- Data1[(i+1):nrow(Data1),] %>% as.data.frame()
    
    Data2 <- setDT(TE)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data2 <- Data2[(i+1):nrow(Data2),] %>% as.data.frame()
    
    Data3 <- setDT(HC)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data3 <- Data3[(i+1):nrow(Data3),] %>% as.data.frame()
    
    Data4 <- setDT(C)[, paste0('lagged', 1:end_ts):= shift(Region,1:end_ts)][]
    Data4 <- Data4[(i+1):nrow(Data4),] %>% as.data.frame()
    
    # Y-vector for OLS
    y_exp <- cbind(t(Data1[1:(length_TS-z),1]), t(Data2[1:(length_TS-z),1]), t(Data3[1:(length_TS-z),1]), 
                   t(Data4[1:(length_TS-z),1])) %>% t()
    
    X <- as.data.frame(matrix(1,end_ts*4,5+i))
    naming <- c("Constant1", "Aux_Constant2", "Aux_Constant3", "Aux_Constant4", "Trend")
    colnames(X) <- naming
    
    # Plug in the auxiliary variables for the different constants
    X[1:end_ts,2] <- 0
    X[(end_ts*2+1):nrow(X),2]<- 0
    
    X[1:(end_ts*2),3] <- 0
    X[(end_ts*3+1):nrow(X),3]<- 0
    
    X[1:(end_ts*3),4] <- 0 

    # quadratic trend
    X[,5] <- (1:end_ts)^2
    
    for (j in 1:i){
      X[1:end_ts,5+j] <- Data1[1:end_ts,1+j]
      X[(end_ts+1):(2*end_ts),5+j] <- Data2[1:end_ts,1+j]
      X[(2*end_ts+1):(3*end_ts),5+j] <- Data3[1:end_ts,1+j]
      X[(3*end_ts+1):(4*end_ts),5+j] <- Data4[1:end_ts,1+j]
    }
  }
  
  # Include GDP for every time series
  X <- cbind(X, Ext_Par[,2+reg_aid])
  
  # Row-bind the other external parameters for the time series
  Ext_Matrix <- data.frame(matrix(0,ncol=4, nrow=end_ts*4))
  Ext_Matrix[1:end_ts,1] <- Ext_Par[,5+reg_aid]
  Ext_Matrix[(end_ts+1):(end_ts*2),2]<- Ext_Par[,14+reg_aid]
  Ext_Matrix[(end_ts*2+1):(end_ts*3),3] <- Ext_Par[,20+reg_aid]
  Ext_Matrix[(end_ts*3+1):nrow(Ext_Matrix),4] <- Ext_Par[,23+reg_aid]
  
  X <- cbind(X, Ext_Matrix)
  
  
  # OLS coefficients
  coefficients_OLS <- data.frame(matrix(ncol=1, nrow=10+i))
  
  # Collect the residuals for the SBC
  residuals <- data.frame(matrix(ncol=1, nrow=end_ts*4))
  
  X_trans <- t(X) %>% as.matrix()
  First_Part <- as.matrix(inv(X_trans%*%as.matrix(X)))
  Second_Part <- as.matrix(X_trans%*%as.matrix(y_exp))
  coefficients_OLS  <- First_Part %*% Second_Part
  
  residuals <- y_exp - as.matrix(X)%*%coefficients_OLS
  
  # SBC for the optimal model specification
  SBC_est <- 
    log((sd(residuals[,1]))^2)+i*log(end_ts)*(1/end_ts)
  
  if(i==0){
    SBC_min <- SBC_est
    lags <- 0
    OLS_solution <- data.frame(matrix(ncol=1, nrow=20))
    OLS_solution[1:10,] <- coefficients_OLS
  }
  
  if(SBC_est < SBC_min){
    SBC_min <- SBC_est
    lags <- i
    OLS_solution[1:(10+i),] <- coefficients_OLS
  }
  
  }
  
  # Collect the forecasts
  Temp <- data.frame(matrix(0, ncol=12, nrow=1))
  names(Temp) <- names(Sales_History)
  
  Forecast_estimate <- data.frame(matrix(nrow = 1, ncol=4))
  
  for(k in 1:4){
  
  if(lags == 0){
    Forecast_estimate[,1] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[7,1]*External_Expectations[(4+reg_ext),k])
    Forecast_estimate[,2] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[2,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[8,1]*External_Expectations[(7+reg_ext),k])
    Forecast_estimate[,3] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[3,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[9,1]*External_Expectations[(10+reg_ext),k])
    Forecast_estimate[,4] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[4,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[10,1]*External_Expectations[(13+reg_ext),k])
  } else if (lags == 1){
    Forecast_estimate[,1] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(1+reg_forecast)]+
                   OLS_solution[7,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[8,1]*External_Expectations[(4+reg_ext),k])
    Forecast_estimate[,2] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[2,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(2+reg_forecast)]+
                   OLS_solution[7,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[9,1]*External_Expectations[(7+reg_ext),k])
    Forecast_estimate[,3] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[3,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(3+reg_forecast)]+
                   OLS_solution[7,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[10,1]*External_Expectations[(10+reg_ext),k])
    Forecast_estimate[,4] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[4,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(4+reg_forecast)]+
                   OLS_solution[7,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[11,1]*External_Expectations[(13+reg_ext),k])
  } else if (lags == 2){
    Forecast_estimate[,1] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(1+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(1+reg_forecast)]+
                   OLS_solution[8,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[9,1]*External_Expectations[(4+reg_ext),k])
    Forecast_estimate[,2] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[2,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(2+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(2+reg_forecast)]+
                   OLS_solution[8,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[10,1]*External_Expectations[(7+reg_ext),k])
    Forecast_estimate[,3] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[3,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(3+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(3+reg_forecast)]+
                   OLS_solution[8,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[11,1]*External_Expectations[(10+reg_ext),k])
    Forecast_estimate[,4] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[4,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(4+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(4+reg_forecast)]+
                   OLS_solution[8,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[12,1]*External_Expectations[(13+reg_ext),k])
  } else if (lags == 3){
    Forecast_estimate[,1] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(1+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(1+reg_forecast)]+
                   OLS_solution[8,1]*Sales_History[(nrow(Sales_History)-2),(1+reg_forecast)]+
                   OLS_solution[9,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[10,1]*External_Expectations[(4+reg_ext),k])
    Forecast_estimate[,2] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[2,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(2+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(2+reg_forecast)]+
                   OLS_solution[8,1]*Sales_History[(nrow(Sales_History)-2),(2+reg_forecast)]+
                   OLS_solution[9,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[11,1]*External_Expectations[(7+reg_ext),k])
    Forecast_estimate[,3] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[3,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(3+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(3+reg_forecast)]+
                   OLS_solution[8,1]*Sales_History[(nrow(Sales_History)-2),(3+reg_forecast)]+
                   OLS_solution[9,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[12,1]*External_Expectations[(10+reg_ext),k])
    Forecast_estimate[,4] <- 
      as.numeric(OLS_solution[1,1]+OLS_solution[4,1]+OLS_solution[5,1]+
                   OLS_solution[6,1]*Sales_History[nrow(Sales_History),(4+reg_forecast)]+
                   OLS_solution[7,1]*Sales_History[(nrow(Sales_History)-1),(4+reg_forecast)]+
                   OLS_solution[8,1]*Sales_History[(nrow(Sales_History)-2),(4+reg_forecast)]+
                   OLS_solution[9,1]*External_Expectations[(1+reg_ext),k]+
                   OLS_solution[13,1]*External_Expectations[(13+reg_ext),k])
  }
    if(region == "Americans"){
      Temp[,1:4] <- Forecast_estimate
    } else if (region == "European"){
      Temp[,5:8] <- Forecast_estimate
    } else if (region == "Asia Pacific"){
      Temp[,9:12] <- Forecast_estimate
    }
    
    Sales_History <- rbind(Sales_History,Temp)
    
  }
  
  if(region == "Americans"){
        return(Sales_History[39:42,1:4])
  } else if (region == "European"){
        return(Sales_History[39:42,5:8])
  } else if (region == "Asia Pacific"){
        return(Sales_History[39:42,9:12])
  }
}

### Forecasts Q3 '22 to Q2 '23

Rolling_Sales_Forecast <- data.frame(matrix(ncol=4, nrow=4))

Rolling_Sales_Forecast[,1:4] <- 
  rolling_forecast(Sales_History_Q2, "Americans", External_Expectations)
Rolling_Sales_Forecast[,5:8] <- 
  rolling_forecast(Sales_History_Q2, "European", External_Expectations)
Rolling_Sales_Forecast[,9:12] <- 
  rolling_forecast(Sales_History_Q2, "Asia Pacific", External_Expectations)





