################################################################################
###               Empirical Simulation in the Master Thesis                  ###
### The Quest for a Modern and Dynamic Budgeting and Forecasting Approach    ###
###               By Max Kneißler (University of Tübingen)                   ###
################################################################################



### Packages for the analysis
library(extrafont)
library(ggplot2)
library(ggpubr)
library(gridExtra)
library(stringr)
library(readxl)
library(stargazer)
library(tidyverse)



### Import of the different sheets in the database
Income_Statement <- read_excel("3M_Data.xlsx", sheet = "Income_Statement")
Sales_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "Sales_per_Division")
OpIncome_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "OpIncome_per_Division")
Sales_per_Region <- read_excel("3M_Data.xlsx", sheet = "Sales_per_Region")
Division_Region_Matrix_Sales <- 
            read_excel("3M_Data.xlsx", sheet = "Division_Region_Matrix_Sales")
Ext_Parameters <- read_excel("3M_Data.xlsx", sheet = "External Parameters")




################################################################################
###             Chapter 5.2: Graphical Illustration of the Data              ###
################################################################################



### Net Sales vs. Operating Income vs. Net Income over time

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



### Net Sales vs. Operating Income in the different divisions over time

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

# Combine the two plots
ggarrange(SalesDiv, NULL, OpIncDiv, nrow=3, 
                      heights = c(1.5, 0.0001, 1.5), align = "h")

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



### Net Sales across the different regions over time

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



### Net Sales across the different regions and divisions over time

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


# Production of the separate plots
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

# Combine the four plots with equal sizes
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


### Build up the training and validation sets

Income_Statement_Train <- Income_Statement %>% subset(select=c(1,8:39))
Income_Statement_Val <- Income_Statement %>% subset(select=c(1:7))

Division_Region_Matrix_Sales <- 
  Division_Region_Matrix_Sales[rev(rownames(Division_Region_Matrix_Sales)),]
DivRegion_Train <- Division_Region_Matrix_Sales[1:112,]
DivRegion_Val <- Division_Region_Matrix_Sales[113:140,]

OpInc_Train <- OpIncome_per_Division %>% subset(select=c(1,8:39))
OpInc_Val <- OpIncome_per_Division %>% subset(select=c(1:7))

SalesDiv_Train <- Sales_per_Division %>% subset(select=c(1,8:39))
SalesDiv_Val <- Sales_per_Division %>% subset(select=c(1:7))

SalesReg_Train <- Sales_per_Region %>% subset(select=c(1,8:39))
SalesReg_Val <- Sales_per_Region %>% subset(select=c(1:7))

Ext_Parameters <- Ext_Parameters[rev(rownames(Ext_Parameters)),]
Ext_Parameters_Train <- Ext_Parameters[1:32,]
Ext_Parameters_Val <- Ext_Parameters[33:38,]


### Linear model as a starting point to define suitable external features

### 1. Training of the Regions

# Linear model to project the sales in the Americans
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

# Summary Statistics of Linear Output 
sink("BaseAmer.txt")
print(summary(ModelAmerSales))
sink()



# Linear model to project the sales in Asia Pacific
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

# Summary Statistics of Linear Output
sink("BaseAsia.txt")
summary(ModelAsiaSales)
sink()

# Linear model to project the sales in Europa / Middle East / Africa
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

# Summary Statistics of Linear Output
sink("BaseEur.txt)")
print(summary(ModelEurSales))
sink()


### 2. Training of the Divisions


# Safety & Industrial

SalesSafeInduTrain <- SalesDiv_Train[1,] %>% t() %>% as.data.frame()
SalesSafeInduTrain <- cbind(Quarter = rownames(SalesSafeInduTrain), SalesSafeInduTrain)
SalesSafeInduTrain <- SalesSafeInduTrain[-1,]
SalesSafeInduTrain <- SalesSafeInduTrain[rev(rownames(SalesSafeInduTrain)),]
rownames(SalesSafeInduTrain) <- 1:nrow(SalesSafeInduTrain)
names(SalesSafeInduTrain)[2] <- "Net_Sales"

SalesSafeInduTrain <- cbind(SalesSafeInduTrain, Ext_Parameters_Train[,2:25])
ModelSafeIndu <- lm(Net_Sales ~ . - Quarter, SalesSafeInduTrain)

# Summary Statistics of Linear Output
sink("BaseSafeIndu.txt)")
print(summary(ModelSafeIndu))
sink()


# Transportation and Electronics

SalesTransElecTrain <- SalesDiv_Train[2,] %>% t() %>% as.data.frame()
SalesTransElecTrain <- cbind(Quarter = rownames(SalesTransElecTrain), SalesTransElecTrain)
SalesTransElecTrain <- SalesTransElecTrain[-1,]
SalesTransElecTrain <- SalesTransElecTrain[rev(rownames(SalesTransElecTrain)),]
rownames(SalesTransElecTrain) <- 1:nrow(SalesTransElecTrain)
names(SalesTransElecTrain)[2] <- "Net_Sales"

SalesTransElecTrain <- cbind(SalesTransElecTrain, Ext_Parameters_Train[,2:25])
ModelTransElec <- lm(Net_Sales ~ . - Quarter, SalesTransElecTrain)

# Summary Statistics of Linear Output
sink("BaseTransElec.txt)")
print(summary(ModelTransElec))
sink()


# Transportation and Electronics

SalesTransElecTrain <- SalesDiv_Train[2,] %>% t() %>% as.data.frame()
SalesTransElecTrain <- cbind(Quarter = rownames(SalesTransElecTrain), SalesTransElecTrain)
SalesTransElecTrain <- SalesTransElecTrain[-1,]
SalesTransElecTrain <- SalesTransElecTrain[rev(rownames(SalesTransElecTrain)),]
rownames(SalesTransElecTrain) <- 1:nrow(SalesTransElecTrain)
names(SalesTransElecTrain)[2] <- "Net_Sales"

SalesTransElecTrain <- cbind(SalesTransElecTrain, Ext_Parameters_Train[,2:25])
ModelTransElec <- lm(Net_Sales ~ . - Quarter, SalesTransElecTrain)

# Summary Statistics of Linear Output
sink("BaseTransElec.txt)")
print(summary(ModelTransElec))
sink()


# Health Care

SalesHealthTrain <- SalesDiv_Train[3,] %>% t() %>% as.data.frame()
SalesHealthTrain <- cbind(Quarter = rownames(SalesHealthTrain), SalesHealthTrain)
SalesHealthTrain <- SalesHealthTrain[-1,]
SalesHealthTrain <- SalesHealthTrain[rev(rownames(SalesHealthTrain)),]
rownames(SalesHealthTrain) <- 1:nrow(SalesHealthTrain)
names(SalesHealthTrain)[2] <- "Net_Sales"

SalesHealthTrain <- cbind(SalesHealthTrain, Ext_Parameters_Train[,2:25])
ModelHealth <- lm(Net_Sales ~ . - Quarter, SalesHealthTrain)

# Summary Statistics of Linear Output
sink("BaseHealth.txt)")
print(summary(ModelHealth))
sink()


# Consumer

SalesConsTrain <- SalesDiv_Train[4,] %>% t() %>% as.data.frame()
SalesConsTrain <- cbind(Quarter = rownames(SalesConsTrain), SalesConsTrain)
SalesConsTrain <- SalesConsTrain[-1,]
SalesConsTrain <- SalesConsTrain[rev(rownames(SalesConsTrain)),]
rownames(SalesConsTrain) <- 1:nrow(SalesConsTrain)
names(SalesConsTrain)[2] <- "Net_Sales"

SalesConsTrain <- cbind(SalesConsTrain, Ext_Parameters_Train[,2:25])
ModelCons <- lm(Net_Sales ~ . - Quarter, SalesConsTrain)

# Summary Statistics of Linear Output
sink("BaseCons.txt)")
print(summary(ModelCons))
sink()



### 3. Training for Divisions and Regions separately


# Filter the 4 segments in the 3 regions
names(DivRegion_Train)[1] <- "Division"
DivRegion_Train <- DivRegion_Train[,1:5] %>% 
            filter(!grepl("Corporate and Unallocated", Division)) %>%
            filter(!grepl("Total Company", Division)) %>%
            filter(!grepl("Elimination of Dual Credit", Division))


# External Parameters only for the time horizon
DivRegExt_Paramter <- Ext_Parameters_Train[17:32,]


### Separate Division Modelling for each region


## Safety & Industrial

SafeInduRegion_Train <- 
            DivRegion_Train %>% filter(Division == "Safety & Industrial")
SafeInduRegion_Train <- cbind(SafeInduRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
SI_Amer <- lm(Americas ~ GDP_USA_Perc_Change + Unemp_USA_Perc + IntRate_USA_Perc + 
                CPI_USA + Avg_Wage_USA, SafeInduRegion_Train)
sink("SI_Amer.txt)")
summary(SI_Amer)
sink()

# Asia Pacific
SI_Asia <- lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change + 
                ExchRate_CHN + CPI_CHN, SafeInduRegion_Train)
sink("SI_Asia.txt)")
summary(SI_Asia)
sink()

# Europe / Middle East / Africa
SI_Eur <- lm(`Europe, Middle East, Africa` ~ ExchRate_GER, SafeInduRegion_Train)
sink("SI_Eur.txt)")
summary(SI_Eur)
sink()


## Transportation & Electronics

TranElecRegion_Train <- 
      DivRegion_Train %>% filter(Division == "Transportation and Electronics")
TranElecRegion_Train <- cbind(TranElecRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
TE_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA, TranElecRegion_Train)
sink("TE_Amer.txt)")
summary(TE_Amer)
sink()


# Asia Pacific
TE_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + CPI_CHN, 
                TranElecRegion_Train)
sink("TE_Asia.txt)")
summary(TE_Asia)
sink()


# Europe / Middle East / Africa
TE_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER, TranElecRegion_Train)
sink("TE_Eur.txt)")
summary(TE_Eur)
sink()



## Health Care

HealthRegion_Train <- DivRegion_Train %>% filter(Division == "Health Care")
HealthRegion_Train <- cbind(HealthRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
HC_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA + Unemp_USA_Perc + CPI_USA, 
     HealthRegion_Train)
sink("HC_Amer.txt)")
summary(HC_Amer)
sink()


# Asia Pacific
HC_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change, 
     HealthRegion_Train)
sink("HC_Asia.txt)")
summary(HC_Asia)
sink()


# Europe / Middle East / Africa
HC_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER + GDP_GER_Perc_Change, 
        HealthRegion_Train)
sink("HC_Eur.txt)")
summary(HC_Eur)
sink()



## Consumer 

ConsRegion_Train <- DivRegion_Train %>% filter(Division == "Consumer")
ConsRegion_Train <- cbind(ConsRegion_Train, DivRegExt_Paramter[,2:25])

# Americans
C_Amer <- 
  lm(Americas ~ GDP_USA_Perc_Change + Avg_Wage_USA, ConsRegion_Train)
sink("C_Amer.txt)")
summary(C_Amer)
sink()


# Asia Pacific
C_Asia <- 
  lm(`Asia Pacific` ~ Umemp_CHN_Perc + Avg_Wage_JPN + GDP_CHN_Perc_Change + 
       ExchRate_CHN, ConsRegion_Train)
sink("C_Asia.txt)")
summary(C_Asia)
sink()


# Europe / Middle East / Africa
C_Eur <- 
  lm(`Europe, Middle East, Africa` ~ ExchRate_GER + Unemp_GER_Perc + ConsBaro_GER, 
     ConsRegion_Train)
sink("C_Eur.txt)")
summary(C_Eur)
sink()




# Validation year 2021


# Compute the predictions for the next quarter

