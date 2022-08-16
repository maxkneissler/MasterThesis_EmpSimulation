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



# Validation year 2021


# Compute the predictions for the next quarter

