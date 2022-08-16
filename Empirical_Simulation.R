################################################################################
###               Empirical Simulation in the Master Thesis                  ###
### The Quest for a Modern and Dynamic Budgeting and Forecasting Approach    ###
###               By Max Kneißler (University of Tübingen)                   ###
################################################################################



# Packages for the analysis
library(tidyverse)
library(readxl)


# Import the different sheets of the 3M reports database
Income_Statement <- read_excel("3M_Data.xlsx", sheet = "Income_Statement")
Sales_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "Sales_per_Division")
OpIncome_per_Division <- 
                    read_excel("3M_Data.xlsx", sheet = "OpIncome_per_Division")
Sales_per_Region <- read_excel("3M_Data.xlsx", sheet = "Sales_per_Region")
Division_Region_Matrix_Sales <- 
            read_excel("3M_Data.xlsx", sheet = "Division_Region_Matrix_Sales")
Ext_Parameters <- read_excel("3M_Data.xlsx", sheet = "External Parameters")


# Import external parameters / factors


# Training for 4/5 years


# Validation year 2021


# Compute the predictions for the next quarter

