# Clear all data
rm(list=ls())

# Load packages
library(tidyverse)
library(lubridate)
library(readxl)
library(xlsx)

# Set file path to desktop

setwd("C:/Users/mkorzec/Downloads")

# Import Sales and Inventory Data
sales <- read_xlsx("2019_Hyundai_Sales_Report.xlsx", sheet = 'Raw Data')
inv <- read_xlsx("2019_Hyundai_Sales_Report.xlsx", sheet = 'Inventory')

###################################################
# 'FL123' 'FL121' 'IL071' ' IL082' 'IL088' "IN048' 'MO046'
###################################################

sales_1 <- sales %>%
  filter(Dealership == 'MO046') %>%
  group_by(Model) %>%
  summarize(Sales = n()) %>%
  mutate(Percent = (Sales / sum(Sales)) * 100) %>%
  mutate(Percent = round(Percent, 2)) %>%
  mutate(Model = toupper(Model))
  
  
inv_1 <- inv %>%
  filter(Dealership == 'MO046') %>%
  group_by(Model) %>%
  mutate(Stock = as.numeric(`Dealer Stock`)) %>%
  summarize(Stock = sum(Stock), 
            Total = sum(`Total`)) %>%
  mutate(Percent_Stock = (Stock / sum(Stock)) * 100,
         Percent_Total = (Total / sum(Total)) * 100) %>%
  mutate(Percent_Stock = round(Percent_Stock, 2),
         Percent_Total = round(Percent_Total, 2))

full <- full_join(sales_1, inv_1, by = 'Model')

MO046 <- full %>%
  mutate(Difference = (Percent_Total - Percent)) %>%
  rename('Percent of Sales' = Percent,
         'In Stock' = Stock,
         'Percent of Stock' = Percent_Stock,
         'Percent of Total' = Percent_Total) %>%
  add_column(Dealership = 'MO046')

# Join data 
data <- rbind(FL121, FL123, IL071, IL082, IL088, IN048, MO046)

# Write to Excel
data <- as.data.frame(data)

write.xlsx(data, file = "2019_Hyundai_Sales_Report.xlsx", 
           sheetName = "Recommended_Inventory", 
           row.names = FALSE, 
           append = TRUE)
