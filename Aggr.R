##########################################################################
# Aggragating Sales from historical data into Monthly sales
##########################################################################

# Clear all data
rm(list=ls())

# Read Excel Workbook file
data <- read_xls("xx###-inventory.xls")

x <- data %>%
  rename(Date = `Retail Date`, Model = `Model Name`) %>%
  filter(Status == 'RDR') %>%
  select(VIN, Model, Date) %>%
  mutate(Date = mdy(Date)) %>%
  group_by(Date) %>%
  summarize(Sales = n()) %>%
  group_by(month = floor_date(Date, "month")) %>%
  summarize(Sales = sum(Sales))

# Export to an existing excel workbook as a new sheet
ts <- as.data.frame(x)

write.xlsx(ts, file = "2019_Hyundai_Sales_Report.xlsx", sheetName = "Time Series", row.names = FALSE, append = TRUE)