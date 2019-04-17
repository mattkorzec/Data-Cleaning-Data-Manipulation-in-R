# Read all data
hist <- read_xlsx("Sales_Per_Month.xlsx")

# Change the type of date
hist$month <- as_date(hist$month)

test <- gather(hist, Dealership, Sales, -month)
test <- arrange(test, month, Dealership)

# Append existing file to the main excel file as a new sheet
test <- as.data.frame(test)
write.xlsx(test, file = "2019_Hyundai_Sales_Report.xlsx", sheetName = "Historical Data", row.names = FALSE, 
           append = TRUE)