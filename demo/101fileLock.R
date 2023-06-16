library(XLConnect)
#test1
wb <- loadWorkbook("./Cars.xlsx", create = TRUE)
createSheet(wb, "Data")
# open file with Ms Excel
writeWorksheet(wb, cars, "Data")
# make some change via MSE, save
saveWorkbook(wb)
# make some change via MSE, save

#pass

#2
wb <- loadWorkbook("Cars.xlsx")
# open MSE
dt.Fees <- data.table::setDT(readWorksheet(wb, "Data"))
# make some change via MSE, save => share violation (lock)
#fail