#Code for the presenter to run that mirrors the README instructions code

#Initial packages
library(dplyr)
library(readr)
library(openxlsx)

#A

training_dataset       <- readr::read_csv("excel_training_dataset.csv")                     #more intelligent - recognises character columns

alt_training_dataset   <- utils::read.csv("~/r-excel-training/excel_training_dataset.csv")  #if we're not in the working directory, we have to specify the full file path

str(training_dataset)      #first two columns are character variables

str(alt_training_dataset)  #first two columns are factor variables

training_dataset       <- readr::read_csv("~/r-excel-training/excel_training_dataset.csv", col_names = TRUE, col_type = list(Year = col_integer(), Prisoner_Type = col_character(), `30_June_Population` = col_integer()))

readr::write_csv(x = training_dataset, path = "example_output.csv")

utils::write.csv(x = training_dataset, "~/r-excel-training/another_output.csv")

#A - Exercise Answers

exercises_dataset      <- readr::read_csv("excel_exercises_dataset.csv")

str(exercises_dataset)

exercises_dataset_2    <- utils::read.csv("excel_exercises_dataset.csv")

exercises_dataset      <- readr::read_csv("excel_exercises_dataset.csv", col_type = list(Year = col_integer(), Prisoner_Type = col_factor(levels = c("Determinate", "Indeterminate", "Non-Criminal", "Recall", "Remand")), `30_Sept_Population` = col_integer()))

str(exercises_dataset)

#B

population_sheet    <- openxlsx::read.xlsx("excel_training_dataset.xlsx", sheet = "Population")   #at the most basic, specify the file path and the name of the sheet

wb_connect       <- openxlsx::loadWorkbook("excel_training_dataset.xlsx")

inflow_sheet_2      <- openxlsx::readWorkbook(wb_connect, "Inflow")  

#specify the name of the workbook object, and the name of the sheet

sheet_names        <- openxlsx::getSheetNames("excel_training_dataset.xlsx")
sheet_names

#B - Exercise Answers

exercises_wb            <- openxlsx::loadWorkbook("excel_exercises_dataset.xlsx")

inflow_exercises        <- openxlsx::read.xlsx("excel_exercises_dataset.xlsx", sheet = "Inflow")   

pop_exercises           <- openxlsx::readWorkbook(exercises_wb, "Population")  

#C

library(openxlsx)

my_workbook   <- openxlsx::createWorkbook()    #creates a blank workbook object

library(dplyr)

population_table     <- population_sheet %>% dplyr::group_by(Year) %>% dplyr::summarise(Population = sum(`30_June_Population`))

openxlsx::addWorksheet(my_workbook, sheet = "Population by Custody")  
openxlsx::addWorksheet(my_workbook, sheet = "Total Population") #creates new sheets in our workbook

openxlsx::writeData(my_workbook, sheet = "Population by Custody", population_sheet)   #writes back out the data frame    
openxlsx::writeData(my_workbook, sheet = "Total Population", population_table)            #writes the dataframes to the separate sheets

#Note that this doesn't actually create an excel file - we are still working with our 'workbook' object in R Studio

openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx")

#First create the styles
boldstyle                 <- openxlsx::createStyle(textDecoration = "bold")
fontstyle                 <- openxlsx::createStyle(fontSize = 11, fontColour = "black", fontName = "Arial")

openxlsx::addStyle(wb = my_workbook, sheet = "Population by Custody", style = boldstyle, rows = 1, cols = 1:3, gridExpand = T, stack = T)

openxlsx::addStyle(wb = my_workbook, sheet = "Total Population", style = fontstyle, rows = 1:nrow(population_table), cols = 1:(ncol(population_table)), gridExpand = T, stack = T)

if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")
openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx") 

openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx", overwrite = TRUE)

#C - Exercise Answers

exercises_wb      <- openxlsx::createWorkbook() 
openxlsx::addWorksheet(exercises_wb, sheet = "Pop Copy")  
openxlsx::addWorksheet(exercises_wb, sheet = "Inflow Copy")
openxlsx::writeData(exercises_wb, sheet = "Pop Copy", pop_exercises)   #writes back out the data frame    
openxlsx::writeData(exercises_wb, sheet = "Inflow Copy", inflow_exercises)

boldstyle                 <- openxlsx::createStyle(textDecoration = "bold")
fontstyle                 <- openxlsx::createStyle(fontSize = 16, 
                                                   fontColour = "red", 
                                                   fontName = "Calibri")

openxlsx::addStyle(wb = exercises_wb, 
                   sheet = "Pop Copy", 
                   style = fontstyle, 
                   rows = 1, cols = 1:3, 
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = exercises_wb, 
                   sheet = "Inflow Copy", 
                   style = boldstyle, 
                   rows = nrow(inflow_exercises) + 1, 
                   cols = 1:(ncol(inflow_exercises)), 
                   gridExpand = T, stack = T)

openxlsx::saveWorkbook(exercises_wb, "C_Exercise_Output.xlsx", overwrite = TRUE) 

#D

template       <- openxlsx::loadWorkbook("Table_Template.xlsx")

#Write data to the workbook object
openxlsx::writeData(template, sheet = "Table 1", population_sheet, xy = c(1,4)) 
openxlsx::writeData(template, sheet = "Table 2", population_table, xy = c(1,4))

#Save the output file
if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")
openxlsx::saveWorkbook(template, "Example_Output.xlsx") 

style_wb        <- openxlsx::loadWorkbook("style_template.xlsx")     #read in the workbook

style_1         <- openxlsx::getStyles(style_wb)[[1]]
style_2         <- openxlsx::getStyles(style_wb)[[2]]

#access styles - we specify the name of the workbook as an argument to the function, then use list subsetting to access the individual style

style_1                 #calling the style shows us what type of object it is

#E 

population_sheet  <- population_sheet %>% dplyr::arrange(Year)

openxlsx::setRowHeights(template, sheet = "Table 1", rows = nrow(population_sheet) + 3, height = 25)
openxlsx::setRowHeights(template, sheet = "Table 1", rows = 4, height = 30)

#increase the row height of the final row and the title row
#remember to take into account that our data frame starts in row 4; whereas setting row heights is referencing the excel sheet (not just our data frame)

openxlsx::setColWidths(template, sheet = "Table 1", cols = 2:3, width = 25)

rightborder  <- openxlsx::createStyle(border = "right", borderColour = "black", borderStyle = "thin")
heavybottom   <- openxlsx::createStyle(border = "bottom", borderColour = "black", borderStyle = "medium")

openxlsx::addStyle(wb = template, sheet = "Table 1", style = rightborder, rows = 4:(nrow(population_sheet)+3), cols = 1, gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 1", style = heavybottom, rows = nrow(population_sheet)+3, cols = 1:(ncol(population_sheet)), gridExpand = T, stack = T)

#we can specify where we want the border, what colour it is, how thick it is, etc.

#remember to gridExpand and stack

rightalign    <- openxlsx::createStyle(halign = "right")
leftalign     <- openxlsx::createStyle(halign = "left")

#Note - use argument 'valign' if you want to vertically align in a cell

openxlsx::addStyle(wb = template, sheet = "Table 1", style = leftalign, rows = 4:(nrow(population_sheet)+3), cols = 1, gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 1", style = rightalign, rows = 4:(nrow(population_sheet)+3), cols = 2:ncol(population_sheet), gridExpand = T, stack = T)

commaformat   <- openxlsx::createStyle(numFmt = "COMMA")

openxlsx::addStyle(wb = template, sheet = "Table 1", style = commaformat, rows = 5:(nrow(population_sheet)+3), cols = 3, gridExpand = T, stack = T)

wrapstyle     <- openxlsx::createStyle(wrapText = TRUE)

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = wrapstyle, 
                   rows = 4, cols = 1:ncol(population_sheet), 
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 1",
                   style = style_1,
                   rows = 1, cols = 1:ncol(population_sheet),
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 2",
                   style = style_2,
                   rows = 7, cols = 2:3,
                   gridExpand = T, stack = T)

#Save the output file
if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")
openxlsx::saveWorkbook(template, "Example_Output.xlsx") 

#E - Exercise Answers

left_dash     <- openxlsx::createStyle(border = "left", borderStyle = "dashed")
double_bottom <- openxlsx::createStyle(border = "bottom", borderStyle = "double")
central_h     <- openxlsx::createStyle(halign = "center")
date_format   <- openxlsx::createStyle(numFmt = "0.00")
yellow_fill   <- openxlsx::createStyle(fgFill = "yellow")

exercises_wb      <- openxlsx::createWorkbook() 
openxlsx::addWorksheet(exercises_wb, sheet = "Pop Copy")  
openxlsx::addWorksheet(exercises_wb, sheet = "Inflow Copy")
openxlsx::writeData(exercises_wb, sheet = "Pop Copy", pop_exercises)   #writes back out the data frame    
openxlsx::writeData(exercises_wb, sheet = "Inflow Copy", inflow_exercises)

openxlsx::addStyle(wb = exercises_wb, sheet = "Pop Copy", style = left_dash, rows = 1:(nrow(pop_exercises)+1), cols = 2:4, gridExpand = T, stack = T)
openxlsx::addStyle(wb = exercises_wb, sheet = "Pop Copy", style = double_bottom, rows = nrow(pop_exercises)+1, cols = 1:ncol(pop_exercises), gridExpand = T, stack = T)
openxlsx::addStyle(wb = exercises_wb, sheet = "Pop Copy", style = central_h, rows = 1:(nrow(pop_exercises)+1), cols = 1, gridExpand = T, stack = T)
openxlsx::addStyle(wb = exercises_wb, sheet = "Inflow Copy", style = date_format, rows = 2:nrow(inflow_exercises), cols = 3, gridExpand = T, stack = T)
openxlsx::addStyle(wb = exercises_wb, sheet = "Inflow Copy", style = yellow_fill, rows = 1, cols = 1:ncol(inflow_exercises), gridExpand = T, stack = T)

if (file.exists("C_Exercise_Output.xlsx")) file.remove("C_Exercise_Output.xlsx")
openxlsx::saveWorkbook(exercises_wb, "C_Exercise_Output.xlsx") 