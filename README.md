# Interfacing with Excel in R
This is a repository for storing materials and code for internal MoJ training on reading from and writing to excel using R. This training is suitable for those who have completed the [Intro R training](https://github.com/moj-analytical-services/IntroRTraining) or have at least reached an equivalent standard to having done this.   

Training is organised through a set of numbered steps to work through. This can be undertaken either in a classroom setting (being run in person/over Teams every few months) or in your own time either by following the instructions below or recordings of a previous session [here](https://web.microsoftstream.com/channel/aa3cda5d-99d6-4e9d-ac5e-6548dd55f52a). See [Remote learning](#remote-learning) for more tips on going through this material in your own time. If you work through the material by yourself please leave feedback about the material [here](https://airtable.com/shr9u2OJB2pW8Y0Af). 

Before starting, it is recommended that you clone this repository into R Studio on the analytical platform, to be able to access supporting data files and templates.

Please follow through training materials below:

# Pre-Materials: Package Options and Caveats

1. As R is open source software, anybody can write and release a package to perform a particular function. In some cases, there will be only option - usually in a niche area, or where the existing package works so well there are no variations or improvements that could be made.

For reading data into R, there are a number of packages available. At the most basic, you can use reading capability in the base R package `utils` - commands such as `read.csv`, `read.table` and `read.delim` will allow you to read in basic dataframes or csv files.

2. Beyond this, more advanced packages have then been created to interface R and Excel. These include (but are not limited to):
* readr
* openxlsx
* readxl
* gdata
* data.table

3. Each package will have specific features and functions, some of which will be more useful in specific situations than others. However, for the purposes of reading and writing basic data frames to and from excel, all will be sufficient. This course therefore focuses on three of the above packages, namely `readr` and `openxlsx` (as well as touching briefly on `utils` functionality)

4. Caveat - this training is designed to allow you to quickly and easily read/write .csv and .xlsx files; and control formatting/styles in excel using R. For detailed workings of specific functions, refer to the package documentation - as a rule this course will only help with the basic arguments that must be specified in functions, not optional extra arguments.


# Section A - Reading and Writing Data Frames from CSV

1. Let's start with the basics - reading in a csv file. If you've ever worked with data in R, the chances are you'll have read in a csv file of the data. We can either use the package `readr` (from the tidyverse) or the base package `utils` to do this.

We pass the name of the file as the argument to the function - this relies on the csv file being located in your working directory. If it isn't, you will have to five the full path.

```
library(readr)  #load the package

training_dataset       <- readr::read_csv("excel_training_dataset.csv")                     #more intelligent - recognises character columns

alt_training_dataset   <- utils::read.csv("~/r-excel-training/excel_training_dataset.csv")  #if we're not in the working directory, we have to specify the full file path

```

2. If we use the command `str(dataframe_name)` we can see how R has interpreted the csv file:

``` 
str(training_dataset)      #first two columns are character variables

str(alt_training_dataset)  #first two columns are factor variables

```

3. There's a couple of arguments in reading csvs, using the `readr::read_csv` function that you'll want to be aware of:

- `col_names` - if this is set to TRUE, the first row of the csv file will be interpreted by R as column names. This is the default setting, so be careful if you don't have column headings in your csv file.

- `col_types` - this can be used to specify what *type* you want each column to be interpreted as. It needs to be a vector, with length equal to the number of columns in the dataset. For example in the above, if we want our first column to be an integer, rather than character, we can override this using the `col_types` argument.

```
training_dataset       <- readr::read_csv("~/r-excel-training/excel_training_dataset.csv", 
                                          col_names = TRUE, 
                                          col_type = list(Year = col_integer(), Prisoner_Type = col_character(), `30_June_Population` = col_integer()))

```

Note that the base `read.csv` is more limited in functionality.  

4. To write back to csv, use `write_csv` (readr - twice as fast as utils) or `write.csv` (utils), specifying the name of the dataframe to be written, and the name of the csv file:

```
readr::write_csv(x = training_dataset, path = "example_output.csv")

utils::write.csv(x = training_dataset, "~/r-excel-training/another_output.csv")

```

There are also options to add or remove row and column names (if not included in your dataset), specify column types, or append to an existing csv. See `?write.csv` for details. 

*Exercises*  

A.1 Read in the csv file 'excel_exercises_dataset.csv'. Try doing this using different packages. What data types have the columns been imported as?

A.2 Read in the dataset specifying column types as follows: integer, factor, integer. *Hint* - you'll need to use an argument 'levels' to specify the categories your factor can take.

# Section B - Reading a .xlsx file and writing to a workbook

1. Now that we can read csv files, lets try reading in data from an xlsx file. An xlsx file will usually be a workbook with multiple sheets of data - even though these can be formatted or more details, we still need these to have the basic structure of a data frame in order to read something sensible.

As outlined in the first section, there are multiple ways to read in xlsx data, however for this training we'll demo using the `openxlsx` package.

R can read in one data frame at a time, therefore when reading xlsx data we'll need to read in one sheet at a time if we want to read straight to a data frame.

```
install.packages("openxlsx") #install the package (ignore if you already have this)

library(openxlsx)     #load the packages

population_sheet    <- openxlsx::read.xlsx("excel_training_dataset.xlsx", sheet = "Population")   #at the most basic, specify the file path and the name of the sheet

```

2. Note that here we have read the worksheet in directly, as a data frame.

An alternative approach would be to read in the *entire workbook* from file, as a workbook object in R. We can then read sheets from this workbook object as we wish. The concept of a workbook object will become more important when we look into writing data back to excel later.

Again we can try this using `openxlsx`:

```
wb_connect       <- openxlsx::loadWorkbook("excel_training_dataset.xlsx")
#read in the whole workbook

```

3. With a workbook object loaded, you can then read in sheets of data from this as needed. The benefit here is that we're only loading from a data store once - subsequent reading of worksheets is then from objects within R.

```
inflow_sheet_2     <- openxlsx::readWorkbook(wb_connect, "Inflow")  
#read in the specific sheet

#specify the name of the workbook object, and the name of the sheet
```

4. We have been specifying the name of the sheets often here. A quick command to get the names of all sheets in the workbook can be used as follows:

```
sheet_names        <- openxlsx::getSheetNames("excel_training_dataset.xlsx")
```

The sheet names are now stored as a character vector.  

*Exercises*  

B.1 Read in the excel workbook from the file 'excel_exercises_dataset.xlsx'. How can you tell this has imported correctly?

B.2 Read in a dataset from the 'Inflow' tab of the exercises xlsx. Do this directly from the xlsx file.

B.3 Now try reading in the dataset from the 'Population' tab of the xlsx. Do this from the loaded workbook in part B.1.

# Section C - Writing to .xlsx and styling the output

1. Now that we know how to write to csv, we can try writing data to an xls file, i.e. one with multiple sheets. The `openxlsx` package can be used to easily manipulate xls files. First, we need to create a 'workbook' object in R Studio - you can think of this as being the same as an excel workbook - i.e. it can store data frames in separate tabs or sheets. To create the workbook:

```
library(openxlsx)

my_workbook   <- openxlsx::createWorkbook()    #creates a blank workbook object

```

2. We next need some data to write into the workbook. From our original csv read, we can create a new table of total population at 30 June each year:

```
library(dplyr)

population_table     <- population_sheet %>% dplyr::group_by(Year) %>% dplyr::summarise(Population = sum(`30_June_Population`))

```
3. And now we can create sheets in our workbook, and write our tables to these worksheets

```
openxlsx::addWorksheet(my_workbook, sheet = "Population by Custody")  
openxlsx::addWorksheet(my_workbook, sheet = "Total Population") 
  #creates new sheets in our workbook

openxlsx::writeData(my_workbook, sheet = "Population by Custody", population_sheet)      
openxlsx::writeData(my_workbook, sheet = "Total Population", population_table)            
  #writes the dataframes to the separate sheets

#Note that this doesn't actually create an excel file - we are still working with our 'workbook' object in R Studio
```

4. Now that we have a ready workbook object, we can write to an excel file. This command will write the file in your current working directory, if we specify a name for the file as a string:

```
openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx")

```

5. Note that we can also add styles to the excel sheets using `openxlsx`. More on this later, but here we can look at some examples of some common styling options you will likely need when creating excel outputs. There are two steps to apply styles to excel tables:

* Create the style using `openxlsx::createStyle`. We can assign this to an object in R, so that once a style is specified we can use it multiple times.

* Add the style to the chosen table. We do this using the command `openxlsx::addStyle`; this is applied *after* data has been written, but *before* we save the workbook object to excel. 

Within the `addStyle` command, we will need to specify the workbook object, which sheet we are on, the style to apply, and the start row/column (optional). We also need to add arguments `gridExpand = TRUE` and `stack = TRUE` to make sure that any overlapping styles are combined, not overwritten (e.g. if we choose to **bold** and *italic* a certain row, these arguments make sure the second command is combined with the first, instead of overwriting it).

Here are some examples:

```
#First create the styles

boldstyle                 <- openxlsx::createStyle(textDecoration = "bold")

fontstyle                 <- openxlsx::createStyle(fontSize = 11, 
                                                   fontColour = "black", 
                                                   fontName = "Arial")
```
```
#Second add the styles to the tables. We'll add one to each of our tables

openxlsx::addStyle(wb = my_workbook, 
                   sheet = "Population by Custody", 
                   style = boldstyle, 
                   rows = 1, cols = 1:3, 
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = my_workbook, 
                   sheet = "Total Population", 
                   style = fontstyle, 
                   rows = 1:nrow(population_table), 
                   cols = 1:(ncol(population_table)), 
                   gridExpand = T, stack = T)

```
To check how this looks, write the workbook to excel again. Note we'll need to remove the xls file first, in order to save back out again.

```
if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")

openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx") 

```  

Or alternatively, we can use the argument `overwrite = TRUE` in the `saveWorkbook` command:

```
openxlsx::saveWorkbook(my_workbook, "Example_Output.xlsx", overwrite = TRUE)

```

*Exercises*  

C.1 Create a new workbook with new sheets 'Pop Copy' and 'Inflow Copy'; within the R Studio Environment. Write some population and inflow data to these new worksheets respectively (using data loaded in exercises A or B).

C.2 Create a new style with font size 16, colour red and font Calibri. Also create a bold style.

C.3 Apply the font style to just the header (top) row in the 'Pop Copy' worksheet. Apply the bold style to just the last row of the 'Inflow Copy' worksheet.

C.4. Write your workbook to excel and open it to check your data has written correctly and styles applied as per part C.3.

# Section D - Reading a table template and writing in position

1. We've seen how to write to an excel file by creating our own new workbook object, but what about reading in an existing xls file? This is useful if our source data is in an xls format (although generally separate csv files is usually recommended), or if we want to create a template for our workbook (e.g. with some base styling or permanent data/text).

`openxlsx` can be used to read in the xls as a workbook object via the `loadWorkbook` function; lets use it to read in the table template that is provided.

```

template       <- openxlsx::loadWorkbook("Table_Template.xlsx")

```

2. The workbook has two sheets, with the same names as our previously created workbook. However, by loading the template, several styles and headings are incorporated. This can be especially useful if there are elements of your xls file that will remain permanent. We can write our existing two tables to this workbook in the same way as before, to check the output.

Note that we also specify the start column and row of where we want the data to be written, to fit the template. For this, use the `xy` argument in `openxlsx`.

```
#Write data to the workbook object

openxlsx::writeData(template, sheet = "Table 1", population_sheet, xy = c(1,4)) 
openxlsx::writeData(template, sheet = "Table 2", population_table, xy = c(1,4))

#Save the output file
if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")
openxlsx::saveWorkbook(template, "Example_Output.xlsx") 
```  

3. We can also use a template styles document to specify commonly used styles. This requires some knowledge of how lists work in R, but has the benefit of no longer being required to manually create multiple different styles.

First, we create a workbook in excel with cells in the first column formatted as desired. This is illustrated by opening the .xlsx file 'styles_example.xlsx' within the repository. Once we have this style template, we use the 'getStyles` function provided by `openxlsx` to read the styles in as objects in R. We use the list notation [[x]] to determine which style we want to access, with x indicating the row number in our workbook.

```
style_wb        <- openxlsx::loadWorkbook("style_template.xlsx")     #read in the workbook

style_1         <- openxlsx::getStyles(style_wb)[[1]]
style_2         <- openxlsx::getStyles(style_wb)[[2]]

#access styles - we specify the name of the workbook as an argument to the function, then use list subsetting to access the individual style

style_1                 #calling the style shows us what type of object it is

```

We will look at how these styles can be applied to workbooks in Part E.

# Section E - Further table styling

The package `openxlsx`, as touched on in Section C, allows us to apply further styling, as covered in this final section of the course. Before we start, let's sort our data by year:

```
population_sheet  <- population_sheet %>% dplyr::arrange(Year)

```

1. When presenting tables in excel, use of white space is key. We can adjust row heights in order to create white space, for example before a total line or after every nth entry. Similarly we can adjust column widths.

Note that here we don't create a style and then write it - we do both steps in one command. 

```
openxlsx::setRowHeights(template, sheet = "Table 1", 
                        rows = nrow(population_sheet) + 3, height = 25)

openxlsx::setRowHeights(template, sheet = "Table 1", 
                        rows = 4, height = 30)

#increase the row height of the final row and the title row
#remember to take into account that our data frame starts in row 4; whereas setting row heights is referencing the excel sheet (not just our data frame)

openxlsx::setColWidths(template, sheet = "Table 1", 
                       cols = 2:3, width = 25)

```

2. We could create borders in our table template (as loaded in Section D), however R lets us create borders remotely, if we wish. This is useful if your table has a dynamic number of rows (e.g. if we're adding a new year to the dataset with each new iteration of the spreadsheet)

Here we create border styles, which we can then apply to Table 1.

```
rightborder  <- openxlsx::createStyle(border = "right", borderColour = "black", borderStyle = "thin")
heavybottom   <- openxlsx::createStyle(border = "bottom", borderColour = "black", borderStyle = "medium")

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = rightborder, 
                   rows = 4:(nrow(population_sheet)+3), cols = 1, 
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = heavybottom, 
                   rows = nrow(population_sheet)+3, cols = 1:(ncol(population_sheet)), 
                   gridExpand = T, stack = T)

#we can specify where we want the border, what colour it is, how thick it is, etc.

#remember to gridExpand and stack
```

3. It is also possible to align text (both horizontally and vertically) - for example the year should usually be left aligned, but the rest of our data we want to be right aligned. Below we create the style and apply

```
rightalign    <- openxlsx::createStyle(halign = "right")
leftalign     <- openxlsx::createStyle(halign = "left")

#Note - use argument 'valign' if you want to vertically align in a cell

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = leftalign, 
                   rows = 4:(nrow(population_sheet)+3), cols = 1, 
                   gridExpand = T, stack = T)

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = rightalign, 
                   rows = 4:(nrow(population_sheet)+3), cols = 2:ncol(population_sheet), 
                   gridExpand = T, stack = T)
```

4. We can format cells remotely, e.g. to show percentages, text, comma formated numbers, dates, etc (for a full list see [here](https://rdrr.io/cran/openxlsx/man/createStyle.html)). Let's make column three a comma formatted number:

```
commaformat   <- openxlsx::createStyle(numFmt = "COMMA")

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = commaformat, 
                   rows = 5:(nrow(population_sheet)+3), cols = 3, 
                   gridExpand = T, stack = T)
```

5. Finally, to stop cell overlap we may want to wrap text in e.g. a title, or more generally when writing text data to excel. Create a style and apply it as usual.

```
wrapstyle     <- openxlsx::createStyle(wrapText = TRUE)

openxlsx::addStyle(wb = template, sheet = "Table 1", 
                   style = wrapstyle, 
                   rows = 4, cols = 1:ncol(population_sheet), 
                   gridExpand = T, stack = T)

```
6. We can also write the styles that we imported earlier from template - i.e. without having to create the new style in R:

```
openxlsx::addStyle(wb = template, sheet = "Table 1",
                   style = style_1,
                   rows = 1, cols = 1:ncol(population_sheet),
                   gridExpand = T, stack = T)
                   
openxlsx::addStyle(wb = template, sheet = "Table 2",
                   style = style_2,
                   rows = 7, cols = 2:3,
                   gridExpand = T, stack = T)
                   
#add the style to random cells to illustrate the styles being applied
```

7. Now we can write the revised workbook to excel, to see the results of our styling.

```
#Save the output file
if (file.exists("Example_Output.xlsx")) file.remove("Example_Output.xlsx")
openxlsx::saveWorkbook(template, "Example_Output.xlsx") 
```  

*Exercises*  

E.1 Create styles for the following:
- left dashed border on the cell
- doube line border on the bottom of the cell
- central horizontal alignment
- number formatted cell with 2 d.p.
- yellow fill (*hint*: use the argument 'fgFill')
You can try doing this using a .xlsx template, or by creating the styles manually in R.

E.2 Apply the styles to a selection of data in the worksheets created in the Exercises to Section C. Note that it will be necessary to re-create your blank workbook and load in the data (i.e. re-run your code) so that the existing styles added in Section C will not be imported. For a quick set-up, run the following code:

```
exercises_wb      <- openxlsx::createWorkbook() 
openxlsx::addWorksheet(exercises_wb, sheet = "Pop Copy")  
openxlsx::addWorksheet(exercises_wb, sheet = "Inflow Copy")
openxlsx::writeData(exercises_wb, sheet = "Pop Copy", pop_exercises)   #writes back out the data frame    
openxlsx::writeData(exercises_wb, sheet = "Inflow Copy", inflow_exercises)
```

# Section F - Reading/Writing from/to Amazon S3

As well as loading/writing datasets stored in your local area of the Analytical Platform, you are likely to want to interface with the Amazon S3 storage areas, for example when storing large datasets or sharing data between individuals in a team.

The MoJ has developed a package s3tools for interfacing with S3, with guidance available in the README of the repository here: [https://github.com/moj-analytical-services/s3tools](https://github.com/moj-analytical-services/s3tools).

In short, the following code can be used to read in a csv from a bucket (data folder) in S3:

```
s3tools::read_using(FUN=readr::read_csv, path = "alpha-everyone/s3tools_tests/iris_base.csv")

#we can specify the read function ('FUN' argument), so that we could change this to read in data from an excel file

#'alpha-everyone' is the name of the bucket in S3, with folders/sub-folders and file names specified in the usual way below this

```

To write to S3, we can either write a specific file from our instance of the platform, or we can write a dataset:

```
s3tools::write_file_to_s3("my_downloaded_file.csv", "alpha-everyone/delete/my_downloaded_file.csv", overwrite =TRUE)

#the first argument is the name of the file, so we can write any file type

#the third argument makes sure that, if the file exists, it will be overwritten with the latest version you are trying to write

s3tools::write_df_to_csv_in_s3(iris, "alpha-everyone/delete/iris.csv", overwrite =TRUE)

#in this second example we write a dataset ('iris') rather than a pre-written file. Otherwise arguments are the same

```

Note there is also guidance on this process in the analytical platform user guidance document: [https://moj-analytical-services.github.io/platform_user_guidance/](https://moj-analytical-services.github.io/platform_user_guidance/)


# Section G - xltabr

The MoJ has created a bespoke package for writing cross-tabulations to excel, for publication purposes.

For more information, please visit the github page for xltabr [https://github.com/moj-analytical-services/xltabr](https://github.com/moj-analytical-services/xltabr) and follow the instructions contained in the README.

# Section H - Example in Practice

Refer to the 'Table Scripts' folder of the following repository: [https://github.com/moj-analytical-services/Coroner-Markdown-Tables](https://github.com/moj-analytical-services/Coroner-Markdown-Tables)

Note that this is just one (quite untidy!) example of reading and writing excel data/tables; and is accessible cross-MoJ at the time of writing of this training.

# Section I - Queries and Further Resources

* Further resources? - Try:
    + openxlsx - [https://cran.r-project.org/web/packages/openxlsx/openxlsx.pdf](https://cran.r-project.org/web/packages/openxlsx/openxlsx.pdf)
    + readr - [https://readr.tidyverse.org/](https://readr.tidyverse.org/)
    + XLConnect - [http://www.milanor.net/blog/steps-connect-r-excel-xlconnect/](http://www.milanor.net/blog/steps-connect-r-excel-xlconnect/)
    + General questions - [https://stackoverflow.com/](https://stackoverflow.com/)
  

* Dealing with untidy input data? (not that MoJ data is ever untidy...)<sup>1</sup> - try using `tidyxl` and `unpivotr`- Duncan Garmondsway (GDS) has written a helpful how to guide here: [https://nacnudus.github.io/unpivotr/articles/worked-examples.html](https://nacnudus.github.io/unpivotr/articles/worked-examples.html)

* Queries on this material? Contact Matt: [matthew.funnell@justice.gov.uk](matthew.funnell@justice.gov.uk) or another member of the R Training Group.

* Suggestions for improvements? Raise an issue on this repository, or email Matt as above.

1 - Sarcasm credit: Mike Hallard
