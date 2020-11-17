#'tablestyle
#'
#'Standard formatting styles for an exampl excel output at the MoJ.
#'
#'INPUTS
#'@param wb - the name of the workbook you are writing data to
#'@param sheet - the sheet of the workbook you are writing styles to
#'@param tibble - a data frame - the table that is being written in to excel
#'@param startrow - a number - the first row in excel where the first line of data (i.e. column headings) will be written
#'@param footerlength - a number - the number of footer notes that are required, including zero/.. explanation but not including title ("notes"). Defaults to zero.
#'Columns are set by default as all columns in the table (i.e. formatting is applied to the whole row)
#'
#'EXAMPLE
#'tablestyle(wb, "Table 4", Table_4_Set, 5, 4) applies all standard table stylings to Table 4 of the workbook 'wb', which is written from row 5 onwards in excel and the footer is length four.
#'
#'Styles applied are as follows:
#'- thick bottom border
#'- increased row heights of the title, top and bottom rows of the table
#'- top align the last row of the table to create white space below it
#'- underline the title headings row
#'- apply standard font (pt 11, arial, black) to the table
#'- apply standard font (pt 10, arial, black) to the footer
#'- apply bold heading for the footer 'Notes'
#'- wrap text in the footer
#'- left align the first column (e.g. for Coroner tables this is nearly always year)
#'- right align all other columns within the table (e.g. these are usually numbers)
#'
#'Should any styles in the table be different to those specified in the function, styles can be overwritten by adding a new style
#'after calling this function
#'

tablestyles     <- function(wb, sheet, tibble, startrow, footerlength = 0){

  #Make sure the startrow is a number > 0
  if(startrow < 1){
    stop("Start row must be a positive integer")
  } else if(is.numeric(startrow == FALSE)){
      stop("Start row must be a number")
    } 
  else {
  
  #create styles for usage
  boldstyle     <- openxlsx::createStyle(textDecoration = "bold")
  fontstyle     <- openxlsx::createStyle(fontSize = 11, fontColour = "black", fontName = "Arial")
  footerfont    <- openxlsx::createStyle(fontSize = 10, fontColour = "black", fontName = "Arial")
  bottomborder  <- openxlsx::createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
  heavybottom   <- openxlsx::createStyle(border = "bottom", borderColour = "black", borderStyle = "medium")
  rightalign    <- openxlsx::createStyle(halign = "right")
  leftalign     <- openxlsx::createStyle(halign = "left")
  topalign      <- openxlsx::createStyle(valign = "top")
  botalign      <- openxlsx::createStyle(valign = "bottom")
  wrapstyle     <- openxlsx::createStyle(wrapText = TRUE)
  
  #get table dimensions for writing styles to
  endrow        <- startrow + nrow(tibble)
  tibblerows    <- startrow:endrow
  tibblecols    <- 1:ncol(tibble)
  
  #set row heights for key rows
  openxlsx::setRowHeights(wb, sheet, rows = c(startrow, startrow + 1, endrow), heights = c(60, 25, 25))
  
  #for loop and conditional to change row height wherever the year ends in 5 or 0 (i.e. every 5 years), as is standard in the example publication
  for (i in tibble$Year){
    lastdigit   <- i %% 5
    if (lastdigit == 0){
      rowi        <- startrow + which(tibble$Year == i)
      openxlsx::setRowHeights(wb, sheet, rows = rowi, height = 25)
      openxlsx::addStyle(wb, sheet, botalign, rows = rowi, cols = tibblecols, gridExpand = T, stack = T)  #bottom align these rows to make sure the gap is consistent
    }
  }
  
  #bottom align the first row
  openxlsx::addStyle(wb, sheet, botalign, rows = startrow + 1, cols = tibblecols, gridExpand = T, stack = T)
  
  #create bottom border and top align the last row (to create whitespace)
  openxlsx::addStyle(wb, sheet, heavybottom, rows = endrow, cols = tibblecols, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, topalign, rows = endrow, cols = tibblecols, gridExpand = T, stack = T)
  
  #title headings row - bold, border, wrap text, right align
  openxlsx::addStyle(wb, sheet, boldstyle, rows = startrow, cols = tibblecols, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, bottomborder, rows = startrow, cols = tibblecols, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, wrapstyle, rows = startrow, cols = tibblecols, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, rightalign, rows = startrow, cols = tibblecols, gridExpand = T, stack = T)
  
  #footer - bold the heading, wrap the text, set the font size/style
  openxlsx::addStyle(wb, sheet, boldstyle, rows = endrow + 2, cols = 1, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, wrapstyle, rows = (endrow + 3):(endrow + 8 + footerlength), cols = 1, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, footerfont, rows = (endrow + 2):(endrow + 8 + footerlength), cols = 1, gridExpand = T, stack = T)
  
  #merge footer cells so that long text gets wrapped correctly
  for(i in 3:(3 + footerlength)){
    openxlsx::mergeCells(wb, sheet, cols = tibblecols, rows = endrow + i)
  }
  
  #font and right align of the whole table
  openxlsx::addStyle(wb, sheet, fontstyle, rows = tibblerows, cols = tibblecols, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet, rightalign, rows = (startrow + 1):endrow, cols = 2:ncol(tibble), gridExpand = T, stack = T)
  
  #left align col 1 - usually the year therefore left align is appropriate
  openxlsx::addStyle(wb, sheet, leftalign, rows = tibblerows, cols = 1, gridExpand = T, stack = T)
  }
}

