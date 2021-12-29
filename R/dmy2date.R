#' @title Create Date column
#' @description Converts day,month,year to a Date column (\code{as.Date})
#' @param d Data frame
#' @param day Day column
#' @param month Month column
#' @param year Year column
#' @param newName Name for new date column (default = "date")
#' @param dateFormat Date format for each column (default = c("\%d","\%m","\%Y"), see \code{strptime} for more codes)
#' @return A data frame
#' @export
#'
#' @examples
#' dat <- data.frame(a=letters[1:2],d=c(10,12),
#'     m=c(6,12),y=c(1988,2012))
#' dmy2date(dat,d,m,y,'newDate')
dmy2date <- function(d,day,month,year,newName='date',dateFormat=c('%d','%m','%Y')){
  require(dplyr)
  day <- enquo(day); month <- enquo(month); year <- enquo(year)
  newName <- sym(newName)

  d %>%
    unite(!!newName,c({{day}},{{month}},{{year}}),sep='_') %>%
    mutate(!!newName := as.Date(!!newName,format=paste(dateFormat,collapse='_')))
}


