#' @title Create a \code{kable} abundance table
#' @description  #Function to make abundance kable table across two marginal variables (col1, col2)
#' @param d Data frame or tibble
#' @param col1 First column to use (e.g. species)
#' @param col2 Second column to use (e.g. site)
#' @param totals Should column and row totals be added?
#' @return A \code{kable} table
#' @export
#' @examples
#' NULL
abundTable <- function(d,col1,col2,totals=FALSE){
  require(tidyverse); require(knitr); require(kableExtra)
  if(missing(col1)|missing(col2)) stop('Specify columns to use')

  col2 <- enquo(col2) #Defuses arguments

  if(totals){
    #Marginal column/row totals
    d <- d %>% count({{col1}},{{col2}}) %>%
      pivot_wider(names_from={{col1}},values_from=n,values_fill=0) %>%
      arrange({{col2}}) %>%
      rowwise() %>% mutate(TOTAL=sum(c_across(-{{col2}}))) %>%
      #Hacky way around dealing with `summarize(across(-col2,sum))` (can't be used within bind_rows)
      column_to_rownames(var=as_label(col2)) %>%
      bind_rows(.,colSums(.)) %>% rownames_to_column(var=as_label(col2)) %>%
      mutate({{col2}} := c({{col2}}[-n()],'TOTAL')) %>%
      column_to_rownames(var=as_label(col2))
  } else {
    #Version without marginal totals
    d <- d %>% count({{col1}},{{col2}}) %>%
      pivot_wider(names_from={{col1}},values_from=n,values_fill=0) %>%
      arrange({{col2}}) %>%
      column_to_rownames(var = as_label(col2))
  }
  return(d)
}
