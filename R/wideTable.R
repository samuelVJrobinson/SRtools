#' @title Create a wide kable table
#' @description Function to make a "wide kable table", with column names replicated across the page
#' @param d Data frame or tibble
#' @param countBy
#' @param countNum
#' @param nsplit How many splits should be made?
#' @param colNames Change column names
#' @param caption Table caption
#' @param returnDF Return results as a data frame?
#' @param ... Other arguments for \code{kable_styling}
#' @return A kable table
#' @export
#' @examples
#' NULL
wideTable <- function(d,countBy={{countBy}},countNum={{countNum}},nsplit=NA,colNames=c("countBy","count"),
                      caption='',returnDF=FALSE,...){
  require(tidyverse)
  require(knitr)
  library(kableExtra)
  if(is.na(nsplit)) stop('Specify number of splits to perform')
  if('sf' %in% class(d)) d <- d %>% sf::st_drop_geometry() #Drop geometry if sf object

  d <- d %>%
    # #Get counts in each category - moved outside of function
    # filter(!is.na({{countBy}})) %>% count({{countBy}}) %>%
    # arrange(n) %>%
    mutate({{countBy}}:=factor({{countBy}},levels={{countBy}})) %>%
    #Splits into nsplit sets of columns
    mutate(ord=cut(rev(rank(as.numeric({{countBy}}))),nsplit,labels=1:nsplit)) %>%
    arrange(ord,desc({{countNum}})) %>% mutate(across(c({{countBy}},{{countNum}}),as.character)) %>%
    group_by(ord) %>% mutate(id=1:n()) %>% ungroup() %>%
    pivot_longer(cols={{countBy}}:{{countNum}}) %>%
    unite(name,name,ord,sep="_") %>% pivot_wider(id_cols=id,names_from=name,values_from=value) %>%
    select(-id) %>% mutate(across(everything(),~ifelse(is.na(.)," ",.)))

  if(returnDF) return(d) #Returns dataframe

  d %>% kable(col.names=rep(colNames,nsplit),align=rep(c("r","l"),nsplit),
              caption=caption) %>%
    column_spec(column=seq(1,nsplit*2,2),italic=TRUE) %>% #Genus names
    kable_styling(latex_options = c("striped"),...)
}
