#' @title Make taxonomic abundance plot
#' @description Function to make bee taxonomic abundance plots (species, genus, family level counts) from data frame.
#' @param d Data frame or tibble
#' @param family Family column
#' @param genus Genus column
#' @param species Species column - can also use genSpp column (see below)
#' @param genSpp GenSpp column (optional - create using \code{\link{makeGenSpp}})
#' @param colourSet Colour set from ColourBrewer (default = "Set1")
#' @param scaleYtext Scaling factor for y-axis text (default = c(1,1,1))
#' @param keepSpp Keep morphospecies? (default = TRUE)
#' @param returnPlots Should plots be returned in a list instead? (default = FALSE)
#' @export
#' @return A \code{ggarrange} (\code{ggplot}) object, showing bee abundances split up by species, genus, and family
#'
#' @examples
#' dat <- data.frame(f=c('Apidae','Apidae','Apidae','Colletidae','Andrenidae','Andrenidae'),
#'     g=c('Bombus','Bombus','Apis','Hylaeus','Lasioglossum','Lasioglossum'),
#'     s=c('rufocinctus','rufocinctus','mellifera','annulatus','zonulum','spp.'))
#' abundPlots(dat,f,g,s,keepSpp=FALSE)
#' abundPlots(dat,f,g,s,keepSpp=TRUE)
abundPlots <- function(d,family=family,genus=genus,species=NULL,genSpp=NULL,colourSet='Set1',
                       scaleYtext=c(1,1,1),keepSpp=TRUE,returnPlots=FALSE){
  require(RColorBrewer)
  require(dplyr); require(tidyr)
  require(ggpubr); require(rlang)
  options(dplyr.summarise.inform=FALSE)

  if('sf' %in% class(d)) d <- d %>% sf::st_drop_geometry() #Drop geometry if sf object

  family <- enquo(family) #Defuse expressions
  genus <- enquo(genus)
  species <- enquo(species)
  genSpp <- enquo(genSpp)

  if(!xor(is.null(get_expr(species)),is.null(get_expr(genSpp)))) stop('species OR genSpp column must be specified')

  if(is.null(get_expr(genSpp))){ #If genSpp column not specified
    d <- makeGenSpp(d,{{genus}},{{species}}) #Create genSpp column
    genSpp <- quo_set_expr(genSpp,expr(genSpp)) #Set genSpp expression
    genSpp <- quo_set_env(genSpp,quo_get_env(species)) #Set genSpp environment
  }

  #Colours for individual families
  famCols <- tibble(cols=brewer.pal(5,colourSet), #Colour scheme for families
                    {{family}} := c('Andrenidae','Apidae','Colletidae','Halictidae','Megachilidae'))

  #Species abundance plot

  #Data for histograms
  plotDat <- d %>% filter(!grepl('spp\\.$',{{genSpp}}) | keepSpp) %>%  #Filter out "spp" unless keepSpp==TRUE
    count({{family}},{{genSpp}}) %>% #Count family and genSpp occurrences
    arrange(desc({{family}}),n) %>% ungroup() %>% #Arrange by family
    mutate({{genSpp}}:=factor({{genSpp}},level={{genSpp}}))

  #Data for coloured background rectangles
  rectDat <- d %>% filter(!grepl('spp\\.$',{{genSpp}}) | keepSpp) %>%
    count({{family}},{{genSpp}}) %>%  #Count family and genSpp occurrences
    group_by({{family}}) %>% summarize(nSpp=n()) %>% ungroup() %>%
    arrange(desc({{family}})) %>% #Arrange by family
    mutate(ymax=cumsum(nSpp)+0.5,ymin=lag(ymax,default=0.5)) %>%
    rowwise() %>% mutate(ymid=mean(c(ymax,ymin))) %>%
    mutate(xmin=0,xmax=max(plotDat$n)) %>%
    left_join(famCols,by=as_label(family))

  #Species plot
  titleText <- paste0('Species (',nrow(plotDat),' total)')
  sppPlot <- ggplot()+ geom_col(data=plotDat,aes(n,{{genSpp}}))+ #Make columns
    geom_rect(data=rectDat, #Make background rectangles
              aes(xmin=xmin,xmax=xmax,ymin=ymin,ymax=ymax,fill={{family}}),
              alpha=0.3,show.legend = FALSE)+
    geom_col(data=plotDat,aes(n,{{genSpp}},fill={{family}}),show.legend = FALSE)+ #Columns (again)
    geom_text(data=rectDat,aes(x=xmax*0.9,y=ymid,label={{family}}),hjust=1)+ #Add text
    theme(axis.text.y=element_text(vjust=0.5,size=8*scaleYtext[1]))+ #Change theme
    labs(y=NULL,x='Number of specimens',title=titleText)+
    scale_fill_manual(values=rev(as.character(rectDat$cols)))

  #Data for genus abundance plots
  plotDat <- d %>%  #Data for histograms
    count({{family}},{{genus}}) %>% #Count family and genus occurrences
    arrange(desc({{family}}),n) %>% ungroup() %>%
    mutate({{genus}}:=factor({{genus}},level={{genus}})) #Re-order genus

  #Data for background rectangles
  rectDat <- d %>%
    group_by({{family}},{{genus}}) %>% summarize(n=n()) %>%
    group_by({{family}}) %>% summarize(nSpp=n()) %>% ungroup() %>%
    arrange(desc({{family}})) %>%
    mutate(ymax=cumsum(nSpp)+0.5,ymin=lag(ymax,default=0.5)) %>%
    rowwise() %>% mutate(ymid=mean(c(ymax,ymin))) %>%
    mutate(xmin=0,xmax=max(plotDat$n)) %>%
    left_join(famCols,by=as_label(family))

  #Genus plot
  titleText <- paste0('Genera (',nrow(plotDat),' total)')
  genPlot <- ggplot()+ geom_col(data=plotDat,aes(n,{{genus}}))+
    geom_rect(data=rectDat,
              aes(xmin=xmin,xmax=xmax,ymin=ymin,ymax=ymax,fill={{family}}),
              alpha=0.3,show.legend = FALSE)+
    geom_col(data=plotDat,aes(n,{{genus}},fill={{family}}),show.legend = FALSE)+
    geom_text(data=rectDat,aes(x=xmax*0.9,y=ymid,label={{family}}),hjust=1)+
    theme(axis.text.y=element_text(vjust=0.5,size=8*scaleYtext[2]))+
    labs(y=NULL,x='Number of specimens',title=titleText)+
    scale_fill_manual(values=rev(as.character(rectDat$cols)))

  #Family abundance plots
  plotDat <- d %>%  #Data for histograms
    group_by({{family}}) %>% summarize(n=n()) %>% ungroup() %>%
    arrange(desc({{family}}),n) %>%
    left_join(famCols,by=as_label(family)) %>%
    mutate({{family}}:=factor({{family}},level={{family}}))

  #Make family plot
  titleText <- paste0('Families (',nrow(plotDat),' total)')
  famPlot <- ggplot()+ geom_col(data=plotDat,aes(n,{{family}}))+ #Make plot
    geom_col(data=plotDat,aes(n,{{family}},fill={{family}}),show.legend = FALSE)+
    theme(axis.text.y=element_text(vjust=0.5,size=8*scaleYtext[3]))+
    labs(y=NULL,x='Number of specimens',title=titleText)+
    scale_fill_manual(values=as.character(plotDat$cols))

  if(returnPlots){
    #Put all plots together into a list
    a <- list(sppPlot,genPlot,famPlot)
  } else {
    #Put all plots together into a single plot
    a <- ggarrange(sppPlot,ggarrange(genPlot,famPlot,nrow=2),ncol=2)
  }
  return(a)
}
