#' @title propAbundPlots
#' @description #Make proportional abundance bar-plots, with a rank-order smoothing line
#' @param d Data frame or tibble
#' @param s Species label column
#' @param N Species count column
#' @param textYRange Numeric vector with proportional y-range for plot text (default: c(0.5, 0.9))
#' @param by Faceting column - usually site (default: NULL)
#' @param Ncol Number of columns, if using facets
#' @param Nrow Number of rows, if using facets
#'
#' @return A \code{ggplot} object
#' @export
#'
#' @examples
#' data(dune,package='vegan')
#'
#' library(tidyverse)
#'
#' dune <- dune %>% mutate(site=1:nrow(.)) %>%
#'   pivot_longer(cols = -site,names_to='species',values_to='counts')
#'
#'   propAbundPlots(dune,s = species,N = counts)
#'   filter(dune,site<10) %>% #First 9 sites
#'     propAbundPlots(s = species,N = counts, by = site, Nrow=3, Ncol=3)
#'
propAbundPlots <- function(d,s,N,textYRange=c(0.5,0.9),by=NULL,Ncol=NA,Nrow=NA){
  require(ggplot2); require(dplyr); require(tidyr);
  require(rlang); require(vegan)
  options(dplyr.summarise.inform=FALSE) #Suppress dplyr messages

  s <- enquo(s) #Defuse expressions
  N <- enquo(N)
  by <- enquo(by)

  #If no facet specified
  if(is.null(get_expr(by))){
    isFacet <- FALSE
    by <- quo_set_expr(by,expr(Site)) #Set facet expression ("Site")
    by <- quo_set_env(by,quo_get_env(s)) #Set facet environment
    d <- d %>% mutate({{by}}:='Site1') #Create column with single site
  } else {
    isFacet <- TRUE #Facets are needed
  }

  #Reshape data into matrix, used in estimateR
  dMat <- d %>%
    group_by({{by}},{{s}}) %>%
    summarize(N = sum({{N}})) %>%
    pivot_wider(names_from={{s}},values_from=N) %>%
    column_to_rownames(var=as_label({{by}})) %>% ungroup() %>%
    as.matrix()

  #Check for empty sites or species
  if(any(rowSums(dMat)==0)){
    zeroSpp <- rownames(dMat)[rowSums(dMat)==0] #Sites that are empty
    dMat <- dMat[rowSums(dMat)!=0,] #Remove rows (sites) with zero counts
    d <- d %>% filter(!{{by}} %in% zeroSpp) #Remove empty species
    warning('Some sites have zero counts and were removed: ',paste0(zeroSpp,collapse=', '))
  }
  if(any(colSums(dMat)==0)){
    zeroSpp <- colnames(dMat)[colSums(dMat)==0] #Species that aren't represented in dataset
    dMat <- dMat[,colSums(dMat)!=0] #Remove columns (species) with zero counts
    d <- d %>% filter(!{{s}} %in% zeroSpp) #Remove empty species
    warning('Some species have zero counts and were removed: ',paste0(zeroSpp,collapse=', '))
  }

  uniqueSpp <- d %>% pull({{s}}) %>% unique()

  #Text for figures
  graphText <- d %>% group_by({{by}}) %>% summarize(N=sum({{N}})) %>% ungroup() %>%
    #Attaches diversity estimates from estimateR (from vegan)
    bind_cols(as.data.frame(t(estimateR(dMat)))) %>%
    mutate(across(everything(),~ifelse(is.nan(.),NA,.))) %>%
    #Trims white space and rounds to nearest 2 decimals
    mutate_at(vars(S.chao1:se.ACE),function(x) trimws(format(round(x,2),nsmall=2))) %>%
    mutate_at(vars(S.chao1:se.ACE),~ifelse(.x=='0.00','0.01',.x)) %>%
    #Position for labels
    mutate(xpos=as.numeric(factor({{by}}))*length(uniqueSpp)) %>%
    mutate({{by}}:=factor({{by}},labels=rownames(dMat)))

  #Proportion abundance
  d1 <- d %>% mutate({{by}}:=as.numeric(factor({{by}}))) %>% group_by({{by}},{{s}}) %>%
    summarize(count=sum({{N}})) %>%
    group_by({{by}}) %>%
    mutate(Proportion=count/sum(count)) %>% #Proportion of sample
    ungroup() %>% arrange({{by}},desc(Proportion)) %>% #Sort by facet and proportion
    mutate(r=row_number()) %>% #Row number as explicit variable
    group_by({{by}}) %>% mutate(ord=row_number()) %>% #Sort within each sample
    mutate(predProp=exp(predict(lm(log(count+1)~log(ord))))/sum(count)) %>% #Predicted proportion
    ungroup() %>%
    mutate({{by}}:=factor({{by}},levels=c(1:nrow(dMat)),labels=rownames(dMat))) #Re-label sample

  ypos <- seq(max(d1$Proportion)*textYRange[2],max(d1$Proportion)*textYRange[1],length.out=3)

  brk <- d1 %>% pull(r) #Set breaks to ranks
  spp <- d1 %>% pull({{s}}) #Set labels to Species

  #Create plot
  p1 <- ggplot(d1,aes(r,Proportion))+ #Use rank rather than spp label
      geom_col(fill='black')+ #Proportion=column height
      geom_smooth(method='nls',se=F,formula=y~exp(b0+b1*log(x+1)), #Nonlinear smoother
                  method.args=list(start=c(b0=1,b1=1)),col='red')+
      scale_x_continuous(breaks = brk, labels = spp,
                         expand = expansion(0,0.2))+ #Expand text size
      theme(axis.text.x=element_text(angle=90,vjust=0.4))+ #Rotate x-axis text
      coord_cartesian(ylim=range(d1$Proportion))+ #Cut off smoother at 0,0.7
      #Add spp richness text
      geom_text(da=graphText,aes(x=xpos,y=ypos[1],label=paste('N =',N)),hjust=1,size=3)+
      geom_text(da=graphText,aes(x=xpos,y=ypos[2],label=paste('S.obs =',S.obs)),hjust=1,size=3)+
      geom_text(da=graphText,aes(x=xpos,y=ypos[3],label=paste('S.Chao1 =',S.chao1,'\u00B1',se.chao1)),
                hjust=1,size=3)+
      labs(x='Species',y='Proportional Abundance')

  if(isFacet){ #If facets are being used
    if(any(is.na(c(Ncol,Nrow)))){ #If columns or rows not specified
      Ncol <- nrow(dMat) #Multiple columns
      Nrow <- 1 #Single row
    }
    p1 <- p1 + facet_wrap({{by}},scales='free_x',ncol=Ncol,nrow=Nrow) #Use free x axis
  }

  return(p1)
}
