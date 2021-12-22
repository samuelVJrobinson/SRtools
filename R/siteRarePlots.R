#' @title Site-based rarefaction plots
#' @description #Function to make "Linc-style" rarefaction plots using \pgk{vegan}
#    Takes matrix of spp abundance, with named sites for each row, and named spp for each column
#' @param d Matrix of species abundance (rows = species, cols = sites)
#' @param Ncol Number of columns for facets
#' @param Nrow Number of rows for facets
#' @param measType Type of diversity predictor ('both','Chao1','ACE','none')
#' @param textRange Proportion upper/lower bounds for text display. Can be 2 values overall, or 2 x N sites (min1,max1,min2,max2,...,minN,maxN).
#' @param seMult Multiplier for SE (default = 1)
#' @param rowOrder Should sites in facets be ordered in row order ('asis'), by diversity ('estDiv'), or # of samples ('Nsamp')?
#' @return A \code{ggplot} object
#' @export
#' @examples
#' library(SRtools)
#' data(dune,package='vegan')
#'
#' #Site-by-site
#' siteRarePlots(dune,Ncol = 5,Nrow=4,measType = 'Chao1',seMult = 1.96)
#'
#' #Overall
#' oDune <-t(as.matrix(colSums(dune)))
#' siteRarePlots(oDune,Ncol = 1,Nrow=1,measType = 'Chao1',seMult = 1.96)
siteRarePlots <- function(d,Ncol=NA,Nrow=NA,measType='both',textRange=c(0.1,0.4),seMult=1,rowOrder='asis'){
  require(vegan)
  require(ggplot2)
  require(dplyr)
  require(tidyr)
  require(tibble)

  #BS checking
  if(sum(c('both','Chao1','ACE','none') %in% measType)!=1) stop("measType must be one of 'both','Chao1','ACE','none'")

  #TO DO:
  # 1) How does function respond to empty rows in matrix?

  if(nrow(d)>1 & any(is.na(c(Ncol,Nrow)))){ #If columns not specified, and Nsites>1
    Ncol <- nrow(d); Nrow <- 1
  }

  if(is.null(rownames(d))) rownames(d) <- as.character(1:nrow(d)) #Create rownames (site names) if necessary

  #Get rarefaction data from 1:N samples
  rareDat <- data.frame(Species=unlist(lapply(1:nrow(d),
                                              function(x) as.vector(rarefy(d[x,],1:sum(d[x,]))))),
                        Sample=unlist(lapply(1:nrow(d),function(x) 1:sum(d[x,]))),
                        Site=unlist(lapply(rownames(d),function(x) rep(x,sum(d[x,]))))) %>%
    mutate(Site=factor(Site,levels=rownames(d)))

  graphText <- data.frame(Site=rownames(d),t(estimateR(d))) %>% #Data for Chao/ACE lines
    mutate(across(S.obs:se.ACE,~ifelse(is.nan(.),NA,.))) %>%
    select(-S.obs) %>% unite(chao,S.chao1,se.chao1,sep='_') %>%
    unite(ACE,S.ACE,se.ACE,sep='_') %>% pivot_longer(chao:ACE,names_to='index') %>%
    separate(value,c('meas','se'),sep='_',convert=T) %>%
    mutate(se=se*seMult) %>% #Multiply by se multiplier
    mutate(index=factor(index,labels=c('ACE','Chao1'))) %>%
    mutate(Site=factor(Site,levels=rownames(d)))

  graphText <- data.frame(graphText[rep(1:nrow(graphText),each=max(rareDat$Sample)),],
                          Sample=rep(1:max(rareDat$Sample),nrow(graphText)))

  graphText2 <- data.frame(N=rowSums(d),t(estimateR(d))) %>%
    rownames_to_column(var='Site') %>% mutate(Site=factor(Site,levels=Site)) %>%
    mutate(across(N:se.ACE,~ifelse(is.nan(.),NA,.))) %>%
    mutate(xpos=max(N)*0.95) %>% #X position
    mutate(ymax=switch(measType,
                       both=max(c(S.chao1+se.chao1,S.ACE+se.ACE,na.rm=TRUE)),
                       Chao1=max(S.chao1+se.chao1,na.rm=TRUE),
                       ACE=max(S.ACE+se.ACE,na.rm=TRUE))) %>%  #Maximum y value
    mutate(ymax=ifelse(is.na(ymax),S.obs,ymax)) %>% #If NAs made it through
    mutate_at(vars(S.chao1:se.ACE),function(x) trimws(format(round(x,2),nsmall=2))) #Trim display variables

  #Re-order sites as necessary
  newOrder <- switch(rowOrder,
                     Ndiv=order(apply(as.matrix(as.numeric(graphText2[,switch(measType,both=c(4,6),Chao1=c(4),ACE=c(6))])),1,max),
                                decreasing=TRUE),
                     Nsamp=order(graphText2$N,decreasing=TRUE),
                     asis=1:nrow(d))
  rareDat <- rareDat %>% mutate(Site=factor(Site,levels=levels(Site)[newOrder]))
  graphText2 <- graphText2 %>% mutate(Site=factor(Site,levels=levels(Site)[newOrder])) %>%
    arrange(Site)

  if(length(textRange)==2){ #If 2 range values provided
    graphText2 <- graphText2 %>%
      mutate(ylwr=ymax*min(textRange),yupr=ymax*max(textRange))
  } else if(length(textRange)==nrow(d)*2){ #If nrow(d)*2 range values provided (min,max,min,max,...)
    graphText2 <- graphText2 %>%
      mutate(ylwr=textRange[seq(1,nrow(d)*2-1,2)],
             yupr=textRange[seq(2,nrow(d)*2,2)]) %>%
      mutate(yupr=yupr*ymax,ylwr=ylwr*ymax)
  } else {
    stop('textRange must 2 or nrow(d)* 2 long')
  }


  p1 <- ggplot(data=rareDat)+
    facet_wrap(~Site,nrow=Nrow,ncol=Ncol) +
    geom_line(aes(Sample,Species),data=rareDat,size=1)+
    geom_point(data=graphText2,aes(x=N,y=S.obs),size=2)+
    geom_segment(data=graphText2,aes(x=N,y=S.obs,xend=N,yend=0),linetype='dashed')+
    geom_segment(data=graphText2,aes(x=0,y=S.obs,xend=N,yend=S.obs),linetype='dashed')+
    geom_text(data=graphText2,aes(x=xpos,y=yupr,label=paste('N =',N)),hjust=1,size=3)+
    geom_text(data=graphText2,aes(x=xpos,y=yupr-(yupr-ylwr)*ifelse(measType=='both',0.33,0.5),
                                  label=paste('S.obs =',S.obs)),hjust=1,size=3)

  if(measType=='both'){ #If using both richness indices
    p1 <- p1 +
      geom_line(data=graphText,aes(x=Sample,y=meas,col=index),size=1)+
      geom_ribbon(data=graphText,aes(x=Sample,ymax=meas+se,ymin=meas-se,fill=index),alpha=0.2)+
      geom_text(data=graphText2,aes(x=xpos,y=yupr-(yupr-ylwr)*0.66,
                                    label=paste('S.chao1 =',S.chao1,'\u00B1',se.chao1)),hjust=1,size=3)+
      geom_text(data=graphText2,aes(x=xpos,y=ylwr,label=paste('S.ACE =',S.ACE,'\u00B1',se.ACE)),hjust=1,size=3)+
      labs(x='Number of Specimens',y='Richness',col='Estimated\nSpecies\nRichness',
           fill='Estimated\nSpecies\nRichness')+
      scale_colour_manual(values=c('red','blue'))+
      scale_fill_manual(values=c('red','blue'))

  } else if(measType=='Chao1'|measType=='ACE') {

    p1 <- p1 +
      geom_line(data=filter(graphText,index==measType),aes(x=Sample,y=meas,col=index),size=1)+
      geom_ribbon(data=filter(graphText,index==measType),aes(x=Sample,ymax=meas+se,ymin=meas-se,fill=index),alpha=0.2)+
      labs(x='Number of Specimens',y='Richness',col='Estimated\nSpecies\nRichness',
           fill='Estimated\nSpecies\nRichness')+
      scale_colour_manual(values=c('black','black'))+
      scale_fill_manual(values=c('black','black'))

    if(measType=='Chao1'){
      p1 <- p1 + geom_text(data=graphText2,aes(x=xpos,y=ylwr,
                                               label=paste('S.chao1 =',S.chao1,'\u00B1',se.chao1)),hjust=1,size=3)
    } else {
      p1 <- p1 + geom_text(data=graphText2,aes(x=xpos,y=ylwr*0.66,
                                               label=paste('S.ACE =',S.ACE,'\u00B1',se.ACE)),hjust=1,size=3)
    }
  }
  return(p1)
}
