#' @title Print insect labels
#' @description Function to print insect labels from a set of csv files. Written for labels used in the U of C insect collection and Galpern lab.
#' @param inputPath Input file with number of labels needed (csv)
#' @param csvPaths Paths to site and trap tables (vector with 2 paths - csv)
#' @param outputPath Output path (docx)
#' @param sepChar Separator characters for \code{inputPath}, \code{site}, and \code{trap} table csvs (default: rep(',',3))
#' @export
#' @details \strong{\code{inputPath}} is a path to a csv with the following column headers:
#'
#' \emph{BTID} (trapping ID), \emph{N} (number of labels needed)
#'
#' \strong{\code{csvPaths[1]}} is a path to the \code{site} csv with the following column headers:
#'
#' \emph{BLID}, \emph{lat}, \emph{lon}, \emph{locality}, \emph{elevation}, \emph{country}, \emph{province}
#'
#' \strong{\code{csvPaths[2]}} is a path to the \code{trap} csv with the following column headers:
#'
#' \emph{BTID}, \emph{BLID}, \emph{collector}, \emph{trapType}, \emph{startYear}, \emph{startMonth} (numeric), \emph{startDay}, \emph{endYear}, \emph{endMonth} (numeric), \emph{endDay}, \emph{lonTrap}, \emph{latTrap}, \emph{elevTrap}
#'
#' If \emph{lonTrap}, \emph{latTrap}, or \emph{elevTrap} values are missing from \code{csvPaths[2]}, values will be filled using \code{site} csv.
#'
#' \strong{\code{output[1]}} is a path to a \emph{docx} file that will be printed. If more than 1 page of labels is required, additional files will be created with numbers appended before the file extension (e.g. \emph{labels_01.docx}, \emph{labels_02.docx})
#'
#' @examples
#' # Make csv files for labels
#'
#' # 2 trapping sessions, with 20 labels each
#' input <- data.frame(BTID = c('10069-1-A-2015','51273-0-B-2019'),
#'                     N = c(20,20))
#'
#' # Site table
#' site <- data.frame(BLID=c(10069,51273),lat=c(50,51),lon=c(-114,-113),
#'                    locality=c('Near Town','Near Mountains'),
#'                    elevation=c(800,1200),country=rep('CANADA',2),
#'                    province=rep('Alberta',2))
#'
#' #Trap table
#' trap <- data.frame(BTID=c('10069-1-A-2015','51273-0-B-2019'),
#'                    BLID=c(10069,51273),collector=c('D. Smith','L. Robbins'),
#'                    trapType=c('Blue Vane','Netting'),
#'                    startYear=c(2012,2021),startMonth=c(6,7),startDay=c(10,20),
#'                    endYear=c(2012,NA),endMonth=c(6,NA),endDay=c(15,NA),
#'                    lonTrap=c(-114.5,-113.1),latTrap=c(51,52.3),
#'                    elevTrap=c(815,1231))
#'
#' #Write csvs to file
#' write.csv(input,'inputcsv.csv')
#' write.csv(site,'sitecsv.csv')
#' write.csv(trap,'trapcsv.csv')
#'
#' #Create labels
#' makeLabels(inputPath = 'inputcsv.csv',
#'            csvPaths = c('sitecsv.csv','trapcsv.csv'),
#'            outputPath = 'labelSheet.docx'
#' )

makeLabels <- function(inputPath,csvPaths,outputPath,sepChar=rep(',',3)){

  require(tidyverse)
  require(officer) # https://ardata-fr.github.io/officeverse/officer-for-word.html

  i2c <- 2.54 #Convert inches to cm

  input <- read.csv(inputPath,strip.white = TRUE,sep=sepChar[1]) %>%
    mutate(across(everything(),~gsub('(^\\s|\\s$)','',.x))) %>% #Strip white space
    mutate(N=as.numeric(N)) %>%
    separate(BTID,c('BLID','Pass','Trap','Year'),sep='-',remove = FALSE)

  #Read in site and trap tables
  site <- read.table(csvPaths[1],header=TRUE,sep=sepChar[2],stringsAsFactors = FALSE,strip.white = TRUE) #%>%
  # mutate(across(everything(),~gsub('(^\\s|\\s$)','',.x))) %>% #Strip white space

  trap <- read.table(csvPaths[2],header=T,sep=sepChar[3],stringsAsFactors = F,strip.white = TRUE) #%>%
  # mutate(across(everything(),~gsub('(^\\s|\\s$)','',.x))) %>% #Strip white space
  # mutate(across(c(pass,startYear:endMinute,lonTrap:elevTrap,startjulian:deployedhours)))

  #BS checking
  if(ncol(site)==1) stop('Site table has only 1 column. Fix separator character?')
  if(ncol(trap)==1) stop('Trap table has only 1 column. Fix separator character?')

  checkBLID <- !(input$BLID %in% site$BLID) & !(input$BLID %in% trap$BLID) #BLID - site IDs
  if(any(checkBLID)){
    print(paste0('BLIDs not found in site/trap tables:\n ',paste(input$BLID[checkBLID],collapse = ',')))
    x <- readline('Continue, and skip entries? (Y/N): ')
    if(x=='Y'|x=='y'){
      input <- input[!checkBLID,] #Remove entries without matching BLIDs
    } else {
      print('Quitting program')
      return()
    }
  } else {
    print('BLIDs OK')
  }
  checkBTID <- !(input$BTID %in% trap$BTID) #BTID - trapping IDs
  if(any(checkBTID)){
    print(paste0('BTIDs not found in trap table: ',paste(input$BTID[checkBTID],collapse = ',')))
    x <- readline('Continue, and skip entries? (Y/N): ')
    if(x=='Y'|x=='y'){
      input <- input[!checkBTID,] #Remove entries without matching BLIDs
    } else {
      print('Quitting program')
      return()
    }
  } else {
    print('BTIDs OK')
  }

  if(any(!(c('country','locality') %in% colnames(site)))){
    stop('"country" and "locality" columns must be present in site table')
  }

  #Page formating properties
  setup <- block_section(
    prop_section(
      page_size = page_size(width = 8.5, height = 11, orient = 'portrait'),
      page_margins = page_mar(bottom=0.4,top=0.5,right=0.5,left=0.5,header=0,footer=0,gutter=0),
      type = 'continuous',
      section_columns = section_columns(widths=rep(2.16/i2c,7),space=0.5/i2c, sep = FALSE)
    )
  )

  parProp <- fp_par(line_spacing=0.35) #Paragraph formatting properties
  textProp <- fp_text(font.family = "Arial",font.size=4) #Text formatting properties
  textProp_superscript <- textProp
  textProp_superscript$vertical.align <- 'superscript'

  #Get matching indices from site and trap tables
  input$siteMatches <- match(input$BLID,site$BLID)
  input$trapMatches <- match(input$BTID,trap$BTID)

  labNum <- 0 #Initialize label number
  inFormat <- '%d.%m.%Y'; outFormat <- '%d.%b.%Y' #Date format strings
  chopLast <- function(x,n) substr(x,1,nchar(x)-n) #Function to remove last n characters from a string

  textBlockList <- lapply(1:nrow(input),function(i){

    sM <- input$siteMatches[i] #Site/trap table matching indices
    tM <- input$trapMatches[i]

    #Site-level parameters
    countryText <- fpar(ftext(paste0(site$country[sM],': ',site$province[sM]),prop=textProp),fp_p = parProp) #Country/province
    locText <- fpar(ftext(site$locality[sM],prop=textProp),fp_p = parProp) #Locality
    degMark <- ftext('o',prop = textProp_superscript) #Degree mark for lat-lon text

    #Trap-level parameters
    trapElev <- trap$elevTrap[tM] #Elevation
    if(trapElev==0 || is.na(trapElev)){ #Elevation not found in trap table
      trapElev <- site$elevation[sM]
      if(trapElev==0 || is.na(trapElev)){
        stop(paste0('No elevation found for BTID: ',input$BTID[i],' in either site or trap table'))
      }
    }
    elevText <- fpar(ftext(paste0('Elev. ',trapElev,' m'),prop=textProp),fp_p = parProp) #Elevation

    #Latitude/longitude
    latLon <- c(as.character(trap$latTrap[tM]),as.character(-trap$lonTrap[tM]))
    if(any(latLon=='0') || any(is.na(latLon))){ #Lat/lon not found in trap table
      latLon <- c(as.character(site$lat[sM]),as.character(-site$lon[sM]))
      if(any(latLon=='0') || any(is.na(latLon))){ #Lat/lon not found in site table
        stop(paste0('No latitude/longitude found for BTID: ',input$BTID[i],' in either site or trap table'))
      }
    }
    latLonText <- fpar(ftext(latLon[1],prop=textProp), degMark, ftext('N, ',prop=textProp), #Latitude/longitude text
                       ftext(latLon[2],prop=textProp), degMark, ftext('W',prop=textProp), fp_p = parProp)

    collectorText <- fpar(ftext(paste0(trap$trapType[tM],'; ',trap$collector[tM]),prop=textProp),fp_p = parProp) #Collection type/collector
    btidText <- fpar(ftext(trap$BTID[tM],prop=textProp),fp_p = parProp)

    #Collection dates
    startText <- toupper(format(as.Date(paste(trap$startDay[tM],trap$startMonth[tM],trap$startYear[tM],sep='.'),
                                        format = inFormat),format = outFormat))
    startText <- ifelse(is.na(startText),'',startText) #If no starting date, shows only end date

    endText <- toupper(format(as.Date(paste(trap$endDay[tM],trap$endMonth[tM],trap$endYear[tM],sep='.'),
                                      format = inFormat),format = outFormat))
    endText <- ifelse(is.na(endText),'',endText) #If no ending date, shows only end date

    if(nchar(startText)>0 & nchar(endText)>0){ #If both dates found
      startEndText <- paste0(chopLast(startText,5),'-',endText) #Start to end of collection
    } else if(nchar(startText)==0 & nchar(endText)==0){ #If no dates found
      warning(paste0('No collection dates found for BTID: ',input$BTID[i]))
      startEndText <- 'No Collection Date'
    } else if(xor(nchar(startText)>0,nchar(endText)>0)){
      startEndText <- paste0(startText,endText) #Only single date listed
    }
    dateText <- fpar(ftext(startEndText,prop=textProp),fp_p = parProp) #Date text

    sepText <- fpar(ftext(' ',prop=textProp),fp_p = parProp) #fp_par(line_spacing=0.5)) #Empty separator text

    #Create block list

    #Text for each label
    entryText <- block_list(countryText,locText,elevText,latLonText,
                            collectorText,btidText,dateText,sepText)

    textBlock <- rep(entryText,input$N[i]) #Replicate over each label

    class(textBlock) <- c('block_list','block') #Reclass as block list

    return(textBlock)
  })

  textBlockList <- do.call('c',textBlockList)

  print(paste0('Making set of ',sum(input$N),' labels'))

  # Labels have 7 text lines + blank line at end = 8 lines
  labLineNum <- 8
  # Page with 0.5" margins can fit 21 x 7 matrix of labels = 147 total
  pageCols <- 7
  pageRows <- 21
  pageLabNum <- pageCols*pageRows

  # Every 21 x 8 = 168 lines should be replaced with a column break on full pages
  colBreakLines <- pageRows*labLineNum

  # On non-full pages, column breaks should make labels span all 7 columns

  #BS Check - number of labels x lines per label should == length(textBlockList)
  if(sum(input$N)*labLineNum != length(textBlockList)) stop("Problem with number of text lines: N_labels x labLineNum != length(textBlockList)")

  fullPages <- sum(input$N) %/% pageLabNum #Full pages of labels
  leftoverLabs <- sum(input$N) %% pageLabNum #Leftover labels

  fullPageSepText <- cumsum(rep(colBreakLines,fullPages*pageCols)) #Separators on full pages
  lastPageSepText <- c( #Number of labels per column
    rep(leftoverLabs %/% pageCols + 1, leftoverLabs %% pageCols),
    rep(leftoverLabs %/% pageCols, pageCols-(leftoverLabs %% pageCols))
  )
  lastPageSepText <- lastPageSepText[1:(pageCols-1)] #Remove last row (no linebreak needed on last column)
  lastPageSepText <- cumsum(lastPageSepText*labLineNum) #Convert to lines

  #If full pages exist, increment to last element of lastPageSepText
  if(length(fullPageSepText)>0) lastPageSepText <- lastPageSepText + fullPageSepText[length(fullPageSepText)]
  allSepText <- c(fullPageSepText,lastPageSepText) #Join together

  if(length(allSepText)>0){ #Insert column breaks
    for(j in allSepText) textBlockList[[j]] <- fpar(ftext('',prop=textProp),
                                                    run_columnbreak(),
                                                    fp_t = textProp,
                                                    fp_p = fp_par(line_spacing=0.5))
  }

  #Problem: body_add_blocks takes a huge amount of time to create files (takes about 1 minute for a single page of labels)
  # Workaround exists on Windows, but not on Mac or Linux - save to individual docx files, then sew together
  # Current workaround: append numbers to the end of outputPath as needed, and create individual docx pages to print

  # #Old solution
  # class(textBlockList) <- c('block_list','block') #Reclass as block_list
  # template <- read_docx() %>% #Read in blank document
  #   body_add_fpar(fpar(' ',fp_t = textProp,fp_p = fp_par(line_spacing=0.5))) %>%  #Add extra line to beginning
  #   body_add_blocks(blocks = textBlockList) %>% #Add textBlock
  #   body_end_block_section(setup)
  # print(template,target=outputPath) #Prints to file at end of label set

  #New solution
  if(fullPages>0 & leftoverLabs>0){
    oPath <- gsub('.docx','',outputPath) #Strip file extension
    nDocs <- sum(fullPages) + ifelse(leftoverLabs>0,1,0) #Number of single-page files to create
    outputPath <- paste0(oPath,'_',sprintf('%02d',1:nDocs),'.docx') #Appends digits to output path
  }

  for(i in 1:length(outputPath)){

    print(paste0('Making label page ',i,' of ',length(outputPath)))

    tBlock <- textBlockList[1:(pmin(pageLabNum*labLineNum,length(textBlockList)))] #Select 1 page-worth of labels
    tBlock <- tBlock[-length(tBlock)] #Gets rid of last run_columnbreak
    class(tBlock) <- c('block_list','block') #Reclass as block_list

    template <- read_docx() %>% #Read in blank document
      body_add_fpar(fpar(' ',fp_t = textProp,fp_p = fp_par(line_spacing=0.5))) %>% #Add extra space to beginning
      body_add_blocks(blocks = tBlock) %>% #Add textBlock
      body_end_block_section(setup)
    print(template,target=outputPath[i]) #Prints to file at end of label set

    if(i!=length(outputPath)){ #If not at end of list
      textBlockList <- textBlockList[(pageLabNum*labLineNum+1):(length(textBlockList))] #Remove printed section of textBlockList
    }
  }
  print('Finished making labels')
}
