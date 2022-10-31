library(testthat)
# context('makeLabels')

# test_that("multiplication works", {
#   expect_equal(2 * 2, 4)
# })

# Make csv files for labels

# 2 trapping sessions, with 20 labels each
input <- data.frame(BTID = c('10069-1-A-2015','51273-0-B-2019'),
                    N = c(20,20))

# Site table
site <- data.frame(BLID=c(10069,51273,99999),lat=c(50,51,99),lon=c(-114,-113,99),
                   locality=c('Near Town','Near Mountains','Nowhere'),
                   elevation=c(800,1200,999),country=rep('CANADA',3),
                   province=rep('Alberta',3))

#Trap table - adding an extra unused BTID
trap <- data.frame(BTID=c('10069-1-A-2015','51273-0-B-2019','99999-9-9-9999'),
                   BLID=c(10069,51273,99999),collector=c('D. Smith','L. Robbins','Nobody'),
                   trapType=c('Blue Vane','Netting','Nothing'),
                   startYear=c(2012,2021,9999),startMonth=c(6,7,9),startDay=c(10,20,9),
                   endYear=c(2012,NA,NA),endMonth=c(6,NA,NA),endDay=c(15,NA,NA),
                   lonTrap=c(-114.5,-113.111111111,-99),latTrap=c(51,52.3333333333,99),
                   elevTrap=c(815,1231,999))

#Write csvs to file
write.csv(input,'inputcsv.csv',row.names = FALSE)
write.csv(site,'sitecsv.csv',row.names = FALSE)
write.csv(trap,'trapcsv.csv',row.names = FALSE)

#Create labels
test_that('Basic function works',{
  expect_output(
  makeLabels(inputPath = 'inputcsv.csv',
             csvPaths = c('sitecsv.csv','trapcsv.csv'),
             outputPath = 'labelSheet.docx',
             platform = 'LibreOffice')
  )
})

saveThis <- input

input <- saveThis
names(input) <- c('a','b')
write.csv(input,'inputcsv.csv',row.names = FALSE)
test_that('Column name error',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

input <- saveThis
input$extra <- 1
write.csv(input,'inputcsv.csv',row.names = FALSE)
test_that('Extra column warning',{
  expect_warning(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
    )})


input <- saveThis
input$N[2] <- 0
write.csv(input,'inputcsv.csv',row.names = FALSE)

test_that('Warns about zero labels',{
  expect_warning(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

input <- saveThis
input$BTID[1] <- ''
write.csv(input,'inputcsv.csv',row.names = FALSE)

test_that('Empty BTID error',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

input <- saveThis
input$BTID[1] <- '10069-1-A'
write.csv(input,'inputcsv.csv',row.names = FALSE)

test_that('Bad BTID error',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

input <- saveThis
write.csv(input,'inputcsv.csv',row.names = FALSE)

#Tests for site and trap csvs

saveThis <- site
site[4,] <- rep(NA,ncol(site))
write.csv(site,'sitecsv.csv',row.names = FALSE)

test_that('Site table blank rows',{
  expect_warning(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

site <- saveThis
site$BLID[1] <- NA
write.csv(site,'sitecsv.csv',row.names = FALSE)

test_that('Site table missing BLIDs',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
  )
})

site <- saveThis
site$country[1] <- NA
write.csv(site,'sitecsv.csv',row.names = FALSE)

test_that('Site table missing country',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
  )
})

site <- saveThis
write.csv(site,'sitecsv.csv',row.names = FALSE)

#Trap table
saveThis <- trap
trap[4,] <- NA
write.csv(trap,'trapcsv.csv',row.names = FALSE)

test_that('Trap table blank rows',{
  expect_warning(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word'))
})

trap <- saveThis
trap$BLID[1] <- NA
write.csv(trap,'trapcsv.csv',row.names = FALSE)

test_that('Trap table missing BLIDs',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
  )
})

trap <- saveThis
trap$BTID[1] <- NA
write.csv(trap,'trapcsv.csv',row.names = FALSE)

test_that('Trap table missing BTIDs',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
  )
})


trap <- saveThis
trap$collector[2] <- NA
write.csv(trap,'trapcsv.csv',row.names = FALSE)

test_that('Trap table missing collector',{
  expect_error(
    makeLabels(inputPath = 'inputcsv.csv',
               csvPaths = c('sitecsv.csv','trapcsv.csv'),
               outputPath = 'labelSheet.docx',
               platform = 'Word')
  )
})

#Cleanup
file.remove(c('inputcsv.csv','sitecsv.csv','trapcsv.csv','labelSheet.docx'))

