###demo read in and paste tracking report
# Bijan Warner
# 9/16/2017

library(XLConnect)
library(stringr)
library(plyr)
library(dplyr)
library(reshape2)

###Read in and prep csv (already conditioned)

tfile <- read.csv("Short Salve Prep 20170916.csv")

##filter

table(tfile$year)


##Filter to only last 3 years of freshmen
tfile <- tfile[tfile$year>=2015,]
tfile <- tfile[tfile$freshman ==1,]


####CREATE table objects, transpose to have year as column
t <- as.data.frame(t
                   (tfile %>%
                       group_by(year) %>%
                       summarise(applicants = sum(applicant),
                                 admits = sum(admit),matrics = sum(matric))))

ncf_i <- as.data.frame((tfile %>%
                          group_by(year,INCOME) %>%
                          summarise(applicants = sum(applicant),
                                    admits = sum(admit),
                                    matrics = sum(matric))))

m <- melt(ncf_i,id=c("year","INCOME"))
m1 <- reshape(m, v.names = "value", idvar = c("INCOME","variable"), timevar = "year",direction = "wide")
m2 <- reshape(m1, idvar = "INCOME", timevar ="variable", direction = "wide")



#### NOW PREP AND LOAD WORKBOOK

#define workbook object (in active directory); but do not over-write (must already exist)
wb <- loadWorkbook("tracking.xlsx", create = FALSE)

#do not overwrite existing style in Excel template
setStyleAction(wb,XLC$"STYLE_ACTION.NONE")

# Create a worksheet
createSheet(wb, name = "paste")
createSheet(wb, name = "NCF_Overview")

# Create a name reference
createName(wb, name = "paste", formula = "paste!$D$2")
createName(wb, name = "ncf", formula = "NCF_Overview!$D$2")


# Write   data.frame '  to the specified named region
writeNamedRegion(wb, t, name = "paste", header=TRUE, rownames="row.names")
writeNamedRegion(wb, m2, name = "ncf", header=TRUE, rownames="row.names")

# Save workbook
saveWorkbook(wb)