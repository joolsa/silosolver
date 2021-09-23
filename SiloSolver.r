## Name: FB Data Analysis
## Version: 0.0.3
## Last update: 2017-05-02

# installing required R pachages (just check for the very first run)
list.of.packages <- c('xlsx', 'tidyr', 'dplyr')
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[, 'Package'])]
if (length(new.packages) > 0) install.packages(new.packages)

# include package namespaces
library(xlsx)
library(tidyr)
library(dplyr)

# read data
theFileByAgeGender <- 'DATA BY AGE AND GENDER.csv'
theFileByRegion <- 'DATA BY REGION.csv'
theDataAG <- read.csv(theFileByAgeGender, header = T, sep = ';', stringsAsFactors = F, encoding = 'latin1') %>%
  dplyr::rename(Amount.Spent = Amount.Spent..EUR.) %>%
  dplyr::rename(Clicks.All = Clicks..All.)
theDataR <- read.csv(theFileByRegion, header = T, sep = ';', stringsAsFactors = F, encoding = 'latin1') %>%
  dplyr::rename(Amount.Spent = Amount.Spent..EUR.) %>%
  dplyr::rename(Clicks.All = Clicks..All.)

# calculate pcts
theDataAG_Clean <- theDataAG %>%
  dplyr::select(Campaign.Name, Link.Clicks, Clicks.All, Impressions, Amount.Spent, Website.Purchases, Total.Conversion.Value, Age, Gender) %>%
  dplyr::mutate_all(funs(replace(., is.na(.), 0))) %>%
  dplyr::group_by(Campaign.Name) %>%
  dplyr::mutate(pctAG.Link.Clicks = Link.Clicks / sum(Link.Clicks, na.rm = T)) %>%
  dplyr::mutate(pctAG.Clicks.All = Clicks.All / sum(Clicks.All, na.rm = T)) %>%
  dplyr::mutate(pctAG.Impressions = Impressions / sum(Impressions, na.rm = T)) %>%
  dplyr::mutate(pctAG.Amount.Spent = Amount.Spent / sum(Amount.Spent, na.rm = T)) %>%
  dplyr::mutate(pctAG.Website.Purchases = Website.Purchases / sum(Website.Purchases, na.rm = T)) %>%
  dplyr::mutate(pctAG.Total.Conversion.Value = Total.Conversion.Value / sum(Total.Conversion.Value, na.rm = T)) %>%
  dplyr::mutate_all(funs(replace(., is.na(.), 0))) %>%
  dplyr::ungroup()
theDataR_Clean <- theDataR %>%
  dplyr::select(Campaign.Name, Link.Clicks, Clicks.All, Impressions, Amount.Spent, Website.Purchases, Total.Conversion.Value, Region) %>%
  dplyr::mutate_all(funs(replace(., is.na(.), 0))) %>%
  dplyr::group_by(Campaign.Name) %>%
  dplyr::mutate(pctR.Link.Clicks = Link.Clicks / sum(Link.Clicks, na.rm = T)) %>%
  dplyr::mutate(pctR.Clicks.All = Clicks.All / sum(Clicks.All, na.rm = T)) %>%
  dplyr::mutate(pctR.Impressions = Impressions / sum(Impressions, na.rm = T)) %>%
  dplyr::mutate(pctR.Amount.Spent = Amount.Spent / sum(Amount.Spent, na.rm = T)) %>%
  dplyr::mutate(pctR.Website.Purchases = Website.Purchases / sum(Website.Purchases, na.rm = T)) %>%
  dplyr::mutate(pctR.Total.Conversion.Value = Total.Conversion.Value / sum(Total.Conversion.Value, na.rm = T)) %>%
  dplyr::mutate_all(funs(replace(., is.na(.), 0))) %>%
  dplyr::mutate(Link.Clicks = NULL, Clicks.All = NULL, Impressions = NULL, Amount.Spent = NULL, Website.Purchases = NULL, Total.Conversion.Value = NULL) %>%
  dplyr::ungroup()

# merging
theData <- NULL
for(icn in 1:length(unique(theDataAG_Clean$Campaign.Name))) {
  CN <- unique(theDataAG_Clean$Campaign.Name)[icn]
  tmpDataAGCN_Clean <- theDataAG_Clean[theDataAG_Clean$Campaign.Name == CN, ]
  tmpDataAGCN <- theDataAG[theDataAG$Campaign.Name == CN, ]
  tmpDataAGCN_Clean$Link.Clicks <- sum(as.numeric(tmpDataAGCN$Link.Clicks), na.rm = T)
  tmpDataAGCN_Clean$Clicks.All <- sum(as.numeric(tmpDataAGCN$Clicks.All), na.rm = T)
  tmpDataAGCN_Clean$Impressions <- sum(as.numeric(tmpDataAGCN$Impressions), na.rm = T)
  tmpDataAGCN_Clean$Amount.Spent <- sum(as.numeric(tmpDataAGCN$Amount.Spent), na.rm = T)
  tmpDataAGCN_Clean$Website.Purchases <- sum(as.numeric(tmpDataAGCN$Website.Purchases), na.rm = T)
  tmpDataAGCN_Clean$Total.Conversion.Value <- sum(as.numeric(tmpDataAGCN$Total.Conversion.Value), na.rm = T)
  tmpDataRCN_Clean <- theDataR_Clean[theDataR_Clean$Campaign.Name == CN, ]
  for(ia in 1:length(unique(tmpDataAGCN_Clean$Age))) {
    A <- unique(tmpDataAGCN_Clean$Age)[ia]
    for(ig in 1:length(unique(tmpDataAGCN_Clean[tmpDataAGCN_Clean$Age == A, ]$Gender))) {
      G <- unique(tmpDataAGCN_Clean[tmpDataAGCN_Clean$Age == A, ]$Gender)[ig]
      tmpData <- dplyr::full_join(tmpDataAGCN_Clean[tmpDataAGCN_Clean$Age == A & tmpDataAGCN_Clean$Gender == G, ], tmpDataRCN_Clean)
      theData <- rbind(theData, tmpData)
    }
  }
}
# averaging
theDataAvg <- theData %>%
  dplyr::mutate(Link.Clicks = Link.Clicks * pctAG.Link.Clicks * pctR.Link.Clicks) %>%
  dplyr::mutate(Clicks.All = Clicks.All * pctAG.Clicks.All * pctR.Clicks.All) %>%
  dplyr::mutate(Impressions = Impressions * pctAG.Impressions * pctR.Impressions) %>%
  dplyr::mutate(Amount.Spent = Amount.Spent * pctAG.Amount.Spent * pctR.Amount.Spent) %>%
  dplyr::mutate(Website.Purchases = Website.Purchases * pctAG.Website.Purchases * pctR.Website.Purchases) %>%
  dplyr::mutate(Total.Conversion.Value = Total.Conversion.Value * pctAG.Total.Conversion.Value * pctR.Total.Conversion.Value) %>%
  dplyr::mutate(pctAG.Link.Clicks = NULL, pctAG.Clicks.All = NULL, pctAG.Impressions = NULL, pctAG.Amount.Spent = NULL, pctAG.Website.Purchases = NULL, pctAG.Total.Conversion.Value = NULL) %>%
  dplyr::mutate(pctR.Link.Clicks = NULL, pctR.Clicks.All = NULL, pctR.Impressions = NULL, pctR.Amount.Spent = NULL, pctR.Website.Purchases = NULL, pctR.Total.Conversion.Value = NULL)

# adding calculated metrics
theDataAvg <- theDataAvg %>%
  dplyr::mutate(CPC = Amount.Spent / Link.Clicks) %>%
  dplyr::mutate(CTR = Link.Clicks / Impressions) %>%
  dplyr::mutate(CVR = Website.Purchases / Link.Clicks)

# cleaning  
theDataAvg <- theDataAvg %>%  
  dplyr::mutate_all(funs(replace(., is.infinite(.), NaN))) %>%
  dplyr::filter(Age != 'Unknown' & Age != 'unknown' & Age != '') %>%
  dplyr::filter(Gender != 'Unknown' & Gender != 'unknown' & Gender != '') %>%
  dplyr::filter(Region != 'Unknown' & Region != 'unknown' & Region != '')

# summarizing data across campaigns
theDataSum <- theDataAvg %>%
  dplyr::group_by(Age, Gender, Region) %>%
  dplyr::mutate(N = n()) %>%
  dplyr::mutate(Sum.Link.Clicks = sum(as.numeric(Link.Clicks), na.rm = T)) %>%
  dplyr::mutate(Sum.Clicks.All = sum(as.numeric(Clicks.All), na.rm = T)) %>%
  dplyr::mutate(Sum.Impressions = sum(as.numeric(Impressions), na.rm = T)) %>%
  dplyr::mutate(Sum.Amount.Spent = sum(as.numeric(Amount.Spent), na.rm = T)) %>%
  dplyr::mutate(Sum.Website.Purchases = sum(as.numeric(Website.Purchases), na.rm = T)) %>%
  dplyr::mutate(Sum.Total.Conversion.Value = sum(as.numeric(Total.Conversion.Value), na.rm = T)) %>%
  dplyr::mutate(Avg.Link.Clicks = mean(as.numeric(Link.Clicks), na.rm = T)) %>%
  dplyr::mutate(Avg.Clicks.All = mean(as.numeric(Clicks.All), na.rm = T)) %>%
  dplyr::mutate(Avg.Impressions = mean(as.numeric(Impressions), na.rm = T)) %>%
  dplyr::mutate(Avg.Amount.Spent = mean(as.numeric(Amount.Spent), na.rm = T)) %>%
  dplyr::mutate(Avg.Website.Purchases = mean(as.numeric(Website.Purchases), na.rm = T)) %>%
  dplyr::mutate(Avg.Total.Conversion.Value = mean(as.numeric(Total.Conversion.Value), na.rm = T)) %>%
  dplyr::mutate(CPC = Sum.Amount.Spent / Sum.Link.Clicks) %>%
  dplyr::mutate(CTR = Sum.Link.Clicks / Sum.Impressions)  %>%
  dplyr::mutate(CVR = Sum.Website.Purchases / Sum.Link.Clicks)  %>%
  dplyr::ungroup() %>%
  dplyr::select(Age, Gender, Region, N,
                Sum.Link.Clicks, Sum.Clicks.All, Sum.Impressions, Sum.Amount.Spent, Sum.Website.Purchases, Sum.Total.Conversion.Value, 
                Avg.Link.Clicks, Avg.Clicks.All, Avg.Impressions, Avg.Amount.Spent, Avg.Website.Purchases, Avg.Total.Conversion.Value,
                CTR, CPC, CVR) %>%
  dplyr::distinct()

# Link.Clicks
M1 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Sum.Link.Clicks, Avg.Link.Clicks) %>%
  dplyr::arrange(desc(Sum.Link.Clicks)) %>%
  dplyr::mutate(Sum.rank = 1:n()) %>%
  dplyr::arrange(desc(Avg.Link.Clicks)) %>%
  dplyr::mutate(Avg.rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

## M1_1 <- theDataAvg[theDataAvg$Age == M1$Age[1] & theDataAvg$Gender == M1$Gender[1] & theDataAvg$Region == M1$Region[1], ]
## M1_1_rest <- theDataAvg[theDataAvg$Age != M1$Age[1] | theDataAvg$Gender != M1$Gender[1] | theDataAvg$Region != M1$Region[1], ]
## M1_2 <- theDataAvg[theDataAvg$Age == M1$Age[2] & theDataAvg$Gender == M1$Gender[2] & theDataAvg$Region == M1$Region[2], ]
## M1_2_rest <- theDataAvg[theDataAvg$Age != M1$Age[2] | theDataAvg$Gender != M1$Gender[2] | theDataAvg$Region != M1$Region[2], ]
## M1_3 <- theDataAvg[theDataAvg$Age == M1$Age[3] & theDataAvg$Gender == M1$Gender[3] & theDataAvg$Region == M1$Region[3], ]
## M1_3_rest <- theDataAvg[theDataAvg$Age != M1$Age[3] | theDataAvg$Gender != M1$Gender[3] | theDataAvg$Region != M1$Region[3], ]
## M1_4 <- theDataAvg[theDataAvg$Age == M1$Age[4] & theDataAvg$Gender == M1$Gender[4] & theDataAvg$Region == M1$Region[4], ]
## M1_4_rest <- theDataAvg[theDataAvg$Age != M1$Age[4] | theDataAvg$Gender != M1$Gender[4] | theDataAvg$Region != M1$Region[4], ]
## M1_5 <- theDataAvg[theDataAvg$Age == M1$Age[5] & theDataAvg$Gender == M1$Gender[5] & theDataAvg$Region == M1$Region[5], ]
## M1_5_rest <- theDataAvg[theDataAvg$Age != M1$Age[5] | theDataAvg$Gender != M1$Gender[5] | theDataAvg$Region != M1$Region[5], ]
## 
## M1$t.test[1]  <- t.test(M1_1$Link.Clicks, M1_1_rest$Link.Clicks)$statistic
## M1$p.value[1] <- t.test(M1_1$Link.Clicks, M1_1_rest$Link.Clicks)$p.value
## M1$t.test[2]  <- t.test(M1_2$Link.Clicks, M1_2_rest$Link.Clicks)$statistic
## M1$p.value[2] <- t.test(M1_2$Link.Clicks, M1_2_rest$Link.Clicks)$p.value
## M1$t.test[3]  <- t.test(M1_3$Link.Clicks, M1_3_rest$Link.Clicks)$statistic
## M1$p.value[3] <- t.test(M1_3$Link.Clicks, M1_3_rest$Link.Clicks)$p.value
## M1$t.test[4]  <- t.test(M1_4$Link.Clicks, M1_4_rest$Link.Clicks)$statistic
## M1$p.value[4] <- t.test(M1_4$Link.Clicks, M1_4_rest$Link.Clicks)$p.value
## M1$t.test[5]  <- t.test(M1_5$Link.Clicks, M1_5_rest$Link.Clicks)$statistic
## M1$p.value[5] <- t.test(M1_5$Link.Clicks, M1_5_rest$Link.Clicks)$p.value

## write.csv(M1, 'Output_Link.Clicks_2tailed.csv', row.names = F)

M1_1      <- theDataAvg[theDataAvg$Age == M1$Age[1] & theDataAvg$Gender == M1$Gender[1] & theDataAvg$Region == M1$Region[1], ]
M1_1_rest <- theDataAvg[theDataAvg$Age != M1$Age[1] | theDataAvg$Gender != M1$Gender[1] | theDataAvg$Region != M1$Region[1], ]
M1_2      <- M1_1_rest[M1_1_rest$Age == M1$Age[2] & M1_1_rest$Gender == M1$Gender[2] & M1_1_rest$Region == M1$Region[2], ]
M1_2_rest <- M1_1_rest[M1_1_rest$Age != M1$Age[2] | M1_1_rest$Gender != M1$Gender[2] | M1_1_rest$Region != M1$Region[2], ]
M1_3      <- M1_2_rest[M1_2_rest$Age == M1$Age[3] & M1_2_rest$Gender == M1$Gender[3] & M1_2_rest$Region == M1$Region[3], ]
M1_3_rest <- M1_2_rest[M1_2_rest$Age != M1$Age[3] | M1_2_rest$Gender != M1$Gender[3] | M1_2_rest$Region != M1$Region[3], ]
M1_4      <- M1_3_rest[M1_3_rest$Age == M1$Age[4] & M1_3_rest$Gender == M1$Gender[4] & M1_3_rest$Region == M1$Region[4], ]
M1_4_rest <- M1_3_rest[M1_3_rest$Age != M1$Age[4] | M1_3_rest$Gender != M1$Gender[4] | M1_3_rest$Region != M1$Region[4], ]
M1_5      <- M1_4_rest[M1_4_rest$Age == M1$Age[5] & M1_4_rest$Gender == M1$Gender[5] & M1_4_rest$Region == M1$Region[5], ]
M1_5_rest <- M1_4_rest[M1_4_rest$Age != M1$Age[5] | M1_4_rest$Gender != M1$Gender[5] | M1_4_rest$Region != M1$Region[5], ]

M1$t.test[1]  <- t.test(M1_1$Link.Clicks, M1_1_rest$Link.Clicks, alternative = 'greater')$statistic
M1$p.value[1] <- t.test(M1_1$Link.Clicks, M1_1_rest$Link.Clicks, alternative = 'greater')$p.value
M1$t.test[2]  <- t.test(M1_2$Link.Clicks, M1_2_rest$Link.Clicks, alternative = 'greater')$statistic
M1$p.value[2] <- t.test(M1_2$Link.Clicks, M1_2_rest$Link.Clicks, alternative = 'greater')$p.value
M1$t.test[3]  <- t.test(M1_3$Link.Clicks, M1_3_rest$Link.Clicks, alternative = 'greater')$statistic
M1$p.value[3] <- t.test(M1_3$Link.Clicks, M1_3_rest$Link.Clicks, alternative = 'greater')$p.value
M1$t.test[4]  <- t.test(M1_4$Link.Clicks, M1_4_rest$Link.Clicks, alternative = 'greater')$statistic
M1$p.value[4] <- t.test(M1_4$Link.Clicks, M1_4_rest$Link.Clicks, alternative = 'greater')$p.value
M1$t.test[5]  <- t.test(M1_5$Link.Clicks, M1_5_rest$Link.Clicks, alternative = 'greater')$statistic
M1$p.value[5] <- t.test(M1_5$Link.Clicks, M1_5_rest$Link.Clicks, alternative = 'greater')$p.value

write.csv(M1, 'Output_Link.Clicks.csv', row.names = F)

# CPC
M2 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Avg.Amount.Spent, Avg.Link.Clicks, CPC) %>%
  dplyr::filter(N >= 10) %>%
  dplyr::arrange(CPC) %>%
  dplyr::mutate(Rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M2_1      <- theDataAvg[theDataAvg$Age == M2$Age[1] & theDataAvg$Gender == M2$Gender[1] & theDataAvg$Region == M2$Region[1], ]
M2_1_rest <- theDataAvg[theDataAvg$Age != M2$Age[1] | theDataAvg$Gender != M2$Gender[1] | theDataAvg$Region != M2$Region[1], ]
M2_2      <- M2_1_rest[M2_1_rest$Age == M2$Age[2] & M2_1_rest$Gender == M2$Gender[2] & M2_1_rest$Region == M2$Region[2], ]
M2_2_rest <- M2_1_rest[M2_1_rest$Age != M2$Age[2] | M2_1_rest$Gender != M2$Gender[2] | M2_1_rest$Region != M2$Region[2], ]
M2_3      <- M2_2_rest[M2_2_rest$Age == M2$Age[3] & M2_2_rest$Gender == M2$Gender[3] & M2_2_rest$Region == M2$Region[3], ]
M2_3_rest <- M2_2_rest[M2_2_rest$Age != M2$Age[3] | M2_2_rest$Gender != M2$Gender[3] | M2_2_rest$Region != M2$Region[3], ]
M2_4      <- M2_3_rest[M2_3_rest$Age == M2$Age[4] & M2_3_rest$Gender == M2$Gender[4] & M2_3_rest$Region == M2$Region[4], ]
M2_4_rest <- M2_3_rest[M2_3_rest$Age != M2$Age[4] | M2_3_rest$Gender != M2$Gender[4] | M2_3_rest$Region != M2$Region[4], ]
M2_5      <- M2_4_rest[M2_4_rest$Age == M2$Age[5] & M2_4_rest$Gender == M2$Gender[5] & M2_4_rest$Region == M2$Region[5], ]
M2_5_rest <- M2_4_rest[M2_4_rest$Age != M2$Age[5] | M2_4_rest$Gender != M2$Gender[5] | M2_4_rest$Region != M2$Region[5], ]

M2$t.test[1]  <- t.test(M2_1$CPC, M2_1_rest$CPC, alternative = 'less')$statistic
M2$p.value[1] <- t.test(M2_1$CPC, M2_1_rest$CPC, alternative = 'less')$p.value
M2$t.test[2]  <- t.test(M2_2$CPC, M2_2_rest$CPC, alternative = 'less')$statistic
M2$p.value[2] <- t.test(M2_2$CPC, M2_2_rest$CPC, alternative = 'less')$p.value
M2$t.test[3]  <- t.test(M2_3$CPC, M2_3_rest$CPC, alternative = 'less')$statistic
M2$p.value[3] <- t.test(M2_3$CPC, M2_3_rest$CPC, alternative = 'less')$p.value
M2$t.test[4]  <- t.test(M2_4$CPC, M2_4_rest$CPC, alternative = 'less')$statistic
M2$p.value[4] <- t.test(M2_4$CPC, M2_4_rest$CPC, alternative = 'less')$p.value
M2$t.test[5]  <- t.test(M2_5$CPC, M2_5_rest$CPC, alternative = 'less')$statistic
M2$p.value[5] <- t.test(M2_5$CPC, M2_5_rest$CPC, alternative = 'less')$p.value

write.csv(M2, 'Output_CPC.csv', row.names = F)

# CTR
M3 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Avg.Link.Clicks, Avg.Impressions, CTR) %>%
  dplyr::filter(N >= 10) %>%
  dplyr::arrange(desc(CTR)) %>%
  dplyr::mutate(Rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M3_1      <- theDataAvg[theDataAvg$Age == M3$Age[1] & theDataAvg$Gender == M3$Gender[1] & theDataAvg$Region == M3$Region[1], ]
M3_1_rest <- theDataAvg[theDataAvg$Age != M3$Age[1] | theDataAvg$Gender != M3$Gender[1] | theDataAvg$Region != M3$Region[1], ]
M3_2      <- M3_1_rest[M3_1_rest$Age == M3$Age[2] & M3_1_rest$Gender == M3$Gender[2] & M3_1_rest$Region == M3$Region[2], ]
M3_2_rest <- M3_1_rest[M3_1_rest$Age != M3$Age[2] | M3_1_rest$Gender != M3$Gender[2] | M3_1_rest$Region != M3$Region[2], ]
M3_3      <- M3_2_rest[M3_2_rest$Age == M3$Age[3] & M3_2_rest$Gender == M3$Gender[3] & M3_2_rest$Region == M3$Region[3], ]
M3_3_rest <- M3_2_rest[M3_2_rest$Age != M3$Age[3] | M3_2_rest$Gender != M3$Gender[3] | M3_2_rest$Region != M3$Region[3], ]
M3_4      <- M3_3_rest[M3_3_rest$Age == M3$Age[4] & M3_3_rest$Gender == M3$Gender[4] & M3_3_rest$Region == M3$Region[4], ]
M3_4_rest <- M3_3_rest[M3_3_rest$Age != M3$Age[4] | M3_3_rest$Gender != M3$Gender[4] | M3_3_rest$Region != M3$Region[4], ]
M3_5      <- M3_4_rest[M3_4_rest$Age == M3$Age[5] & M3_4_rest$Gender == M3$Gender[5] & M3_4_rest$Region == M3$Region[5], ]
M3_5_rest <- M3_4_rest[M3_4_rest$Age != M3$Age[5] | M3_4_rest$Gender != M3$Gender[5] | M3_4_rest$Region != M3$Region[5], ]

M3$t.test[1]  <- t.test(M3_1$CTR, M3_1_rest$CTR, alternative = 'greater')$statistic
M3$p.value[1] <- t.test(M3_1$CTR, M3_1_rest$CTR, alternative = 'greater')$p.value
M3$t.test[2]  <- t.test(M3_2$CTR, M3_2_rest$CTR, alternative = 'greater')$statistic
M3$p.value[2] <- t.test(M3_2$CTR, M3_2_rest$CTR, alternative = 'greater')$p.value
M3$t.test[3]  <- t.test(M3_3$CTR, M3_3_rest$CTR, alternative = 'greater')$statistic
M3$p.value[3] <- t.test(M3_3$CTR, M3_3_rest$CTR, alternative = 'greater')$p.value
M3$t.test[4]  <- t.test(M3_4$CTR, M3_4_rest$CTR, alternative = 'greater')$statistic
M3$p.value[4] <- t.test(M3_4$CTR, M3_4_rest$CTR, alternative = 'greater')$p.value
M3$t.test[5]  <- t.test(M3_5$CTR, M3_5_rest$CTR, alternative = 'greater')$statistic
M3$p.value[5] <- t.test(M3_5$CTR, M3_5_rest$CTR, alternative = 'greater')$p.value

write.csv(M3, 'Output_CTR.csv', row.names = F)

# Website.Purchases
M4 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Sum.Website.Purchases, Avg.Website.Purchases) %>%
  dplyr::arrange(desc(Sum.Website.Purchases)) %>%
  dplyr::mutate(Sum.rank = 1:n()) %>%
  dplyr::arrange(desc(Avg.Website.Purchases)) %>%
  dplyr::mutate(Avg.rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M4_1      <- theDataAvg[theDataAvg$Age == M4$Age[1] & theDataAvg$Gender == M4$Gender[1] & theDataAvg$Region == M4$Region[1], ]
M4_1_rest <- theDataAvg[theDataAvg$Age != M4$Age[1] | theDataAvg$Gender != M4$Gender[1] | theDataAvg$Region != M4$Region[1], ]
M4_2      <- M4_1_rest[M4_1_rest$Age == M4$Age[2] & M4_1_rest$Gender == M4$Gender[2] & M4_1_rest$Region == M4$Region[2], ]
M4_2_rest <- M4_1_rest[M4_1_rest$Age != M4$Age[2] | M4_1_rest$Gender != M4$Gender[2] | M4_1_rest$Region != M4$Region[2], ]
M4_3      <- M4_2_rest[M4_2_rest$Age == M4$Age[3] & M4_2_rest$Gender == M4$Gender[3] & M4_2_rest$Region == M4$Region[3], ]
M4_3_rest <- M4_2_rest[M4_2_rest$Age != M4$Age[3] | M4_2_rest$Gender != M4$Gender[3] | M4_2_rest$Region != M4$Region[3], ]
M4_4      <- M4_3_rest[M4_3_rest$Age == M4$Age[4] & M4_3_rest$Gender == M4$Gender[4] & M4_3_rest$Region == M4$Region[4], ]
M4_4_rest <- M4_3_rest[M4_3_rest$Age != M4$Age[4] | M4_3_rest$Gender != M4$Gender[4] | M4_3_rest$Region != M4$Region[4], ]
M4_5      <- M4_4_rest[M4_4_rest$Age == M4$Age[5] & M4_4_rest$Gender == M4$Gender[5] & M4_4_rest$Region == M4$Region[5], ]
M4_5_rest <- M4_4_rest[M4_4_rest$Age != M4$Age[5] | M4_4_rest$Gender != M4$Gender[5] | M4_4_rest$Region != M4$Region[5], ]

M4$t.test[1]  <- t.test(M4_1$Website.Purchases, M4_1_rest$Website.Purchases, alternative = 'greater')$statistic
M4$p.value[1] <- t.test(M4_1$Website.Purchases, M4_1_rest$Website.Purchases, alternative = 'greater')$p.value
M4$t.test[2]  <- t.test(M4_2$Website.Purchases, M4_2_rest$Website.Purchases, alternative = 'greater')$statistic
M4$p.value[2] <- t.test(M4_2$Website.Purchases, M4_2_rest$Website.Purchases, alternative = 'greater')$p.value
M4$t.test[3]  <- t.test(M4_3$Website.Purchases, M4_3_rest$Website.Purchases, alternative = 'greater')$statistic
M4$p.value[3] <- t.test(M4_3$Website.Purchases, M4_3_rest$Website.Purchases, alternative = 'greater')$p.value
M4$t.test[4]  <- t.test(M4_4$Website.Purchases, M4_4_rest$Website.Purchases, alternative = 'greater')$statistic
M4$p.value[4] <- t.test(M4_4$Website.Purchases, M4_4_rest$Website.Purchases, alternative = 'greater')$p.value
M4$t.test[5]  <- t.test(M4_5$Website.Purchases, M4_5_rest$Website.Purchases, alternative = 'greater')$statistic
M4$p.value[5] <- t.test(M4_5$Website.Purchases, M4_5_rest$Website.Purchases, alternative = 'greater')$p.value

write.csv(M4, 'Output_Website.Purchases.csv', row.names = F)

# Total.Conversion.Value
M5 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Sum.Total.Conversion.Value, Avg.Total.Conversion.Value) %>%
  dplyr::arrange(desc(Sum.Total.Conversion.Value)) %>%
  dplyr::mutate(Sum.rank = 1:n()) %>%
  dplyr::arrange(desc(Avg.Total.Conversion.Value)) %>%
  dplyr::mutate(Avg.rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M5_1      <- theDataAvg[theDataAvg$Age == M5$Age[1] & theDataAvg$Gender == M5$Gender[1] & theDataAvg$Region == M5$Region[1], ]
M5_1_rest <- theDataAvg[theDataAvg$Age != M5$Age[1] | theDataAvg$Gender != M5$Gender[1] | theDataAvg$Region != M5$Region[1], ]
M5_2      <- M5_1_rest[M5_1_rest$Age == M5$Age[2] & M5_1_rest$Gender == M5$Gender[2] & M5_1_rest$Region == M5$Region[2], ]
M5_2_rest <- M5_1_rest[M5_1_rest$Age != M5$Age[2] | M5_1_rest$Gender != M5$Gender[2] | M5_1_rest$Region != M5$Region[2], ]
M5_3      <- M5_2_rest[M5_2_rest$Age == M5$Age[3] & M5_2_rest$Gender == M5$Gender[3] & M5_2_rest$Region == M5$Region[3], ]
M5_3_rest <- M5_2_rest[M5_2_rest$Age != M5$Age[3] | M5_2_rest$Gender != M5$Gender[3] | M5_2_rest$Region != M5$Region[3], ]
M5_4      <- M5_3_rest[M5_3_rest$Age == M5$Age[4] & M5_3_rest$Gender == M5$Gender[4] & M5_3_rest$Region == M5$Region[4], ]
M5_4_rest <- M5_3_rest[M5_3_rest$Age != M5$Age[4] | M5_3_rest$Gender != M5$Gender[4] | M5_3_rest$Region != M5$Region[4], ]
M5_5      <- M5_4_rest[M5_4_rest$Age == M5$Age[5] & M5_4_rest$Gender == M5$Gender[5] & M5_4_rest$Region == M5$Region[5], ]
M5_5_rest <- M5_4_rest[M5_4_rest$Age != M5$Age[5] | M5_4_rest$Gender != M5$Gender[5] | M5_4_rest$Region != M5$Region[5], ]

M5$t.test[1]  <- t.test(M5_1$Total.Conversion.Value, M5_1_rest$Total.Conversion.Value, alternative = 'greater')$statistic
M5$p.value[1] <- t.test(M5_1$Total.Conversion.Value, M5_1_rest$Total.Conversion.Value, alternative = 'greater')$p.value
M5$t.test[2]  <- t.test(M5_2$Total.Conversion.Value, M5_2_rest$Total.Conversion.Value, alternative = 'greater')$statistic
M5$p.value[2] <- t.test(M5_2$Total.Conversion.Value, M5_2_rest$Total.Conversion.Value, alternative = 'greater')$p.value
M5$t.test[3]  <- t.test(M5_3$Total.Conversion.Value, M5_3_rest$Total.Conversion.Value, alternative = 'greater')$statistic
M5$p.value[3] <- t.test(M5_3$Total.Conversion.Value, M5_3_rest$Total.Conversion.Value, alternative = 'greater')$p.value
M5$t.test[4]  <- t.test(M5_4$Total.Conversion.Value, M5_4_rest$Total.Conversion.Value, alternative = 'greater')$statistic
M5$p.value[4] <- t.test(M5_4$Total.Conversion.Value, M5_4_rest$Total.Conversion.Value, alternative = 'greater')$p.value
M5$t.test[5]  <- t.test(M5_5$Total.Conversion.Value, M5_5_rest$Total.Conversion.Value, alternative = 'greater')$statistic
M5$p.value[5] <- t.test(M5_5$Total.Conversion.Value, M5_5_rest$Total.Conversion.Value, alternative = 'greater')$p.value

write.csv(M5, 'Output_Total.Conversion.Value.csv', row.names = F)

# Clicks.All
M6 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Sum.Clicks.All, Avg.Clicks.All) %>%
  dplyr::arrange(desc(Sum.Clicks.All)) %>%
  dplyr::mutate(Sum.rank = 1:n()) %>%
  dplyr::arrange(desc(Avg.Clicks.All)) %>%
  dplyr::mutate(Avg.rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M6_1      <- theDataAvg[theDataAvg$Age == M6$Age[1] & theDataAvg$Gender == M6$Gender[1] & theDataAvg$Region == M6$Region[1], ]
M6_1_rest <- theDataAvg[theDataAvg$Age != M6$Age[1] | theDataAvg$Gender != M6$Gender[1] | theDataAvg$Region != M6$Region[1], ]
M6_2      <- M6_1_rest[M6_1_rest$Age == M6$Age[2] & M6_1_rest$Gender == M6$Gender[2] & M6_1_rest$Region == M6$Region[2], ]
M6_2_rest <- M6_1_rest[M6_1_rest$Age != M6$Age[2] | M6_1_rest$Gender != M6$Gender[2] | M6_1_rest$Region != M6$Region[2], ]
M6_3      <- M6_2_rest[M6_2_rest$Age == M6$Age[3] & M6_2_rest$Gender == M6$Gender[3] & M6_2_rest$Region == M6$Region[3], ]
M6_3_rest <- M6_2_rest[M6_2_rest$Age != M6$Age[3] | M6_2_rest$Gender != M6$Gender[3] | M6_2_rest$Region != M6$Region[3], ]
M6_4      <- M6_3_rest[M6_3_rest$Age == M6$Age[4] & M6_3_rest$Gender == M6$Gender[4] & M6_3_rest$Region == M6$Region[4], ]
M6_4_rest <- M6_3_rest[M6_3_rest$Age != M6$Age[4] | M6_3_rest$Gender != M6$Gender[4] | M6_3_rest$Region != M6$Region[4], ]
M6_5      <- M6_4_rest[M6_4_rest$Age == M6$Age[5] & M6_4_rest$Gender == M6$Gender[5] & M6_4_rest$Region == M6$Region[5], ]
M6_5_rest <- M6_4_rest[M6_4_rest$Age != M6$Age[5] | M6_4_rest$Gender != M6$Gender[5] | M6_4_rest$Region != M6$Region[5], ]

M6$t.test[1]  <- t.test(M6_1$Clicks.All, M6_1_rest$Clicks.All, alternative = 'greater')$statistic
M6$p.value[1] <- t.test(M6_1$Clicks.All, M6_1_rest$Clicks.All, alternative = 'greater')$p.value
M6$t.test[2]  <- t.test(M6_2$Clicks.All, M6_2_rest$Clicks.All, alternative = 'greater')$statistic
M6$p.value[2] <- t.test(M6_2$Clicks.All, M6_2_rest$Clicks.All, alternative = 'greater')$p.value
M6$t.test[3]  <- t.test(M6_3$Clicks.All, M6_3_rest$Clicks.All, alternative = 'greater')$statistic
M6$p.value[3] <- t.test(M6_3$Clicks.All, M6_3_rest$Clicks.All, alternative = 'greater')$p.value
M6$t.test[4]  <- t.test(M6_4$Clicks.All, M6_4_rest$Clicks.All, alternative = 'greater')$statistic
M6$p.value[4] <- t.test(M6_4$Clicks.All, M6_4_rest$Clicks.All, alternative = 'greater')$p.value
M6$t.test[5]  <- t.test(M6_5$Clicks.All, M6_5_rest$Clicks.All, alternative = 'greater')$statistic
M6$p.value[5] <- t.test(M6_5$Clicks.All, M6_5_rest$Clicks.All, alternative = 'greater')$p.value

write.csv(M6, 'Output_Clicks.All.csv', row.names = F)

# CVR
M7 <- theDataSum %>%
  dplyr::select(Age, Gender, Region, N, Avg.Website.Purchases, Avg.Link.Clicks, CVR) %>%
  dplyr::filter(N >= 10) %>%
  dplyr::arrange(desc(CVR)) %>%
  dplyr::mutate(Rank = 1:n()) %>%
  dplyr::mutate(t.test = NA) %>%
  dplyr::mutate(p.value = NA)

M7_1      <- theDataAvg[theDataAvg$Age == M7$Age[1] & theDataAvg$Gender == M7$Gender[1] & theDataAvg$Region == M7$Region[1], ]
M7_1_rest <- theDataAvg[theDataAvg$Age != M7$Age[1] | theDataAvg$Gender != M7$Gender[1] | theDataAvg$Region != M7$Region[1], ]
M7_2      <- M7_1_rest[M7_1_rest$Age == M7$Age[2] & M7_1_rest$Gender == M7$Gender[2] & M7_1_rest$Region == M7$Region[2], ]
M7_2_rest <- M7_1_rest[M7_1_rest$Age != M7$Age[2] | M7_1_rest$Gender != M7$Gender[2] | M7_1_rest$Region != M7$Region[2], ]
M7_3      <- M7_2_rest[M7_2_rest$Age == M7$Age[3] & M7_2_rest$Gender == M7$Gender[3] & M7_2_rest$Region == M7$Region[3], ]
M7_3_rest <- M7_2_rest[M7_2_rest$Age != M7$Age[3] | M7_2_rest$Gender != M7$Gender[3] | M7_2_rest$Region != M7$Region[3], ]
M7_4      <- M7_3_rest[M7_3_rest$Age == M7$Age[4] & M7_3_rest$Gender == M7$Gender[4] & M7_3_rest$Region == M7$Region[4], ]
M7_4_rest <- M7_3_rest[M7_3_rest$Age != M7$Age[4] | M7_3_rest$Gender != M7$Gender[4] | M7_3_rest$Region != M7$Region[4], ]
M7_5      <- M7_4_rest[M7_4_rest$Age == M7$Age[5] & M7_4_rest$Gender == M7$Gender[5] & M7_4_rest$Region == M7$Region[5], ]
M7_5_rest <- M7_4_rest[M7_4_rest$Age != M7$Age[5] | M7_4_rest$Gender != M7$Gender[5] | M7_4_rest$Region != M7$Region[5], ]

M7$t.test[1]  <- t.test(M7_1$CVR, M7_1_rest$CVR, alternative = 'greater')$statistic
M7$p.value[1] <- t.test(M7_1$CVR, M7_1_rest$CVR, alternative = 'greater')$p.value
M7$t.test[2]  <- t.test(M7_2$CVR, M7_2_rest$CVR, alternative = 'greater')$statistic
M7$p.value[2] <- t.test(M7_2$CVR, M7_2_rest$CVR, alternative = 'greater')$p.value
M7$t.test[3]  <- t.test(M7_3$CVR, M7_3_rest$CVR, alternative = 'greater')$statistic
M7$p.value[3] <- t.test(M7_3$CVR, M7_3_rest$CVR, alternative = 'greater')$p.value
M7$t.test[4]  <- t.test(M7_4$CVR, M7_4_rest$CVR, alternative = 'greater')$statistic
M7$p.value[4] <- t.test(M7_4$CVR, M7_4_rest$CVR, alternative = 'greater')$p.value
M7$t.test[5]  <- t.test(M7_5$CVR, M7_5_rest$CVR, alternative = 'greater')$statistic
M7$p.value[5] <- t.test(M7_5$CVR, M7_5_rest$CVR, alternative = 'greater')$p.value

write.csv(M7, 'Output_CVR.csv', row.names = F)
  
# save summary
wb <- xlsx::createWorkbook()
wsh <- xlsx::createSheet(wb, sheetName = 'Summary')

rows <- xlsx::createRow(wsh, rowIndex = 1)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Cheapest to reach (CPC)')
xlsx::addDataFrame(as.data.frame(head(M2, 5)), wsh, startRow = 2, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 8)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Most likely to click (CTR)')
xlsx::addDataFrame(as.data.frame(head(M3, 5)), wsh, startRow = 9, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 15)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Most likely to buy (CVR)')
xlsx::addDataFrame(as.data.frame(head(M7, 5)), wsh, startRow = 16, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 22)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Highest sales value (€)')
xlsx::addDataFrame(as.data.frame(head(M5, 5)), wsh, startRow = 23, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 29)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Most All clicks')
xlsx::addDataFrame(as.data.frame(head(M6, 5)), wsh, startRow = 30, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 36)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Most Website clicks')
xlsx::addDataFrame(as.data.frame(head(M1, 5)), wsh, startRow = 37, startColumn = 1, row.names = F)

rows <- xlsx::createRow(wsh, rowIndex = 43)
cell <- xlsx::createCell(rows, colIndex = 1)
xlsx::setCellValue(cell[[1, 1]], 'Most Purchases')
xlsx::addDataFrame(as.data.frame(head(M4, 5)), wsh, startRow = 44, startColumn = 1, row.names = F)

xlsx::saveWorkbook(wb, "Output_SUMMARY.xlsx")

