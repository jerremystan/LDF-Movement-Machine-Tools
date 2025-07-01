library(dplyr)
library(readr)
library(writexl)
library(readxl)
library(tidyverse)

ldf_df <- read_excel("LDF Movement 24Q4 v.1.00.xlsx", sheet = "R dataset")
inc_df <- read_excel("Incurred for R.v.2.xlsx", sheet = "data")

Quarter_Prev <- "2024Q3"
Quarter_Now <- "2024Q4"

significant_level <- 0.3
#Significant level artinya berapa nilai selisih Net Incurred terhadap Total Incurred per Accident Period tersebut

group_inc <- inc_df %>%
  group_by(`LoB`,`Client ID`,`Client Name`,`Loss - Quarter (YYYYQn)`,`RPT - Quarter (YYYYQn)`) %>%
  summarise(`Gross Incurred` = sum(`Gross Incurred`)
            ,`Net Incurred` = sum(`Net Incurred`))


wide_inc <- pivot_wider(
  group_inc,
  names_from = `RPT - Quarter (YYYYQn)`,
  values_from = c("Gross Incurred","Net Incurred")
)
wide_inc[is.na(wide_inc)] <- 0
wide_inc$`Diff Gross` <- wide_inc$`Gross Incurred_2024Q4` - wide_inc$`Gross Incurred_2024Q3`
wide_inc$`Diff Net` <- wide_inc$`Net Incurred_2024Q4` - wide_inc$`Net Incurred_2024Q3`
wide_inc[is.na(wide_inc)] <- 0
wide_inc$`Abs Diff Gross` <- abs(wide_inc$`Diff Gross`)
wide_inc$`Abs Diff Net` <- abs(wide_inc$`Diff Net`)
wide_inc[is.na(wide_inc)] <- 0

cols_to_summarise <- names(wide_inc)[5:10]  # Dynamically extract column names

wide_inc_no_client <- wide_inc %>%
  group_by(`LoB`, `Loss - Quarter (YYYYQn)`) %>%
  summarise(`Total Diff Gross` = sum(`Diff Gross`)
            ,`Total Diff Net` = sum(`Diff Net`)
            ,`Abs Total Diff Gross` = abs(sum(`Diff Gross`))
            ,`Abs Total Diff Net` = abs(sum(`Diff Net`))
            )

#summarise(across(all_of(cols_to_summarise), first, .names = "{.col}"))

wide_inc_sig <- wide_inc %>%
  left_join(wide_inc_no_client, by = c("LoB", "Loss - Quarter (YYYYQn)")) %>%
  mutate(
    `Net Significance` = `Abs Diff Net`/`Abs Total Diff Net`
  )

wide_inc_max <- wide_inc_sig %>%
  group_by(`LoB`, `Loss - Quarter (YYYYQn)`) %>%
  reframe(
          ,`Client ID` = `Client ID`[which.max(abs(`Diff Net`))]
          ,`Client Name` = `Client Name`[which.max(abs(`Diff Net`))]
          ,`Gross Incurred_2024Q3` = `Gross Incurred_2024Q3`[which.max(abs(`Diff Net`))]
          ,`Gross Incurred_2024Q4` = `Gross Incurred_2024Q4`[which.max(abs(`Diff Net`))]
          ,`Net Incurred_2024Q3` = `Net Incurred_2024Q3`[which.max(abs(`Diff Net`))]
          ,`Net Incurred_2024Q4` = `Net Incurred_2024Q4`[which.max(abs(`Diff Net`))]
          ,`Diff Gross` = `Diff Gross`[which.max(abs(`Diff Net`))]
          ,`Diff Net` = `Diff Net`[which.max(abs(`Diff Net`))]
          ,`Abs Diff Gross` = abs(`Diff Gross`[which.max(abs(`Diff Net`))])
          ,`Abs Diff Net` = abs(`Diff Net`[which.max(abs(`Diff Net`))])
  ) %>%
  left_join(wide_inc_no_client, by = c("LoB", "Loss - Quarter (YYYYQn)")) %>%
  mutate(
    `Net Significance` = `Abs Diff Net`/`Abs Total Diff Net`
  ) %>%
  select(
    colnames(wide_inc_sig)
  )

wide_inc_sig <- wide_inc_sig[wide_inc_sig$`Net Significance` > significant_level,]

new_ldf_1 <- left_join(ldf_df,wide_inc_sig, by = c("LoB","Loss - Quarter (YYYYQn)"))

new_ldf_2 <- new_ldf_1 %>%
  filter(is.na(`Net Significance`))  %>%
  select(
    Field
    ,LoB
    ,`Loss - Quarter (YYYYQn)`
  ) %>%
  left_join(wide_inc_max, by = c("LoB","Loss - Quarter (YYYYQn)"))

new_ldf_1 <- new_ldf_1[!is.na(new_ldf_1$`Net Significance`),]

desc_ldf <- rbind(new_ldf_1, new_ldf_2)

convert_to_label <- function(x) {
  if (x >= 1e9) {
    paste0(round(x / 1e9, 2), " Bio")  # Billions
  } else if (x >= 1e6) {
    paste0(round(x / 1e6, 2), " Mio")  # Millions
  } else if (x >= 1e3) {
    paste0(round(x / 1e3, 2), "k")     # Thousands
  } else {
    as.character(round(x, 2))          # Less than 1k, round and leave as-is
  }
}

desc_ldf$`Diff Gross Label` <- sapply(desc_ldf$`Abs Diff Gross`, convert_to_label)
desc_ldf$`Diff Net Label` <- sapply(desc_ldf$`Abs Diff Net`, convert_to_label)

arranged_ldf <- desc_ldf %>%
  arrange(desc(Field), LoB, desc(`Loss - Quarter (YYYYQn)`), desc(`Net Significance`))

arranged_ldf <- arranged_ldf %>%
  group_by(Field, LoB, `Loss - Quarter (YYYYQn)`) %>%
  mutate(Rank = rank(-`Net Significance`))

arranged_ldf$Description <- NA

for (i in 1:nrow(arranged_ldf)){

  min_rank <- min(arranged_ldf$Rank[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i]])
  
if(i == 1 | arranged_ldf$Rank[i] == min_rank){ #Baris pertama atau rank terkecil dalam setiap group
  
  #NEW REGISTERED
  if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf$`Net Incurred_2024Q3`[i] < 1){
    arranged_ldf$Description[i] <- paste0("New Registered of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  #Lambang Diff dengan Total sama
  
  else if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf$`Net Incurred_2024Q3`[i] > 1 & arranged_ldf$`Total Diff Net`[i] > 0){
    arranged_ldf$Description[i] <- paste0("Increased Incurred of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  } 
  else if(arranged_ldf$`Diff Net`[i] < 0 & arranged_ldf$`Net Incurred_2024Q3`[i] > 1 & arranged_ldf$`Total Diff Net`[i] < 0){
    arranged_ldf$Description[i] <- paste0("Decreased Incurred of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  #Lambang Diff dengan Total beda
  
  else if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf$`Net Incurred_2024Q3`[i] > 1 & arranged_ldf$`Total Diff Net`[i] < 0){
    arranged_ldf$Description[i] <- paste0("Offset by Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  } 
  else if(arranged_ldf$`Diff Net`[i] < 0 & arranged_ldf$`Net Incurred_2024Q3`[i] > 1 & arranged_ldf$`Total Diff Net`[i] > 0){
    arranged_ldf$Description[i] <- paste0("Offset by Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  else{
    arranged_ldf$Description[i] <- NA
  }
  
}else{
  
  if(#New Registered
    arranged_ldf$`Diff Net`[i] > 0 
     & arranged_ldf$`Net Incurred_2024Q3`[i] < 1 
    ){
    arranged_ldf$Description[i] <- paste0("New Registered of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
  
  else if(#Increased Incurred Inline Total
    arranged_ldf$`Diff Net`[i] > 0 
    & arranged_ldf$`Net Incurred_2024Q3`[i] > 1
    & arranged_ldf$`Total Diff Net`[i] > 0
  ){
    if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 0 
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Offset by Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0 
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 1
    ){
      arranged_ldf$Description[i] <- paste0("Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
  }
  
  else if(#Increased Incurred Not Inline Total
    arranged_ldf$`Diff Net`[i] > 0 
    & arranged_ldf$`Net Incurred_2024Q3`[i] > 1
    & arranged_ldf$`Total Diff Net`[i] < 0
  ){
    if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 0 
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Offset by Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0 
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 1
    ){
      arranged_ldf$Description[i] <- paste0("Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
  }
  
  else if(#Decreased Incurred Inline Total
    arranged_ldf$`Diff Net`[i] < 0 
    & arranged_ldf$`Net Incurred_2024Q3`[i] > 1
    & arranged_ldf$`Total Diff Net` < 0
  ){
    if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Offset by Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 1
    ){
      arranged_ldf$Description[i] <- paste0("Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
  }
  
  else if(#Decreased Incurred Not Inline Total
    arranged_ldf$`Diff Net`[i] < 0 
    & arranged_ldf$`Net Incurred_2024Q3`[i] > 1
    & arranged_ldf$`Total Diff Net` > 0
  ){
    if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 1
    ){
      arranged_ldf$Description[i] <- paste0("Offset by Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
    else if(
      arranged_ldf$`Diff Net`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] > 0
      & arranged_ldf$`Net Incurred_2024Q3`[arranged_ldf$Field == arranged_ldf$Field[i] & arranged_ldf$LoB == arranged_ldf$LoB[i] & arranged_ldf$`Loss - Quarter (YYYYQn)` == arranged_ldf$`Loss - Quarter (YYYYQn)`[i] & arranged_ldf$Rank == min_rank] < 1
    ){
      arranged_ldf$Description[i] <- paste0("Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
    }
  }
  
  else{
    arranged_ldf$Description[i] <- NA
  }
  
}
  
}


result_ldf <- arranged_ldf %>%
  select(Field, LoB,`Loss - Quarter (YYYYQn)`, Description)

library(openxlsx)

wb <- createWorkbook()

addWorksheet(wb, "Result")
addWorksheet(wb, "Detail")

writeData(wb, sheet = "Result", result_ldf)
writeData(wb, sheet = "Detail", arranged_ldf)

saveWorkbook(wb, paste0("LDF Movement Result Significant Description - ",Quarter_Now," - ",significant_level,".xlsx"), overwrite = TRUE)
