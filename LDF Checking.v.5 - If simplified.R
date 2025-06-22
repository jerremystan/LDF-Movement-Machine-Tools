library(dplyr)
library(readr)
library(writexl)
library(readxl)
library(tidyverse)

ldf_df <- read_excel("Input/LDF Movement 25Q1 v.1.00.xlsx", sheet = "R dataset")
inc_df <- read_excel("Input/Incurred for R.v.2.xlsx", sheet = "data")

Quarter_Prev <- "2024Q4"
Quarter_Now <- "2025Q1"

significant_level <- 0.05
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
#wide_inc$`Diff Gross` <- wide_inc$`Gross Incurred_2024Q4` - wide_inc$`Gross Incurred_2024Q3`
wide_inc$`Diff Gross` <- wide_inc[[paste0("Gross Incurred_",Quarter_Now)]] - wide_inc[[paste0("Gross Incurred_",Quarter_Prev)]]
#wide_inc$`Diff Net` <- wide_inc$`Net Incurred_2024Q4` - wide_inc$`Net Incurred_2024Q3`
wide_inc$`Diff Net` <- wide_inc[[paste0("Net Incurred_",Quarter_Now)]] - wide_inc[[paste0("Net Incurred_",Quarter_Prev)]]
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
          ,!!paste0("Gross Incurred_", Quarter_Prev) := .data[[paste0("Gross Incurred_", Quarter_Prev)]][which.max(abs(`Diff Net`))]
          ,!!paste0("Gross Incurred_", Quarter_Now) := .data[[paste0("Gross Incurred_", Quarter_Now)]][which.max(abs(`Diff Net`))]
          ,!!paste0("Net Incurred_", Quarter_Prev) := .data[[paste0("Net Incurred_", Quarter_Prev)]][which.max(abs(`Diff Net`))]
          ,!!paste0("Net Incurred_", Quarter_Now) := .data[[paste0("Net Incurred_", Quarter_Now)]][which.max(abs(`Diff Net`))]
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

  #NEW REGISTERED
  if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf[[paste0("Net Incurred_", Quarter_Prev)]][i] < 1){
    arranged_ldf$Description[i] <- paste0("New Registered of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  #Lambang Diff dengan Total sama
  
  else if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf[[paste0("Net Incurred_", Quarter_Prev)]][i] > 1 & arranged_ldf$`Total Diff Net`[i] > 0){
    arranged_ldf$Description[i] <- paste0("Increased Incurred of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  } 
  else if(arranged_ldf$`Diff Net`[i] < 0 & arranged_ldf[[paste0("Net Incurred_", Quarter_Prev)]][i] > 1 & arranged_ldf$`Total Diff Net`[i] < 0){
    arranged_ldf$Description[i] <- paste0("Decreased Incurred of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  #Lambang Diff dengan Total beda
  
  else if(arranged_ldf$`Diff Net`[i] > 0 & arranged_ldf[[paste0("Net Incurred_", Quarter_Prev)]][i] > 1 & arranged_ldf$`Total Diff Net`[i] < 0){
    arranged_ldf$Description[i] <- paste0("Offset by Increased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  } 
  else if(arranged_ldf$`Diff Net`[i] < 0 & arranged_ldf[[paste0("Net Incurred_", Quarter_Prev)]][i] > 1 & arranged_ldf$`Total Diff Net`[i] > 0){
    arranged_ldf$Description[i] <- paste0("Offset by Decreased of ", arranged_ldf$`Client Name`[i], " by IDR ", arranged_ldf$`Diff Gross Label`[i], " / IDR ", arranged_ldf$`Diff Net Label`[i], " (Gross/Net)")
  }
  
  else{
    arranged_ldf$Description[i] <- NA
  }
  
  
}

#Move Offset to last Rank

max_rank <- 10

arranged_2 <- arranged_ldf

for (i in 1:max_rank){
  temp_rank <- max_rank
  for (j in 1:nrow(arranged_2)){
    
    temp_rank <- temp_rank + 1
    min_rank <- min(arranged_2$Rank[arranged_2$Field == arranged_2$Field[j] & arranged_2$LoB == arranged_2$LoB[j] & arranged_2$`Loss - Quarter (YYYYQn)` == arranged_2$`Loss - Quarter (YYYYQn)`[j]], na.rm = TRUE)
    
    if(!is.na(min_rank) && !is.na(arranged_2$Rank[j]) && 
       arranged_2$Rank[j] == min_rank && 
       !is.na(arranged_2$Description[j]) && startsWith(arranged_2$Description[j], "Offset"))
    {
      
      arranged_2$Rank[j] <- temp_rank
      
    }
    else{
      
    }
    
  }
  
}

arranged_2 <- arranged_2 %>%
  arrange(Field, LoB, `Loss - Quarter (YYYYQn)`, Rank)

result_ldf <- arranged_2 %>%
  select(Field, LoB,`Loss - Quarter (YYYYQn)`, Description)

library(openxlsx)

wb <- createWorkbook()

addWorksheet(wb, "Result")
addWorksheet(wb, "Detail")

writeData(wb, sheet = "Result", result_ldf)
writeData(wb, sheet = "Detail", arranged_2)

saveWorkbook(wb, paste0("LDF Movement Result Significant Description Simplified Code - ",Quarter_Now," - ",significant_level,".xlsx"), overwrite = TRUE)
