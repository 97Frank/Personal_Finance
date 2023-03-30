#Functions used in script----
load <- function(){
  list.of.packages <- c("tidyverse",
                        "haven",
                        "expss",
                        "readxl", 
                        "writexl",
                        "officer",
                        "rvg",
                        "ggrepel", 
                        "systemfonts", 
                        "rvest", 
                        "lubridate",
                        "gridExtra",
                        "flextable",
                        "qualtRics",
                        "randomForest")
  
  
  
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  
  
  if(length(new.packages)) {
    install.packages(new.packages)
  }
  
  
  lapply(list.of.packages, require, character.only = TRUE)
  
  
}
# Define the spread_transaction function
spread_transaction <- function(transactions, search_text, n_months, transaction_date = NULL) {
  
  # Find the transaction to spread
  if (is.null(transaction_date)) {
    transaction_to_spread <- transactions %>%
      filter(str_detect(Text, search_text)) %>%
      slice(1) # In case there are multiple matches, take the first one
  } else {
    # Convert transaction_date to Date object if it's not
    if (!inherits(transaction_date, "Date")) {
      transaction_date <- as.Date(transaction_date)
    }
    
    transaction_to_spread <- transactions %>%
      filter(str_detect(Text, search_text) & Transaktionsdatum == transaction_date) %>%
      slice(1) # In case there are still multiple matches, take the first one
  }
  
  # Calculate the monthly amount
  monthly_amount <- transaction_to_spread$Belopp / n_months
  
  # Create the new transactions
  new_transactions <- tibble(
    Transaktionsdatum = seq.Date(
      transaction_to_spread$Transaktionsdatum,
      by = "month",
      length.out = n_months
    ),
    Text = rep(transaction_to_spread$Text, n_months),
    Belopp = rep(monthly_amount, n_months),
    Match = rep(transaction_to_spread$Match, n_months),
    Main = rep(transaction_to_spread$Main, n_months),
    Sub = rep(transaction_to_spread$Sub, n_months),

  )
  
  # Remove the original transaction
  if (is.null(transaction_date)) {
    transactions <- transactions %>%
      filter(!str_detect(Text, search_text))
  } else {
    transactions <- transactions %>%
      filter(!(str_detect(Text, search_text) & Transaktionsdatum == transaction_date))
  }
  
  # Add the new transactions
  transactions <- bind_rows(transactions, new_transactions)
  
  return(transactions)
}
# Define the kvitta_transactions function
kvitta_transactions <- function(transactions, cost_text, cost_date, income_text, income_date) {
  # Convert the dates to Date objects if they're not
  if (!inherits(cost_date, "Date")) {
    cost_date <- as.Date(cost_date)
  }
  if (!inherits(income_date, "Date")) {
    income_date <- as.Date(income_date)
  }
  
  # Find the cost and income transactions
  cost_transaction <- transactions %>% 
    filter(Text == cost_text & Transaktionsdatum == cost_date) %>%
    slice(1) # In case there are multiple matches, take the first one
  
  income_transaction <- transactions %>% 
    filter(Text == income_text & Transaktionsdatum == income_date) %>%
    slice(1) # In case there are multiple matches, take the first one
  
  # Subtract the income Belopp from the cost Belopp
  org_cost_belopp <- cost_transaction$Belopp
  new_cost_belopp <- cost_transaction$Belopp - income_transaction$Belopp
  
  # Update the cost transaction's Belopp and kr columns
  transactions <- transactions %>%
    mutate(Belopp = case_when(
      Text == cost_text & Transaktionsdatum == cost_date ~ new_cost_belopp,
      Text == income_text & Transaktionsdatum == income_date ~ Belopp - (org_cost_belopp - new_cost_belopp),
      TRUE ~ Belopp))
  
  
  return(transactions)
}

load()
#MAIN SCRIPT HB----
#HB transactions
transaktioner <- read_excel("HB_Uttag_0330.xlsx", skip = 4) %>% 
  mutate(Text = str_remove_all(string = Text, pattern = "[^[:alnum:] ]"),
         Transaktionsdatum = lubridate::as_date(Transaktionsdatum),
         Belopp = as.numeric(gsub("-","",gsub(",",".",gsub(" ","",Belopp))))) %>% 
  select(c(3, 5, 7))

#category excel/google sheet
data <- read_excel("Kategorier_0330.xlsx") %>% 
  mutate(Text = str_remove_all(string = Text, pattern = "[^[:alnum:] ]"))




#for loop to map categories to transactions

transaktioner$Match = "NO MATCH"
transaktioner$Main = "NO MATCH"
transaktioner$Sub = "NO MATCH"

for (i in 1:nrow(transaktioner)) {
  for (z in 1:nrow(data)) {
    TF <- str_detect(string = transaktioner$Text[i], pattern =  data$Text[z])
    if (TF == TRUE) {
      transaktioner$Match[i] = data$Text[z]
      transaktioner$Main[i] = data$Main[z]
      transaktioner$Sub[i] = data$Sub[z]
    }
  }
}


#Filter the data from excluded transactions

filtered <- transaktioner %>% 
  filter(Main == "Filter") 
HB <- transaktioner %>% 
  filter(Main != "Filter") %>% 
  mutate(Main = case_when(
    str_detect(Text,"Rocker AB") & Transaktionsdatum == "2023-02-28" ~ "Nöje & shopping",
    TRUE ~ Main),
    Sub = case_when(
    str_detect(Text,"Rocker AB") & Transaktionsdatum == "2023-02-28" ~ "Hemelektronik",
    TRUE ~ Sub),
    Konto = "HANDELSBANKEN")


# Use the spread function to put amount to multiple periods
HB <- spread_transaction(HB, "Rocker AB", 12, "2023-02-28")
HB <- spread_transaction(HB, "Datacamp", 12, "2023-03-09")
HB <- spread_transaction(HB, "CENTRALA STUDI", 12, "2023-02-20")
HB <- spread_transaction(HB, "Pixel 7", 12, "2023-01-26")



#MAIN SCRIPT AMEX----
amex <- read_excel("amex_0330.xlsx",skip=6) %>% 
  select(1:3) %>% 
  mutate(Beskrivning = toupper(str_remove_all(string = Beskrivning, pattern = "[^[:alnum:] ]")),
         Datum = lubridate::mdy(Datum),
         Konto = "AMERICAN EXPRESS")
colnames(amex)[1:2] <- c("Transaktionsdatum","Text")

data <- read_excel("Kategorier_0330.xlsx") %>% 
  mutate(Text = toupper(str_remove_all(string = Text, pattern = "[^[:alnum:] ]")))



amex$Match = "NO MATCH"
amex$Main = "NO MATCH"
amex$Sub = "NO MATCH"

for (i in 1:nrow(amex)) {
  for (z in 1:nrow(data)) {
    TF <- str_detect(string = amex$Text[i], pattern =  data$Text[z])
    if (TF == TRUE) {
      amex$Match[i] = data$Text[z]
      amex$Main[i] = data$Main[z]
      amex$Sub[i] = data$Sub[z]
    }
  }
}



filtered_amex <- amex %>% 
  filter(Main == "Filter") 
no_match_amex <- amex %>% 
  filter(Main == "NO MATCH") 
#COMBINE AND WRITE FINAL DATAFRAMES----
DF <- bind_rows(HB,amex)
