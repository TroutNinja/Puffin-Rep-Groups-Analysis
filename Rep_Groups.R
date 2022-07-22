setwd("/Users/matt/Desktop/Puffin/Projects/Rep_Group_Project")
rm(list = ls())


library(tidyverse)
library(tidyverse)
library(tidyxl)
library(dplyr)
library(pdftools)
library(lubridate)
library(readxl)
library(writexl)
library(glue)
library(arsenal)

# import data rep group data that is currently in Business Central
rep_groups <- read_csv("Reps_Groups.csv")

###################################
# Outdoor Channel Rep Group Data
###################################

# import dataframe of accurate Reps / Groups & Customers for Outdoor
reps <- readxl::read_xlsx("Outdoor/Updating Rep Group_Salesperson.xlsx", col_names = FALSE)
colnames(reps) <- c("id", "name", "correct_group", "correct_rep")

pattern <- c("JG2HI", "ALGER SALES", "HOUSE ACCOUNTS", "GIB CARSON", "DIVERSE MARKETING",
                 "PORTICO", "JG2HI", "RITZ SISTERS", "PORTICO COLLECTION", "TWIST")

reps <- reps %>% filter(!correct_group %in% pattern)

##############################
# Gift Channel Rep Group Data
##############################

# Import Ritz Data and Clean / Transform
ritz <- readxl::read_xls("Gift/Ritz rep list .xls")
ritz <- ritz %>% rename(name = `Shipping Name`,
                        correct_rep = `Current Sales Rep`) %>%
  mutate(correct_group = "RITZ SISTERS",
         correct_rep = str_to_upper(correct_rep),
         id = NA) %>% 
  select(id, name, correct_group, correct_rep)


# Import Twist Data and Clean / Transform
twist <- readxl::read_xls("Gift/Twist Rep Group .xls")
twist <- twist %>% rename(name = `Customer Name`,
                          correct_rep = Salesperson) %>% 
  mutate(correct_group = "TWIST",
         correct_rep = str_to_upper(correct_rep),
         id = NA) %>% 
  select(id, name, correct_group, correct_rep)


# Import JG2HI Data and Clean / Transform
JG2HI <- readxl::read_xls("Gift/JG2HI rep list.xls", sheet = "Report")
JG2HI <- JG2HI %>% rename(name = `Billing Name`,
                          correct_rep = `Current Sales Rep`) %>% 
  mutate(correct_group = "JG2HI",
         correct_rep = str_to_upper(correct_rep),
         id = NA) %>% 
  select(id, name, correct_group, correct_rep)


# Bind Rows for Outdoor and Gift Records
reps <- do.call(rbind, list(reps, ritz, twist, JG2HI))


#######################
# Merge Data Frames
#######################

combined <- left_join(rep_groups, reps, by = c("Name" = "name"), keep = TRUE)

combined <- combined %>% 
  mutate(group_matches_system = if_else(`Rep Group` == correct_group, TRUE, FALSE)) %>% 
  mutate(rep_matches_system = if_else(Salesperson_Code == correct_rep, TRUE, FALSE)) %>% 
  select(No, id, Name, name, `Rep Group`, correct_group, Salesperson_Code, correct_rep, group_matches_system, rep_matches_system)

combined <- combined %>% rename(BC_Customer_No = No,
                    Correct_Customer_No = id,
                    BC_Customer_Name = Name,
                    Correct_Customer_Name = name,
                    BC_Rep_Group = `Rep Group`,
                    Correct_Rep_Group = correct_group,
                    BC_Salesperson = Salesperson_Code,
                    Correct_Rep = correct_rep,
                    Group_Matches_BC = group_matches_system,
                    Rep_Matches_BC = rep_matches_system
                    )

combined_false <- combined %>% filter(Group_Matches_BC == FALSE)

combined %>% filter(BC_Customer_Name == "Duluth Trading Co.")

writexl::write_xlsx(combined, "Rep_Group_Analysis.xlsx")



