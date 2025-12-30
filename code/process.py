# Libraries
import os
import pandas as pd
import re

# Open files
os.chdir(f"C:/Users/kevinat/Documents/GitHub/lseg-registrations/res/")
dat_today_fname = f"ProductRegistrationSummaryRequest_20251202.csv"
dat_ongoing_fname = f"AccountstoCheck-LSEG.xlsx"

dat_today = pd.read_csv(dat_today_fname)
dat_ongoing = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)

# Remove full duplicate records from today's file
dat_today.drop(dat_today[dat_today.duplicated(keep = 'first')].index, inplace=True)
dat_today.reset_index(inplace=True, drop=True)

dat_today.drop(dat_today[pd.merge(dat_today, dat_ongoing, on=list(dat_today.columns), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, inplace=True)
dat_today.reset_index(inplace=True, drop=True)

# Flag potential duplicates
dat_today["Flag_Duplicate"] = ""

## Name
dat_today.loc[dat_today[dat_today.loc[:, ["FIRST NAME", "LAST NAME"]].duplicated(keep = 'first')].index, "Flag_Duplicate"] = dat_today.loc[dat_today[dat_today.loc[:, ["FIRST NAME", "LAST NAME"]].duplicated(keep = 'first')].index, "Flag_Duplicate"] + "Duplicated new name on " + re.search('\\_(.*)\\.', dat_today_fname).group(1) + ". "
dat_today.loc[dat_today[pd.merge(dat_today, dat_ongoing, on=list(["FIRST NAME", "LAST NAME"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, "Flag_Duplicate"] = dat_today.loc[dat_today[pd.merge(dat_today, dat_ongoing, on=list(["FIRST NAME", "LAST NAME"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, "Flag_Duplicate"] + "Duplicated old name on " + re.search('\\_(.*)\\.', dat_today_fname).group(1) + ". "

## Email
dat_today["Email_Prefix"] = dat_today["COMPANY EMAIL"].str.extract(r'(.+?(?=\@))')
dat_today.loc[dat_today[dat_today.loc[:, ["Email_Prefix"]].duplicated(keep = 'first')].index, "Flag_Duplicate"] = dat_today.loc[dat_today[dat_today.loc[:, ["Email_Prefix"]].duplicated(keep = 'first')].index, "Flag_Duplicate"] + "Duplicated new email prefix on " + re.search('\\_(.*)\\.', dat_today_fname).group(1) + ". "
dat_today.loc[dat_today[pd.merge(dat_today, dat_ongoing, on=list(["Email_Prefix"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, "Flag_Duplicate"] = dat_today.loc[dat_today[pd.merge(dat_today, dat_ongoing, on=list(["Email_Prefix"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, "Flag_Duplicate"] + "Duplicated old email prefix on " + re.search('\\_(.*)\\.', dat_today_fname).group(1) + ". "

# Merge data to ongoing file
dat_today.sort_values(by=["LAST NAME", "FIRST NAME", "COMPANY EMAIL"], inplace=True, ignore_index=True)
dat_ongoing = pd.concat([dat_ongoing, dat_today], sort=False, ignore_index=True).fillna("")
dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="ProductRegistrationSummaryReque", index=False)
