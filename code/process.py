# Libraries
import os
import pandas as pd
import re
from datetime import datetime
from itertools import compress

# Prompt user for files

# Access files
os.chdir(f"C:/Users/kevinat/Documents/GitHub/lseg-registrations/res/")
dat_today_fname = f"TEST-ProductRegistrationSummaryRequest_20251202.csv"
dat_ongoing_fname = f"TEST-AccountstoCheck-LSEG.xlsx"

dat_today = pd.read_csv(dat_today_fname)
dat_ongoing = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)

# Remove full duplicate records from today's file
dat_today.drop(dat_today[dat_today.duplicated(keep = 'first')].index, inplace=True)
dat_today.reset_index(inplace=True, drop=True)

dat_today.drop(dat_today[pd.merge(dat_today, dat_ongoing, on=list(dat_today.columns), how='left', indicator=True).loc[:, '_merge'] == 'both'].index, inplace=True)
dat_today.reset_index(inplace=True, drop=True)

# Flag potential issues
dat_today["Email_local-part"] = dat_today["COMPANY EMAIL"].str.extract(r'(.+?(?=\@))')
dat_today["New_Warning"] = ""
file_date = re.search('\\_(.*)\\.', dat_today_fname).group(1)

## Alumni email
mask = dat_today[dat_today["COMPANY EMAIL"].str.contains(r"(^[a-zA-Z]+\.[a-zA-Z]+\.w[a-zA-Z]\d\d@[a-zA-Z]*\.*upenn\.edu$)", na=False)].index
dat_today.loc[mask, "New_Warning"] = dat_today.loc[mask, "New_Warning"] + f"Alumni-style email. "

## Duplicate name
mask = dat_today[dat_today.loc[:, ["FIRST NAME", "LAST NAME"]].duplicated(keep = False)].index
dat_today.loc[mask, "New_Warning"] = dat_today.loc[mask, "New_Warning"] + f"Repeated name in {file_date} file. "

mask = dat_today[pd.merge(dat_today, dat_ongoing.drop_duplicates(subset=["FIRST NAME", "LAST NAME"]), on=list(["FIRST NAME", "LAST NAME"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index
dat_today.loc[mask, "New_Warning"] = dat_today.loc[mask, "New_Warning"] + f"Name from {file_date} exists in old file. "

## Duplicate email
mask = dat_today[dat_today.loc[:, ["Email_local-part"]].duplicated(keep = False)].index
dat_today.loc[mask, "New_Warning"] = dat_today.loc[mask, "New_Warning"] + f"Repeated email prefix in {file_date} file. "

mask = dat_today[pd.merge(dat_today, dat_ongoing.drop_duplicates(subset=["Email_local-part"]), on=list(["Email_local-part"]), how='left', indicator=True).loc[:, '_merge'] == 'both'].index
dat_today.loc[mask, "New_Warning"] = dat_today.loc[mask, "New_Warning"] + f"Email prefix from {file_date} exists in old file. "

# Mark today's records as new
dat_today["New_Record"] = file_date

# Merge data to ongoing file
dat_today.sort_values(by=["LAST NAME", "FIRST NAME", "COMPANY EMAIL"], inplace=True, ignore_index=True)
dat_ongoing = pd.concat([dat_ongoing, dat_today], sort=False, ignore_index=True).fillna("")
dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="in", index=False)
# Attempt to append to file, but it overwrites anyway
#dat_today.to_excel(dat_ongoing_fname, sheet_name="in", startrow=len(dat_ongoing.index)+1, columns=dat_today.columns.tolist(), index=False, header=False)

# Prompt user to look up accounts and fill fields for next step

# Re-access file
dat_ongoing = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)

# Assess next actions
mask_iterate = list(dat_ongoing.index)

## Follow-up due
mask = dat_ongoing[dat_ongoing.loc[:, "Followup_DueDate"] == ""].index
mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

mask = dat_ongoing[pd.to_datetime(dat_ongoing.loc[:, "Followup_DueDate"], format='YYYY-MM-DD') <= datetime.today()].index
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Follow up. "
dat_ongoing.loc[mask, "Followup_DueDate"] = ""

## Alumni account


# Merge data to ongoing file
dat_ongoing["New_Record"] = ""
dat_ongoing.sort_values(by=["Take_Action", "LAST NAME", "FIRST NAME", "COMPANY EMAIL"], inplace=True, ignore_index=True)
dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="in", index=False)
