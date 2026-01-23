# Libraries
import os
import pandas as pd
import re
from datetime import datetime
from datetime import timedelta
from itertools import compress

# Prompt user for files

# Access files
os.chdir(f"C:/Users/kevinat/Documents/GitHub/lseg-registrations/res/")

dat_today_fname = f"TEST-ProductRegistrationSummaryRequest_20251202.csv"
dat_ongoing_fname = f"TEST-AccountstoCheck-LSEG.xlsx"
dat_ongoing_fname2 = f"TEST-AccountstoCheck-LSEG_part2.xlsx" #for testing stage 2 only
#OR
dat_today_fname = f"ProductRegistrationSummaryRequest_20251202.csv"
dat_ongoing_fname = f"AccountstoCheck-LSEG.xlsx"

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
#following line is commented out for testing; swap with the next line for production
#dat_ongoing = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)
dat_ongoing = pd.read_excel(dat_ongoing_fname2, na_values=[], keep_default_na=False)

# Assess next actions
mask_iterate = list(dat_ongoing.index)

## Send follow-up
mask = dat_ongoing[pd.to_datetime(dat_ongoing.loc[:, "Followup_DueDate"], format='YYYY-MM-DD') <= datetime.today()].index
dat_ongoing.loc[mask, "Email_Text"] = "Optional placeholder text for multiple accounts email template."
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Send follow-up email. "
dat_ongoing.loc[mask, "Followup_DueDate"] = ""

mask = dat_ongoing[dat_ongoing.loc[:, "Followup_DueDate"] == ""].index
mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

## Fix patron label
mask = dat_ongoing[dat_ongoing.loc[:, "New_Record"] != ""].index
mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

for i in mask_iterate:
  if dat_ongoing.loc[i, "Graduation Year"] == "ALUM" and dat_ongoing.loc[i, "LABEL"] != "Alumni":
    dat_ongoing.loc[i, "LABEL"] = "Alumni"
    dat_ongoing.loc[i, "Take_Action"] = dat_ongoing.loc[i, "Take_Action"] + "Change LSEG Label to Alumni. "
  elif dat_ongoing.loc[i, "Graduation Year"] == "N/A" and bool(re.search(r"Staff", dat_ongoing.loc[i, "Notes"])) and dat_ongoing.loc[i, "LABEL"] != "Staff":
    dat_ongoing.loc[i, "LABEL"] = "Staff"
    dat_ongoing.loc[i, "Take_Action"] = dat_ongoing.loc[i, "Take_Action"] + "Change LSEG Label to Staff. "
  elif (isinstance(dat_ongoing.loc[i, "Graduation Year"], int) or dat_ongoing.loc[i, "Graduation Year"] == "N/A") and (bool(re.search(r"PhD", dat_ongoing.loc[i, "Notes"])) or bool(re.search(r"Faculty", dat_ongoing.loc[i, "Notes"]))) and dat_ongoing.loc[i, "LABEL"] != "Faculty/PhD":
    dat_ongoing.loc[i, "LABEL"] = "Faculty/PhD"
    dat_ongoing.loc[i, "Take_Action"] = dat_ongoing.loc[i, "Take_Action"] + "Change LSEG Label to Faculty/PhD. "
  elif (isinstance(dat_ongoing.loc[i, "Graduation Year"], int) or dat_ongoing.loc[i, "Graduation Year"] == "Unknown") and not bool(re.search(r"PhD", dat_ongoing.loc[i, "Notes"])) and dat_ongoing.loc[i, "LABEL"] != "Student":
    dat_ongoing.loc[i, "LABEL"] = "Student"
    dat_ongoing.loc[i, "Take_Action"] = dat_ongoing.loc[i, "Take_Action"] + "Change LSEG Label to Student. "

## Remove licenses
mask = dat_ongoing[dat_ongoing.loc[:, "Has Licenses in backend"] == "Yes"].index
mask2 = dat_ongoing[dat_ongoing.loc[:, "LABEL"] == "Alumni"].index
mask = list(compress(mask2, [i in set(mask) for i in mask2]))
mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Unassign all licenses. "
dat_ongoing.loc[mask, "New_Record"] = ""

mask = dat_ongoing[dat_ongoing.loc[:, "LABEL"] != "Alumni"].index
mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

## Create account
mask = sorted(list(set(list(dat_ongoing[pd.to_numeric(dat_ongoing["Appears in backend"], errors="coerce") <= 0].index) + list(dat_ongoing[dat_ongoing.loc[:, "Appears in backend"] == "No"].index))))
mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
dat_ongoing.loc[mask, "Email_Text"] = "Placeholder text: Your account is ready."
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Create an LSEG account and notify by email. "
dat_ongoing.loc[mask, "New_Record"] = ""

mask = dat_ongoing[dat_ongoing.loc[:, "New_Record"] != ""].index
mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

## De-duplicate accounts
mask = dat_ongoing[pd.to_numeric(dat_ongoing["Appears in backend"], errors="coerce") > 1].index
mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
dat_ongoing.loc[mask, "Followup_DueDate"] = "DUE DATE TBD" ####should be today + N days -- see next line for attempted logic
pd.Series([pd.to_datetime((datetime.today() + timedelta(days = 7)), format='YYYY-MM-DD')]).dt.date
dat_ongoing.loc[mask, "Email_Text"] = "Placeholder text: You created accounts under {dat_ongoing.loc[mask, 'COMPANY EMAIL']} and __. Which would you prefer to keep?"
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Patron has multiple LSEG accounts; ask which they'd perfer to keep. "
dat_ongoing.loc[mask, "New_Record"] = ""

## Add licenses
mask = dat_ongoing[dat_ongoing.loc[:, "Has Licenses in backend"] == "No"].index
mask2 = dat_ongoing[dat_ongoing.loc[:, "LABEL"] != "Alumni"].index
mask = list(compress(mask2, [i in set(mask) for i in mask2]))
mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Assign licenses. "

## No further action
dat_ongoing.loc[:, "New_Record"] = ""

# Merge data to ongoing file
dat_ongoing.sort_values(by=["Take_Action"], inplace=True, ignore_index=True)
dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="in", index=False)
