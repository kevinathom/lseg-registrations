# Lippincott Library LSEG Workspace Registration Audit Automation
This tool automates portions of the daily LSEG Workspace registrations audit workflow.

## Context and process
In certain academic implementations of LSEG Workspace, patrons can self-register for access. These registrations succeed inconsistently. In some cases, the correct licenses get assigned to the appropriate individuals, requiring no administrator support. In others, problems occur (e.g.: alumni receive access they should not have; accounts get created without licenses for accessing subscribed content) that require intervention. LSEG provides daily files listing registrants, facilitating audits by administrators within the subscribing library.

The Lippincott Library of the Wharton School developed informal criteria and personal workflows for auditing and correcting new LSEG Workspace accounts. These workflows require detailed knowledge of what should happen for different patron categories, what problems can occur, and how to fix each problem. They also require detailed attention to identify issue such as duplicate registrations and conflicts with details presented in different systems.

This tool attempts to reduce the need for specialized knowledge and detailed attention. It does this in two stages:
1. It accepts [an ongoing "Accounts to Check" file](/assets/AccountstoCheck-LSEG_NewCols.xlsx) plus one day's registration files from LSEG, de-duplicates records and labels potential issues, and returns an updated "Accounts to Check" file for a human to review and augment. Processing steps:
	1. Flag new records as new
	1. Date new records
	1. Remove new identical records
	1. Extract and store the email local-part
	1. Flag alumni email addresses
	1. Flag repeated names
	1. Flag repeated email local-parts
	1. Append new records to the "Accounts to Check" file
1. It uses the human-augmented "Accounts to Check" file, labels records with recommended next actions, and returns an updated "Accounts to Check" file for a human to review and implement. Processing steps:
	1. Identify records that require follow-up from a prior action
	1. Identify records that should have their patron-type label fixed
	1. Identify records that need licenses removed
	1. Identify records that need an account created
	1. Identify records that have duplicate accounts and need one or more removed
	1. Identify records that need licenses assigned
	1. Assign a unique identifier to records without an identical
	1. Remove flags for new records

While the process does not integrate with proprietary systems to automate user lookups or make account changes directly, it does make the daily registrations audit less taxing for librarians to perform at scale, makes the work feasible with less procedural knowledge, and leaves humans in the loop to verify every meaningful action.

## Notes about files
- The ["Accounts to Check" file](/assets/AccountstoCheck-LSEG_NewCols.xlsx) available here offers a template that works with the scripted process. Notes and color-coding identify how different fields are intended to work with this tool. It is meant to be used iteratively, loaded and appended to each day.
- The Python script [process_gui.py](/code/process_gui.py) holds the tool's latest code. It probably should be a later version of [process.py](/code/process.py), but the separate files avoided having to use the repository's history as a reference while implementing the tool's GUI.
- The [proposed process flow diagram](/assets/flowchart_LSEG-2025-12_proposed.drawio.png) reflects a preliminary target process. The tool's final state may resemble this workflow broadly but deviates in some details.

I compile an executable via this line in a terminal: `python -m PyInstaller --onefile -w 'control.py'`