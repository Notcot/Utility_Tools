# For extracting cable patching info from a Master Patching List
- In purpose of data center patching info management
- To aid cable migrations

## REQUIRE:
- python
- openpyxl

## Description:
Given theres two excel sheets, one being a master and the other being the one on the job, and a primary key is used for both sheets. This tool matches the keys and extact info from master when the keys exist in both sheets. Producing a seperate excel sheet with only the relevant data.

## Beware
Tool will crash if any of the excel sheets contain too many empty rows and columns. Avoid exceeding 50000 rows and columns.
