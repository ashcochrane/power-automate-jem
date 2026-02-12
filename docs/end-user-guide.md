# End User Guide — Mobility Parking Pass Automation

This guide explains how to send mobility parking passes for an Eden Park event.

## What you need

- The CSV file exported from the Eden Park website (the ballot entries)
- Access to the `Data.xlsx` file in OneDrive/SharePoint

## Step 1 — Open the Excel file

Open `Data.xlsx` from the `/Automation/Input/` folder in OneDrive (or SharePoint).

You will see a table called `tblLetters` with column headers.

## Step 2 — Paste the CSV data

1. Open the CSV file from the website export
2. Select all the data rows (not the header row)
3. Copy them
4. In `Data.xlsx`, click the first empty cell under the headers
5. Paste

Make sure the columns line up. The key columns are:

| Column | What it is |
|---|---|
| Name (First) | Recipient's first name |
| Name (Last) | Recipient's last name |
| Email (Enter Email) | Where the pass gets sent |
| Mobility or Disability Parking Permit Number | Their permit number |
| Vehicle number plate | Their plate number |
| Vehicle type | Standard vehicle or Van |
| Event (columns 8 and 9) | Event name and date |

## Step 3 — Check for duplicates

Before sending, check for duplicate entries. The spreadsheet has conditional formatting already set up — duplicates will be highlighted automatically.

**If duplicates are highlighted (coloured cells):**

1. Look at the **Mobility or Disability Parking Permit Number** column — duplicates are highlighted in red
2. Look at the **Vehicle number plate** column — duplicates are highlighted in yellow
3. For each set of duplicates, keep the most recent entry (latest Entry Date) and delete the older rows

**If conditional formatting is not set up yet:**

1. Select the **Mobility or Disability Parking Permit Number** column
2. Go to **Home > Conditional Formatting > Highlight Cell Rules > Duplicate Values**
3. Choose red highlight, click OK
4. Repeat for the **Vehicle number plate** column with yellow highlight
5. Delete the older duplicate rows, keeping the most recent entry for each

## Step 4 — Run the flow

1. Go to [Power Automate](https://make.powerautomate.com)
2. Find the flow called **Mobility Parking Pass**
3. Click **Run**
4. Wait for it to complete

## Step 5 — Check the results

Go back to `Data.xlsx` and check the `SendStatus` column:

| Status | Meaning |
|---|---|
| **Sent** | Pass was emailed successfully |
| **Error** | Something went wrong — check `ErrorMessage` column |
| *(blank)* | Not yet processed |

If any rows show **Error**, fix the issue (usually a missing or invalid email address) and run the flow again. It will only process rows that are not already marked as Sent.

## Tips

- Always check for duplicates before running
- Start with a small batch (5-10 rows) for a new event to make sure everything looks right
- The flow only sends to rows where `SendStatus` is blank — previously sent rows are skipped
- If you need to re-send to someone, clear their `SendStatus` cell and run again
