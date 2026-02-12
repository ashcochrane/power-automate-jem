# IT Setup Guide — Mobility Parking Pass Automation

This guide covers how to import and configure the Power Automate flow in your Microsoft 365 tenant.

## Prerequisites

- Microsoft 365 Business licence (or above) with:
  - Power Automate
  - OneDrive for Business or SharePoint Online
  - Excel Online (Business)
  - Word Online (Business)
  - Outlook
- An account with permissions to import Power Automate flows
- An account (or shared mailbox) to send the emails from

## Step 1 — Create the folder structure

In OneDrive for Business (or a SharePoint document library), create:

```
/Automation/
├── Input/
├── Templates/
└── Output/
    ├── PDF/
    └── DOCX/
```

The `PDF` folder stores the final passes. The `DOCX` folder is used temporarily during processing (the flow creates a Word file, converts it to PDF, then deletes the Word file automatically).

## Step 2 — Upload the templates

Copy these files from the `templates/` folder in this bundle:

| File | Upload to |
|---|---|
| `LetterTemplate.docx` | `/Automation/Templates/` |
| `Data.xlsx` | `/Automation/Input/` |

The Word template must be a `.docx` file with Plain Text Content Controls. See `docs/flow-build-guide.md` for details on the content control tags.

The Excel file must contain a named table called `tblLetters`. See `docs/flow-build-guide.md` for the full column list.

## Step 3 — Import the flow

There are two import methods. Try Method A first.

### Method A — Import Package (preferred)

1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **My flows** in the left sidebar
3. Click **Import** > **Import Package (Legacy)**
4. Upload `solution/MobilityParkingPass.zip` from this bundle
5. Power Automate will show the package contents and ask you to configure connections
6. For each of the 4 connections, click **Select during import** and choose an existing connection or create a new one:

| Connection | What it connects to |
|---|---|
| Excel Online (Business) | The account that has access to `Data.xlsx` |
| Word Online (Business) | The account that has access to `LetterTemplate.docx` |
| OneDrive for Business | The storage where `/Automation/` lives |
| Outlook | The account (or shared mailbox) that sends the emails |

7. Click **Import**

### Method B — Code View (backup)

If the package import fails:

1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **Create** > **Instant cloud flow**
3. Name it `Mobility Parking Pass`, select **Manually trigger a flow**, click **Create**
4. In the flow editor, click the **Code View** toggle (top right, `</>` icon)
5. Open `solution/flow-definition.json`
6. Copy the contents of the `properties.definition` object (just the definition, not the whole file)
7. Paste it into the code view, replacing the existing JSON
8. Click **Save**
9. Switch back to the visual designer

For full details on both methods, see `solution/IMPORT-INSTRUCTIONS.md`.

## Step 4 — Configure the flow

After import, open the flow in the visual editor. Several actions have placeholder values that need to point to your environment.

### 4.1 — List rows present in a table
- **Location:** Your OneDrive for Business or SharePoint site
- **File:** Browse to `/Automation/Input/Data.xlsx`
- **Table:** Select `tblLetters`

### 4.2 — Populate a Microsoft Word template (inside Try scope)
- **Location:** Your OneDrive for Business or SharePoint site
- **File:** Browse to `/Automation/Templates/LetterTemplate.docx`
- The 6 template fields should auto-populate if the content control tags match:
  `FullName`, `PermitNumber`, `VehiclePlate`, `VehicleType`, `EventName`, `EventDate`

### 4.3 — Create DOCX file (inside Try scope)
- **Folder path:** `/Automation/Output/DOCX`

### 4.4 — Convert Word Document to PDF (inside Try scope)
- **Location:** Should auto-detect from the created DOCX file
- If not, set to the same OneDrive/SharePoint location as step 4.1

### 4.5 — Create PDF file (inside Try scope)
- **Folder path:** `/Automation/Output/PDF`

### 4.6 — Send an email (inside Try scope)
- **From:** Confirm the sending account or shared mailbox
- Review the email body and update the contact details if needed

### 4.7 — Update a row — Success (inside Try scope)
- **Location, File, Table:** Same as step 4.1

### 4.8 — Update a row — Error (inside Catch scope)
- **Location, File, Table:** Same as step 4.1

### Save the flow after configuring all actions.

## Step 5 — Test with sample data

1. Open `Data.xlsx` in `/Automation/Input/`
2. Paste the 3 rows from `samples/SampleData.csv` into the `tblLetters` table
3. **Change the `Email` column** in all 3 rows to your own email address
4. Go to Power Automate and click **Run** on the flow
5. Confirm all of the following:

| Check | What to look for |
|---|---|
| Flow completed | Run history shows "Succeeded" |
| PDF files created | 3 PDFs in `/Automation/Output/PDF/` |
| PDF content correct | Open each PDF — all 6 fields filled in |
| Filenames correct | e.g. `LT38000001 - Smith.pdf` |
| Emails received | 3 emails in your inbox |
| Email subject correct | Shows event name and date |
| PDF attached | Attachment opens correctly |
| Excel updated | All 3 rows show `SendStatus = Sent` |
| Timestamps present | `SentAt` column has values |

6. Test error handling: add a row with a blank email, run again, confirm it shows `SendStatus = Error`
7. Test idempotency: run again with no new rows — confirm nothing is re-sent
8. Delete the test rows when done

## Step 6 — Hand over to the operator

Share the [End User Guide](end-user-guide.md) with whoever will run this day-to-day. This is a one-page guide covering: paste CSV data, check for duplicates, click Run, check results.

## Troubleshooting

| Problem | Check |
|---|---|
| Package import fails | Try Method B (Code View). See `solution/IMPORT-INSTRUCTIONS.md` |
| Flow can't find Excel file | Verify the file path in the List rows action and that the connection account has access |
| Word template fields are blank | Verify content control tags match exactly: `FullName`, `PermitNumber`, `VehiclePlate`, `VehicleType`, `EventName`, `EventDate` |
| PDF conversion fails | Ensure the Word Online (Business) connector is available in your tenant. Try OneDrive > Convert file as alternative |
| Email not sending | Check the Outlook connection and that the sending account has a mailbox |
| Row not updating to Sent | Verify `tblLetters` has `Permit No.` as a key column and the Update Row action references it |
| DOCX files building up | Check the Delete DOCX file action is enabled — it should clean up after each successful run |
| Duplicate emails sent | Ensure the filter query `SendStatus ne 'Sent'` is set on the List rows action |
