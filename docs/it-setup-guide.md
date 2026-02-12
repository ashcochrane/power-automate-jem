# IT Setup Guide — Mobility Parking Pass Automation

This guide covers how to import and configure the Power Automate solution in your Microsoft 365 tenant.

## Prerequisites

- Microsoft 365 Business licence (or above) with:
  - Power Automate
  - OneDrive for Business or SharePoint Online
  - Excel Online (Business)
  - Word Online (Business)
  - Outlook
- An account with permissions to import Power Automate solutions
- An account (or shared mailbox) to send the emails from

## Step 1 — Create the folder structure

In OneDrive for Business (or a SharePoint document library), create:

```
/Automation/
├── Input/
├── Templates/
├── Output/
│   └── PDF/
└── Logs/          (optional)
```

## Step 2 — Upload the templates

Copy these files from the `templates/` folder in this bundle:

| File | Upload to |
|---|---|
| `LetterTemplate.docx` | `/Automation/Templates/` |
| `Data.xlsx` | `/Automation/Input/` |

## Step 3 — Import the Power Automate solution

1. Go to [Power Automate](https://make.powerautomate.com)
2. In the left sidebar, click **Solutions**
3. Click **Import solution** in the top toolbar
4. Upload `solution/MobilityParkingPass.zip` from this bundle
5. Follow the prompts to create or select connections:

| Connection | What it connects to |
|---|---|
| Excel Online (Business) | The account that has access to `Data.xlsx` |
| Word Online (Business) | The account that has access to `LetterTemplate.docx` |
| OneDrive for Business | The storage where `/Automation/` lives |
| Outlook | The account (or shared mailbox) that sends the emails |

6. Complete the import

## Step 4 — Configure the flow

After import, open the flow and verify these settings:

### List rows action
- **Location:** Your OneDrive for Business or SharePoint site
- **File:** `/Automation/Input/Data.xlsx`
- **Table:** `tblLetters`

### Populate Word template action
- **Location:** Your OneDrive for Business or SharePoint site
- **File:** `/Automation/Templates/LetterTemplate.docx`

### Create file (PDF) action
- **Folder path:** `/Automation/Output/PDF`

### Send email action
- **From:** Confirm the sending account or shared mailbox
- Update the contact email/phone in the email body if needed

### Save the flow

## Step 5 — Test with sample data

1. Open `Data.xlsx` in `/Automation/Input/`
2. Paste the 3 rows from `samples/SampleData.csv` into the `tblLetters` table
3. Update the `Email` column to your own test email addresses
4. Go to Power Automate and run the flow manually
5. Confirm:
   - You receive the emails
   - The PDF passes look correct
   - The Excel rows update to `SendStatus = Sent`
6. Delete the test rows when done

## Step 6 — Hand over to the operator

Share the [End User Guide](end-user-guide.md) with whoever will run this day-to-day.

## Troubleshooting

| Problem | Check |
|---|---|
| Flow fails to find Excel file | Verify the file path and that the connection account has access |
| Word template fields are blank | Verify content control tags match exactly: `FullName`, `PermitNumber`, `VehiclePlate`, `VehicleType`, `EventName`, `EventDate` |
| PDF conversion fails | Ensure Word Online (Business) connector is available in your tenant |
| Email not sending | Check the Outlook connection and that the sending account has a mailbox |
| Row not updating | Verify `tblLetters` has a key column and the Update Row action is configured to use it |
