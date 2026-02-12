# Mobility Parking Pass Automation — Design

## Context

Eden Park runs a mobility parking ballot for each event. People with disability parking permits apply through the website. Successful applicants receive a personalized parking pass as a PDF via email.

Currently this is manual. This project automates the flow: CSV data in, personalized PDF passes out, emailed to each recipient.

## Deliverable

A Power Automate solution package + supporting files, delivered to the client's IT team for import into their M365 tenant. The end user is non-technical — their only job is pasting data into Excel and clicking Run.

## Architecture

```
OneDrive for Business / SharePoint
└── /Automation/
    ├── /Input/
    │   └── Data.xlsx           (tblLetters table)
    ├── /Templates/
    │   └── LetterTemplate.docx (Word template with content controls)
    ├── /Output/
    │   └── /PDF/               (generated passes)
    └── /Logs/                  (optional)
```

## Data

### Source

CSV export from the Eden Park website form. Relevant columns:

| CSV Column | Used As | Example |
|---|---|---|
| Name (First) | FirstName | Magda |
| Name (Last) | LastName | Clayton |
| Email (Enter Email) | Email | bclayton@xtra.co.nz |
| Mobility or Disability Parking Permit Number | PermitNumber | LT38445958 |
| Vehicle number plate | VehiclePlate | NNN501 |
| Vehicle type | VehicleType | Standard vehicle |
| Event (col 8) | EventName | Blues v Chiefs |
| Event (col 9) | EventDate | Saturday 14 February |

### Ignored columns

Phone, Consent fields, Created By, Entry Date, Source Url, Transaction Id, Payment fields, User Agent, User IP, Submission Speed — form metadata not needed for the pass.

### Automation tracking columns (added to Excel)

| Column | Purpose |
|---|---|
| SendStatus | Blank / Sent / Error |
| SentAt | Timestamp when emailed |
| ErrorMessage | Captured on failure |

### Deduplication

Handled manually in Excel before running the flow. Conditional formatting highlights rows with duplicate mobility permit numbers or duplicate vehicle plates. The operator resolves duplicates before clicking Run.

## Word Template

The current pass was designed in Adobe InDesign and saved as RTF. It must be recreated as a `.docx` with Plain Text Content Controls for Power Automate.

### Layout

Single page, folds in half:
- Top half (back when folded): Terms and conditions
- Bottom half (front): Eden Park branding, event info, person's details

### Content controls (6 fields)

| Tag | Source |
|---|---|
| FullName | FirstName + " " + LastName |
| PermitNumber | Mobility or Disability Parking Permit Number |
| VehiclePlate | Vehicle number plate |
| VehicleType | Vehicle type |
| EventName | Event name |
| EventDate | Event date |

All other content (logo, T&Cs, layout) is static.

## Power Automate Flow

### Trigger

Instant cloud flow — manual button press.

### Steps

1. **List rows** from `tblLetters` in `Data.xlsx`, filtered: `SendStatus ne 'Sent'`
2. **Apply to each** row:

| Step | Action | Detail |
|---|---|---|
| 2a | Compose safe filename | `PermitNumber - LastName`, sanitized |
| 2b | Populate Word template | Maps 6 content controls from Excel row |
| 2c | Convert to PDF | Word Online connector |
| 2d | Create file | Save PDF to `/Output/PDF/` |
| 2e | Send email | To row's Email, PDF attached |
| 2f | Update row | `SendStatus = Sent`, `SentAt = utcNow()` |

### Error handling

Steps 2b–2f wrapped in a **Try** scope. A **Catch** scope (configured to run on failure/timeout/skip) writes `SendStatus = Error` and `ErrorMessage` back to the row. One bad row does not kill the batch.

## Email

**Subject:** Your Mobility Parking Pass — [EventName], [EventDate]

**Body:**

Hi [FirstName],

Please find your Mobility Parking Pass attached for [EventName] on [EventDate].

Please print this pass and display it clearly on your dashboard when parked.

If you have any questions, please contact [contact email/phone].

Kind regards,
Eden Park

**Attachment:** `[PermitNumber] - [LastName].pdf`

## Deliverable Bundle

```
jem/
├── templates/
│   ├── LetterTemplate.docx        # Pass with 6 content controls
│   ├── Data.xlsx                   # tblLetters table with all columns
│   └── email-body.txt             # Email text with placeholders
├── solution/
│   └── MobilityParkingPass.zip    # Exported Power Automate solution
├── docs/
│   ├── it-setup-guide.md          # Import, connect, configure
│   ├── end-user-guide.md          # Paste CSV, dedup, Run
│   └── flow-build-guide.md        # Rebuild/modify reference
├── samples/
│   ├── SampleData.csv             # 3 fake test rows
│   └── SampleOutput.pdf           # Example generated pass
└── README.md
```

## Responsibilities

| Who | Does what |
|---|---|
| Developer (you) | Recreate Word template, build flow in M365, export solution, assemble bundle |
| Client IT | Import solution, create connections, set up folder structure |
| End user | Paste CSV into Excel, dedup with conditional formatting, click Run |
