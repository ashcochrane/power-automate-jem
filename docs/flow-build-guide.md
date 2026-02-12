# Flow Build Guide — Developer Reference

This guide is your reference for building (or rebuilding) the Power Automate flow from scratch.

## Overview

The flow reads unsent rows from Excel, generates a personalised parking pass PDF for each, emails it, and marks the row as sent.

## Prerequisites

- An M365 tenant with Power Automate, Excel Online (Business), Word Online (Business), OneDrive for Business, and Outlook
- The folder structure and templates set up per the IT Setup Guide
- `LetterTemplate.docx` with content controls in place (see Word Template section below)

## Word Template Setup

The template must be a `.docx` file with Plain Text Content Controls. To set these up:

1. Open Word and enable the **Developer** tab:
   - **Mac:** Word > Preferences > Ribbon & Toolbar > enable Developer
   - **Windows:** File > Options > Customize Ribbon > enable Developer
2. Open `LetterTemplate.docx`
3. For each dynamic field, place your cursor where the value should appear, then:
   - Developer tab > **Plain Text Content Control**
   - Select the control > click **Properties**
   - Set both **Title** and **Tag** to the exact values below

| Tag | What it displays | Maps from Excel |
|---|---|---|
| `FullName` | Recipient's full name | Compose: `FirstName` + " " + `LastName` |
| `PermitNumber` | Mobility parking permit # | `Mobility or Disability Parking Permit Number` |
| `VehiclePlate` | Vehicle registration | `Vehicle number plate` |
| `VehicleType` | Vehicle type | `Vehicle type` |
| `EventName` | Event name | `Event` (column 8) |
| `EventDate` | Event date | `Event` (column 9) |

4. Save to `/Automation/Templates/LetterTemplate.docx`

## Excel Setup

Open `Data.xlsx` and set up the table:

1. Select the header row and at least one data row
2. **Insert > Table** (tick "My table has headers")
3. Name the table: **Table Design > Table Name:** `tblLetters`
4. Ensure these columns exist (add if missing):

**Data columns** (from CSV):
- `Permit No.`
- `Name (First)`
- `Name (Middle)`
- `Name (Last)`
- `Name (Suffix)`
- `Email (Enter Email)`
- `Phone`
- `Event` (event name)
- (unnamed column 9 — event date)
- `Mobility or Disability Parking Permit Number`
- `Vehicle number plate`
- `Vehicle type`

**Automation columns** (add these at the end):
- `SendStatus` (text — blank, Sent, or Error)
- `SentAt` (text — timestamp)
- `ErrorMessage` (text)

## Create the Flow

Go to [Power Automate](https://make.powerautomate.com) > **Create** > **Instant cloud flow** > **Manually trigger a flow**.

### Step 1 — List rows present in a table

| Setting | Value |
|---|---|
| Connector | Excel Online (Business) |
| Action | List rows present in a table |
| Location | OneDrive for Business (or SharePoint) |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |
| Filter Query | `SendStatus ne 'Sent'` |
| Top Count | `50` (increase once tested) |

### Step 2 — Apply to each

Power Automate creates this automatically from the List rows output.

Inside the loop, create a **Scope** called `Try` and put all the following actions inside it.

### Step 2a — Compose safe filename

| Setting | Value |
|---|---|
| Action | Data Operations > Compose |
| Inputs | Expression: `replace(replace(replace(concat(items('Apply_to_each')?['Mobility or Disability Parking Permit Number'],' - ',items('Apply_to_each')?['Name (Last)']),'/','-'),'\','-'),':','-')` |

Rename this action to `SafeBaseName`.

### Step 2b — Compose full name

| Setting | Value |
|---|---|
| Action | Data Operations > Compose |
| Inputs | Expression: `concat(items('Apply_to_each')?['Name (First)'],' ',items('Apply_to_each')?['Name (Last)'])` |

Rename this action to `FullName`.

### Step 2c — Populate Word template

| Setting | Value |
|---|---|
| Connector | Word Online (Business) |
| Action | Populate a Microsoft Word template |
| Location | OneDrive for Business (or SharePoint) |
| File | `/Automation/Templates/LetterTemplate.docx` |
| FullName | Output of `FullName` compose |
| PermitNumber | `Mobility or Disability Parking Permit Number` from Excel |
| VehiclePlate | `Vehicle number plate` from Excel |
| VehicleType | `Vehicle type` from Excel |
| EventName | `Event` (column 8) from Excel |
| EventDate | Column 9 from Excel |

### Step 2d — Convert Word document to PDF

| Setting | Value |
|---|---|
| Connector | Word Online (Business) |
| Action | Convert Word Document to PDF |
| File | Body output from Step 2c (the populated template content) |

If this action is not available, try: OneDrive for Business > Convert file.

### Step 2e — Create PDF file

| Setting | Value |
|---|---|
| Connector | OneDrive for Business (or SharePoint) |
| Action | Create file |
| Folder path | `/Automation/Output/PDF` |
| File name | Expression: `concat(outputs('SafeBaseName'),'.pdf')` |
| File content | Output from Step 2d (PDF content) |

### Step 2f — Send email

| Setting | Value |
|---|---|
| Connector | Outlook |
| Action | Send an email (V2) |
| To | `Email (Enter Email)` from Excel |
| Subject | Expression: `concat('Your Mobility Parking Pass — ',items('Apply_to_each')?['Event'],' ',<event date column>)` |
| Body | See `templates/email-body.txt` — replace placeholders with dynamic content |
| Attachment Name | Expression: `concat(outputs('SafeBaseName'),'.pdf')` |
| Attachment Content | File content from Step 2e (or use Get file content on the created PDF) |

### Step 2g — Update Excel row (success)

| Setting | Value |
|---|---|
| Connector | Excel Online (Business) |
| Action | Update a row |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |
| Key Column | `Permit No.` |
| Key Value | `Permit No.` from current row |
| SendStatus | `Sent` |
| SentAt | Expression: `utcNow()` |

### Error Handling — Catch scope

After the `Try` scope, add a new **Scope** called `Catch`.

Configure the Catch scope's **run after** settings:
- Tick: **has failed**
- Tick: **has timed out**
- Tick: **is skipped**

Inside the Catch scope, add:

#### Update Excel row (error)

| Setting | Value |
|---|---|
| Connector | Excel Online (Business) |
| Action | Update a row |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |
| Key Column | `Permit No.` |
| Key Value | `Permit No.` from current row |
| SendStatus | `Error` |
| ErrorMessage | Expression: `string(result('Try'))` (captures the error details) |

## Export the Solution

Once the flow is tested and working:

1. Go to **Solutions** in Power Automate
2. If the flow is not in a solution, create one and add the flow to it
3. Click **Export** on the solution
4. Save the `.zip` to `solution/MobilityParkingPass.zip` in this repo

## Testing Checklist

- [ ] Word template fields populate correctly (all 6 fields)
- [ ] PDF conversion works
- [ ] Filename is valid (no illegal characters)
- [ ] Email arrives with correct subject, body, and attachment
- [ ] Excel row updates to `Sent` with timestamp
- [ ] Row with missing email triggers Catch scope and writes `Error`
- [ ] Row with special characters in name produces a valid filename
- [ ] Running the flow twice does not re-send already-sent rows
- [ ] A batch of 10+ rows processes without errors
