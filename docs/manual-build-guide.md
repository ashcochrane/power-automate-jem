# Manual Build Guide — Mobility Parking Pass Automation

A complete, step-by-step guide to building this automation from scratch. Follow in order.

---

## Phase 1: Set Up the Folder Structure (5 minutes)

### 1.1 Open OneDrive for Business (or SharePoint)

1. Go to https://portal.office.com and sign in
2. Click the **OneDrive** icon (or navigate to your SharePoint site)

### 1.2 Create the folders

Create this exact structure:

```
Automation/
├── Input/
├── Templates/
└── Output/
    └── PDF/
```

How to do it:

1. Click **New > Folder**, name it `Automation`
2. Open `Automation`, click **New > Folder**, name it `Input`
3. Click **New > Folder**, name it `Templates`
4. Click **New > Folder**, name it `Output`
5. Open `Output`, click **New > Folder**, name it `PDF`

---

## Phase 2: Prepare the Excel File (15 minutes)

### 2.1 Create the workbook

1. Inside `/Automation/Input/`, click **New > Excel workbook**
2. Name it `Data.xlsx`
3. Open it in Excel Online (or download and open in desktop Excel)

### 2.2 Set up the column headers

In Row 1, type these headers across the columns exactly as shown:

| Column | Header |
|---|---|
| A | `Permit No.` |
| B | `Name (First)` |
| C | `Name (Middle)` |
| D | `Name (Last)` |
| E | `Name (Suffix)` |
| F | `Email (Enter Email)` |
| G | `Phone` |
| H | `Event` |
| I | `EventDate` |
| J | `Mobility or Disability Parking Permit Number` |
| K | `Vehicle number plate` |
| L | `Vehicle type` |
| M | `SendStatus` |
| N | `SentAt` |
| O | `ErrorMessage` |

Columns A–L match the CSV export from the Eden Park website. Columns M–O are for automation tracking.

### 2.3 Convert the range to a Table

1. Click on cell **A1**
2. Press **Ctrl+T** (Windows) or **Cmd+T** (Mac)
3. Tick **My table has headers** and click OK
4. The range is now a formatted Table

### 2.4 Name the table

1. Click anywhere inside the table
2. Go to the **Table Design** tab (appears when the table is selected)
3. On the far left, find the **Table Name** field
4. Replace the default name with: `tblLetters`
5. Press Enter

### 2.5 Set up conditional formatting for duplicates

This helps the operator spot duplicates before running the flow.

**For Permit Number:**

1. Select the entire **Mobility or Disability Parking Permit Number** column (click the column header in the table)
2. Go to **Home > Conditional Formatting > Highlight Cell Rules > Duplicate Values**
3. Choose a highlight colour (e.g. light red) and click OK

**For Vehicle Plate:**

1. Select the entire **Vehicle number plate** column
2. Go to **Home > Conditional Formatting > Highlight Cell Rules > Duplicate Values**
3. Choose a different highlight colour (e.g. yellow) and click OK

### 2.6 Save the workbook

If you're in Excel Online, it saves automatically. If you're in desktop Excel, save and close, then upload to `/Automation/Input/` in OneDrive.

---

## Phase 3: Build the Word Template (20–30 minutes)

### 3.1 Enable the Developer tab

**On Mac:**
1. Open Word
2. Go to **Word > Preferences > Ribbon & Toolbar**
3. In the right column, tick **Developer**
4. Click Save

**On Windows:**
1. Open Word
2. Go to **File > Options > Customize Ribbon**
3. In the right column, tick **Developer**
4. Click OK

### 3.2 Recreate the pass layout

Open a new blank document and recreate the Eden Park mobility parking pass layout. The pass is a single A4 page that folds in half:

**Top half (becomes the back when folded):**
- "FOLD HERE" line at the very top
- "By using this pass you accept and will abide by the below conditions."
- Bullet list of all terms and conditions
- Disclaimer paragraph at the bottom of this section

**Bottom half (the front of the pass):**
- Eden Park logo
- Event name and date
- "MOBILITY PARKING PASS" heading
- Fields for: NAME, PERMIT, VEHICLE PLATE, VEHICLE TYPE
- Any additional branding or instructions

Match the fonts, colours, and layout from the original InDesign template as closely as possible.

### 3.3 Add content controls for each dynamic field

For each field that changes per person, you will insert a Plain Text Content Control.

**To add a content control:**

1. Place your cursor exactly where the dynamic value should appear (e.g. after "NAME: ")
2. Go to the **Developer** tab
3. Click **Plain Text Content Control** (the `Aa` icon)
4. A grey placeholder box appears — this is the content control

**To set the tag (critical — Power Automate maps to these tags):**

1. Click on the content control to select it
2. Go to the **Developer** tab
3. Click **Properties**
4. In the dialog:
   - **Title:** Enter the tag name (see table below)
   - **Tag:** Enter the exact same tag name
5. Click OK

**Do this for all 6 fields:**

| Where on the pass | Title & Tag | Example value |
|---|---|---|
| After "NAME: " | `FullName` | Magda Clayton |
| After "PERMIT: " (the permit number) | `PermitNumber` | LT38445958 |
| After "PLATE: " or vehicle plate field | `VehiclePlate` | NNN501 |
| After "VEHICLE TYPE: " | `VehicleType` | Standard vehicle |
| The event name area | `EventName` | Blues v Chiefs |
| The event date area | `EventDate` | Saturday 14 February |

### 3.4 Test the content controls

Click on each content control — you should see the tag name appear. Type some test text into each one to confirm they work. Then undo back to the placeholder state.

### 3.5 Save the template

1. Save as `LetterTemplate.docx` (standard .docx, not .dotx)
2. Upload to `/Automation/Templates/` in OneDrive

---

## Phase 4: Create the Power Automate Flow (30–45 minutes)

### 4.1 Start a new flow

1. Go to https://make.powerautomate.com
2. Click **Create** in the left sidebar
3. Choose **Instant cloud flow**
4. Name it: `Mobility Parking Pass`
5. Select trigger: **Manually trigger a flow**
6. Click **Create**

### 4.2 Add Step 1 — List rows from Excel

1. Click **+ New step**
2. Search for: `Excel Online (Business)`
3. Select the action: **List rows present in a table**
4. Configure:

| Field | Value |
|---|---|
| Location | OneDrive for Business |
| Document Library | OneDrive (or your SharePoint library) |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |

5. Click **Show advanced options**
6. Set:

| Field | Value |
|---|---|
| Filter Query | `SendStatus ne 'Sent'` |
| Top Count | `20` |

The filter query ensures only unsent rows are processed. Set Top Count to 20 while testing — increase it later for production batches.

### 4.3 Add Step 2 — Apply to each

Power Automate will automatically wrap the next actions in an **Apply to each** loop when you reference the Excel rows output. If it doesn't:

1. Click **+ New step**
2. Search for **Apply to each**
3. In the "Select an output from previous steps" field, choose the `value` output from the List rows action

Everything from here goes **inside** this Apply to each loop.

### 4.4 Add a Try scope

Inside the Apply to each:

1. Click **Add an action**
2. Search for **Scope**
3. Rename it to `Try`

All of the following actions (4.5 through 4.10) go **inside this Try scope**.

### 4.5 Compose the full name

Inside the Try scope:

1. Click **Add an action**
2. Search for **Compose** (under Data Operations)
3. Rename it to `FullName`
4. In the **Inputs** field, click **Expression** and enter:

```
concat(items('Apply_to_each')?['Name (First)'],' ',items('Apply_to_each')?['Name (Last)'])
```

5. Click OK

### 4.6 Compose a safe file name

1. Click **Add an action** (still inside Try scope)
2. Search for **Compose**
3. Rename it to `SafeBaseName`
4. In the **Inputs** field, click **Expression** and enter:

```
replace(replace(replace(concat(items('Apply_to_each')?['Mobility or Disability Parking Permit Number'],' - ',items('Apply_to_each')?['Name (Last)']),'/','-'),'\','-'),':','-')
```

5. Click OK

This creates a filename like `LT38445958 - Clayton` with dangerous characters stripped out.

### 4.7 Populate the Word template

1. Click **Add an action** (still inside Try scope)
2. Search for `Word Online (Business)`
3. Select: **Populate a Microsoft Word template**
4. Configure:

| Field | Value |
|---|---|
| Location | OneDrive for Business |
| Document Library | OneDrive |
| File | `/Automation/Templates/LetterTemplate.docx` |

5. Once the file is selected, Power Automate reads the content controls and shows fields for each tag. Map them:

| Template Field | Value (Dynamic content) |
|---|---|
| FullName | Output of the `FullName` compose step |
| PermitNumber | `Mobility or Disability Parking Permit Number` from Excel |
| VehiclePlate | `Vehicle number plate` from Excel |
| VehicleType | `Vehicle type` from Excel |
| EventName | `Event` from Excel (column H — the event name) |
| EventDate | `EventDate` from Excel (column I — the event date) |

If the template fields don't appear, double-check that the .docx has content controls with the correct tags.

### 4.8 Convert to PDF

1. Click **Add an action** (still inside Try scope)
2. Search for `Word Online (Business)`
3. Select: **Convert Word Document to PDF**
4. In the **File** field, select the **Microsoft Word Document** output (the body) from the Populate step

If you don't see "Convert Word Document to PDF":
- Try searching for `OneDrive for Business` > **Convert file**
- As a last resort, you may need a third-party connector (Encodian, Plumsail) — but most business tenants have the Word Online connector

### 4.9 Save the PDF

1. Click **Add an action** (still inside Try scope)
2. Search for `OneDrive for Business` (or SharePoint)
3. Select: **Create file**
4. Configure:

| Field | Value |
|---|---|
| Folder Path | `/Automation/Output/PDF` |
| File Name | Expression: `concat(outputs('SafeBaseName'),'.pdf')` |
| File Content | The **PDF document** output from the Convert step |

### 4.10 Send the email

1. Click **Add an action** (still inside Try scope)
2. Search for `Outlook`
3. Select: **Send an email (V2)**
4. Configure:

| Field | Value |
|---|---|
| To | `Email (Enter Email)` from Excel (dynamic content) |
| Subject | Expression: `concat('Your Mobility Parking Pass — ',items('Apply_to_each')?['Event'],' ',items('Apply_to_each')?['EventDate'])` |
| Body | See below |

**Email body** (paste this into the Body field, then replace the bracketed parts with dynamic content):

```
Hi [Name (First) from Excel],

Please find your Mobility Parking Pass attached for [Event from Excel] on [EventDate from Excel].

Please print this pass and display it clearly on your dashboard when parked.

If you have any questions, please contact [insert contact email here].

Kind regards,
Eden Park
```

Replace each `[bracketed part]` by clicking in the text and selecting the matching dynamic content from Excel.

**Add the attachment:**

1. Click **Show advanced options**
2. Under **Attachments**, click the "T" icon to switch to array mode, or click **Add new item**
3. Set:

| Field | Value |
|---|---|
| Attachments Name - 1 | Expression: `concat(outputs('SafeBaseName'),'.pdf')` |
| Attachments Content - 1 | The **File Content** output from the Create file step (4.9) |

### 4.11 Update the Excel row — success

1. Click **Add an action** (still inside Try scope)
2. Search for `Excel Online (Business)`
3. Select: **Update a row**
4. Configure:

| Field | Value |
|---|---|
| Location | OneDrive for Business |
| Document Library | OneDrive |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |
| Key Column | `Permit No.` |
| Key Value | `Permit No.` from current Excel row (dynamic content) |
| SendStatus | `Sent` |
| SentAt | Expression: `utcNow()` |

This marks the row so it won't be processed again.

---

## Phase 5: Add Error Handling (10 minutes)

### 5.1 Add a Catch scope

1. Click **Add an action** — but make sure you're adding it **after** the Try scope, still inside the Apply to each loop (not inside Try)
2. Search for **Scope**
3. Rename it to `Catch`

### 5.2 Configure Catch to run on failure

1. Click the **three dots** (⋯) on the Catch scope header
2. Select **Configure run after**
3. Untick: **is successful**
4. Tick: **has failed**
5. Tick: **has timed out**
6. Tick: **is skipped**
7. Click **Done**

This means the Catch scope only runs when the Try scope fails.

### 5.3 Add the error update action inside Catch

1. Inside the Catch scope, click **Add an action**
2. Search for `Excel Online (Business)` > **Update a row**
3. Configure:

| Field | Value |
|---|---|
| Location | OneDrive for Business |
| Document Library | OneDrive |
| File | `/Automation/Input/Data.xlsx` |
| Table | `tblLetters` |
| Key Column | `Permit No.` |
| Key Value | `Permit No.` from current Excel row |
| SendStatus | `Error` |
| ErrorMessage | Expression: `string(result('Try'))` |

The `result('Try')` expression captures the error details from whichever action failed inside the Try scope.

---

## Phase 6: Test the Flow (15 minutes)

### 6.1 Add test data

1. Open `Data.xlsx` in `/Automation/Input/`
2. Add 3 test rows using the sample data from `samples/SampleData.csv`
3. **Change the email addresses** in all 3 rows to your own email address

### 6.2 Run the flow

1. Go to Power Automate
2. Find `Mobility Parking Pass`
3. Click **Run**
4. Wait for it to complete (watch the run history)

### 6.3 Verify each step

Check every one of these:

| Check | What to look for |
|---|---|
| Flow completed | Run history shows "Succeeded" |
| PDF files created | Look in `/Automation/Output/PDF/` for 3 PDF files |
| PDF content | Open each PDF — all 6 fields should be filled in correctly |
| Filenames | Should be like `LT38000001 - Smith.pdf` |
| Emails received | Check your inbox for 3 emails |
| Email subject | Should show the event name and date |
| Email attachment | PDF should be attached and openable |
| Excel updated | `SendStatus` should say `Sent` for all 3 rows |
| Timestamps | `SentAt` column should have a timestamp |

### 6.4 Test error handling

1. Add a 4th test row with an invalid or blank email address
2. Run the flow again
3. The 3 already-sent rows should be skipped (filter query)
4. The bad row should show `SendStatus = Error` and `ErrorMessage` should contain details
5. Delete the test rows when done

### 6.5 Test idempotency

1. Run the flow again with no new rows (all marked as Sent)
2. Confirm the flow completes without doing anything — no duplicate emails

---

## Phase 7: Export the Solution (5 minutes)

### 7.1 Add the flow to a solution

1. In Power Automate, go to **Solutions** in the left sidebar
2. Click **New solution**
3. Fill in:
   - **Display name:** Mobility Parking Pass
   - **Publisher:** Select default or create one
   - **Version:** 1.0.0
4. Click **Create**
5. Open the solution
6. Click **Add existing > Cloud flow**
7. Select your `Mobility Parking Pass` flow
8. Click **Add**

### 7.2 Export

1. From the Solutions list, click the **three dots** (⋯) on your solution
2. Click **Export**
3. Choose **Unmanaged** (this lets the client's IT make changes if needed)
4. Click **Export**
5. Save the downloaded `.zip` file

### 7.3 Add to the repo

Save the `.zip` file as `solution/MobilityParkingPass.zip` in the Jem repo.

---

## Phase 8: Prepare the Deliverable (5 minutes)

Your final deliverable bundle should contain:

```
jem/
├── templates/
│   ├── LetterTemplate.docx        ← the Word template you built
│   ├── Data.xlsx                   ← the Excel file you set up
│   └── email-body.txt             ← email text reference
├── solution/
│   └── MobilityParkingPass.zip    ← the exported solution
├── docs/
│   ├── it-setup-guide.md          ← hand this to the client's IT
│   ├── end-user-guide.md          ← hand this to the operator
│   ├── flow-build-guide.md        ← keep for yourself
│   └── manual-build-guide.md      ← this document
├── samples/
│   ├── SampleData.csv             ← test data
│   └── SampleOutput.pdf           ← example of a generated pass (add one)
└── README.md
```

Hand the client's IT team:
1. The `solution/MobilityParkingPass.zip` file
2. The `templates/` folder contents
3. The `docs/it-setup-guide.md`
4. The `docs/end-user-guide.md`
5. A sample PDF output so they know what success looks like

---

## Quick Reference: All Expressions Used in the Flow

| Action | Expression |
|---|---|
| FullName | `concat(items('Apply_to_each')?['Name (First)'],' ',items('Apply_to_each')?['Name (Last)'])` |
| SafeBaseName | `replace(replace(replace(concat(items('Apply_to_each')?['Mobility or Disability Parking Permit Number'],' - ',items('Apply_to_each')?['Name (Last)']),'/','-'),'\','-'),':','-')` |
| PDF filename | `concat(outputs('SafeBaseName'),'.pdf')` |
| Email subject | `concat('Your Mobility Parking Pass — ',items('Apply_to_each')?['Event'],' ',items('Apply_to_each')?['EventDate'])` |
| Sent timestamp | `utcNow()` |
| Error message | `string(result('Try'))` |
| Filter query | `SendStatus ne 'Sent'` |
