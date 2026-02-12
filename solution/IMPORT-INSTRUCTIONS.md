# How to Import the Power Automate Flow

Two methods — try Method 1 first. If it doesn't work, use Method 2.

---

## Method 1: Import Package (preferred)

1. Go to https://make.powerautomate.com
2. Click **My flows** in the left sidebar
3. Click **Import** > **Import Package (Legacy)** in the top toolbar
4. Upload `MobilityParkingPass.zip`
5. Power Automate will show the package contents and ask you to configure connections
6. For each connection (Excel Online, Word Online, OneDrive, Outlook):
   - Click **Select during import**
   - Choose an existing connection or create a new one
   - Sign in with the account that has access to the OneDrive/SharePoint files
7. Click **Import**
8. Once imported, open the flow and configure the file paths (see below)

---

## Method 2: Code View (backup)

If the package import doesn't work:

1. Go to https://make.powerautomate.com
2. Click **Create** > **Instant cloud flow**
3. Name it: `Mobility Parking Pass`
4. Select trigger: **Manually trigger a flow**
5. Click **Create**
6. In the flow editor, click the **Code View** toggle (top right, looks like `</>`)
7. Select all the existing JSON and delete it
8. Open `flow-definition.json` from this folder
9. Copy the contents of `properties.definition` (just the definition object, not the whole file)
10. Paste it into the code view
11. Click **Save**
12. Switch back to the visual designer and configure the file paths (see below)

---

## After Import: Configure File Paths

The flow has placeholder values for file locations. You need to set these for your environment.

Open the flow in the visual editor and update these actions:

### 1. List rows present in a table
- **Location:** OneDrive for Business (or your SharePoint site)
- **Document Library:** OneDrive (or your library name)
- **File:** Navigate to `/Automation/Input/Data.xlsx`
- **Table:** Select `tblLetters`

### 2. Populate a Microsoft Word template
- **Location:** OneDrive for Business (or your SharePoint site)
- **Document Library:** OneDrive (or your library name)
- **File:** Navigate to `/Automation/Templates/LetterTemplate.docx`
- The 6 template fields (FullName, PermitNumber, etc.) should auto-map if the tags match

### 3. Create DOCX file
- **Folder Path:** `/Automation/Output/DOCX`
- Create this folder in OneDrive if it doesn't exist

### 4. Convert Word Document to PDF
- **Location:** Should auto-detect from the DOCX file created in step 3
- If not, set to the same OneDrive/SharePoint location

### 5. Create PDF file
- **Folder Path:** `/Automation/Output/PDF`

### 6. Update a row — Success (inside Try scope)
- **Location, File, Table:** Same as step 1

### 7. Update a row — Error (inside Catch scope)
- **Location, File, Table:** Same as step 1

### 8. Send an email
- Verify the sending account is correct
- Update the email body contact details if needed

---

## After Configuration: Test

1. Add 2-3 test rows to `Data.xlsx` (use your own email addresses)
2. Run the flow manually
3. Verify: emails arrive, PDFs are correct, Excel rows update to "Sent"
4. Check the `/Automation/Output/PDF/` folder for the generated files

See `docs/end-user-guide.md` for ongoing usage.
