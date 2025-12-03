# PST File Troubleshooting Guide

## Issue with test.pst

Your `test.pst` file cannot be parsed by the pst-extractor library. Here's what we know:

### File Information
- **Size**: 46MB
- **Format**: Microsoft Outlook email folder (<=2002) - ANSI format
- **Error**: `PSTFile::findBtreeItem Unable to find 97`

### Why This Happens

The pst-extractor library is a pure JavaScript implementation that works with most PST files, but it has known limitations with:

1. **Corrupted PST files** - Files with internal structure damage
2. **Encrypted PST files** - Password-protected archives
3. **Certain older ANSI PST formats** - Some PST 97-2002 files have unusual structures
4. **PST files created by third-party tools** - Non-standard implementations

The error "Unable to find 97 is desc: true" suggests the B-tree index structure in your PST file is either:
- Corrupted
- Uses a format variation not supported by pst-extractor
- Has been modified by repair tools

## Solutions

### Option 1: Repair the PST File (Recommended)

Use Microsoft's built-in repair tool:

#### On Windows:
1. **Use SCANPST.EXE** (Inbox Repair Tool)
   ```
   Location: C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE
   (Path varies by Office version)
   ```

2. **Steps:**
   - Close Outlook
   - Run SCANPST.EXE
   - Select your test.pst file
   - Click "Start" to scan
   - If errors found, click "Repair"
   - After repair, try fixECalendar again

#### On Mac:
Microsoft doesn't provide SCANPST for Mac, but you can:
- Use Outlook for Mac to import the PST and export a new one
- Use a Windows VM to run SCANPST

### Option 2: Use Outlook to Re-export

If you have access to Outlook:

1. **Open the PST in Outlook**
   - File → Open → Outlook Data File → Select test.pst

2. **Export the Calendar**
   - File → Open & Export → Import/Export
   - Choose "Export to a file"
   - Select "Outlook Data File (.pst)"
   - Choose only the Calendar folder
   - Save as "test-repaired.pst"

3. **Try fixECalendar again**
   ```bash
   fixECalendar test-repaired.pst --output calendar.ics
   ```

### Option 3: Alternative Export Methods

#### Export Directly from Outlook:
1. Open Outlook
2. File → Save Calendar
3. Choose "iCalendar Format (.ics)"
4. This creates an ICS file directly (no conversion needed!)

#### Use Outlook's built-in export:
1. File → Open & Export → Import/Export
2. Choose "Export to a file"
3. Select "iCalendar Format (.ics)"
4. Select the Calendar folder

### Option 4: Try Alternative PST Libraries

If the above doesn't work, there are other PST parsing options:

#### 1. **readpst** (Linux/Mac command-line tool)
```bash
# Install
brew install libpst  # Mac
sudo apt-get install pst-utils  # Linux

# Convert PST to mbox or text
readpst -r -o output_dir test.pst

# Then use another tool to convert calendar entries
```

#### 2. **PST Walker** (Windows GUI tool)
- Free tool for exploring PST files
- Can export to various formats
- Download: https://www.pstwalker.com/

#### 3. **Python: pypff** (Advanced)
```python
import pypff

pst_file = pypff.file()
pst_file.open("test.pst")
root = pst_file.get_root_folder()
# Navigate to calendar and extract
```

## Testing with a Different PST File

To verify fixECalendar works, try creating a test PST:

### On Windows with Outlook:
1. Create a new PST file in Outlook
2. Add a few test calendar events
3. Export that PST
4. Try fixECalendar with the new PST

### Generate a Sample PST:
If you have multiple PST files, try fixECalendar with your other archives to see if they work.

## Next Steps

1. **Try repairing test.pst** using SCANPST.EXE (Windows)
2. **If you have other PST files**, test with those
3. **Report the issue** to pst-extractor:
   - https://github.com/epfromer/pst-extractor/issues
   - Include: PST format, error message, file size

## Alternative: Use the Working PST Files

If you have other PST files that work:

```bash
# Process all working PST files
fixECalendar working1.pst working2.pst working3.pst

# Check what was imported
fixECalendar --db-stats

# Export
fixECalendar --output calendar.ics
```

## For Your Large 6.5GB Files

The good news: Newer PST files (Unicode format, post-2003) typically work better with pst-extractor. Your large 6.5GB files are likely in Unicode format and should work fine.

Test with one of your large files:
```bash
fixECalendar /path/to/large-file.pst --verbose
```

## Technical Details of the Error

The error `findBtreeItem Unable to find 97` means:
- The library is trying to read node ID 97 from the PST's B-tree
- This node should exist according to the file structure
- But it's either missing or the index is corrupted

This typically indicates:
- File corruption
- Non-standard PST format variation
- Issues with how the PST was created/exported

## Contact & Support

If you continue having issues:

1. Check if your other PST files work
2. Try the repair option
3. Consider using Outlook's built-in ICS export as a workaround
4. Open an issue on GitHub with:
   - PST file size
   - PST format (ANSI/Unicode)
   - Outlook version that created it
   - Error message

The fixECalendar tool itself is working correctly - the issue is specific to how this particular PST file is structured.
