# Microsoft Office Macros and VBA Scripts

## General

### Arrays

#### GetLength(_arr_)
Returns the number of items in the array.

#### GetUBound(_arr_)
Returns the array's highest index. If the array is empty, Null is returned.

#### GetLBound(_arr_)
Returns the array's lowest index. If the array is empty, Null is returned.

#### Push(_arr_, _item_)
Adds an item to the end of a dynamic array.

#### Pop(_arr_)
Removes and returns the last item in a dynamic array, or returns Null if the array is already empty.

#### Unshift(_arr_, _item_)
Adds an item to the beginning of a dynamic array.

#### Shift(_arr_)
Removes and returns the first item in a dynamic array, or returns Null if the array is already empty.

#### assignAtIndex(_arr_, _index_, _item_)
Assigns a value to a position in an array.

### Math

#### StandardRound(_pValue_, Optional _pDecimalPlaces_ = 0, Optional _pSymmetricRounding = False)
This is a slightly modified script based on one from [TechOnTheNet.com]([https://www.techonthenet.com/excel/formulas/round_vba.php).

Returns the value rounded to the specified number of decimal places.

Negative values can be rounded differently depending on `pSymmetricRounding`. For example:
If `False`, -1.5 would round to -1.
If `True`, -1.5 would round to -2.

### Files

#### GetFolderName(Optional _pathAndOrFileFilter_)
Displays the folder picker dialog box, optionally with an initial directory and/or file type filter. Returns the selected path or an empty string.

#### GetFileName(Optional _pathAndOrFileFilter_)
Displays the file picker dialog box, optionally with an initial directory and/or file type filter. Returns the selected path or an empty string.

#### ChangeDirectory(_path_)
Changes the current directory. Returns `True` if successful.

#### GetFileCount(_thePath_, Optional _fileFilter_ = "\*.\*")
Returns the number of files in `thePath` and its subdirectories that match the filter.

## Excel

### Copy & Paste

Be aware that `Range.Copy` and `Range.PasteSpecial` use the clipboard. This can cause problems if you copy/cut anything else on your computer while they are being used by this script.

#### CopyPasteValues(_fromRange_, _toCell_)
Copy & paste values only (without using the clipboard). Returns the range that the values were copied to.

#### CopyPaste(_fromRange_, _toRange_)
Copy & paste everything (values, formats, comments, et al.). This uses the clipboard.

#### CopyPasteSpecial(_fromRange_, _toRange_)
Use the Paste Special feature to copy & paste only specific attributes of a range. This uses the clipboard.

### Miscellaneous

#### GetLastRow(Optional _sheet_)
Returns the number of the last row containing a cell with data.

#### ToLowerCase()
Converts values in the selected range to lower case.

#### ToUpperCase()
Converts values in the selected range to upper case.

#### ToProperCase()
Converts values in the selected range to title case.

#### ConvertToText()
Applies number format to the selected range.

#### ConvertToGeneral()
Applies general format to the selected range.

#### ForceNumberFormat()
Updates selected cell values to match the number format of their cell.  
If there are formulas in the selection, the user will be asked whether to convert them to values or not.

#### PercentToInteger()
Multiplies numeric values in the selection by 100 (e.g., `0.5` becomes `50`).  
Converts string values denoting a percentage to a number value (e.g., `50%` becomes `50`).

### Macro-enabled Templates

#### Crosstab to List
In the active sheet of a selected workbook, converts the data from crosstab-format to list format.

#### List Files
Generates a list of all the files in a selected directory and its subdirectories.

#### Collate Workbooks
Combines data from the active sheet of each of the workbooks in the specified folder and its subfolders.  
If there is a header row, it gives the option to add a new column for each unique header it finds. Filenames can optionally be included in column A.

#### Import PDF Forms
This will import the field values from the PDF forms selected, appending each form's data to the next empty row in this spreadsheet. Currently, only text and checkbox fields are imported.

New column headers will be added as new field names are discovered. The fields are not imported in any particular order. If you would like the columns in a certain order, you can pre-populate the column headers. Before importing, put the matching field names in the order you wish, starting in column B.

Requires Adobe Acrobat.

## Web Browser Scripts (JavaScript)

#### ExcelWorkbook(tableElems, sheetnames = [], documentProperties = {})
Class to convert an array of HTML table elements into sheets in an XML-based Excel document. This does not retain styling.

`documentProperties` can contain Excel document metadata (e.g., Title, Subject, Author, Keywords, Description, Created).

Properties:
- `xml` - The raw XML of the Excel workbook.
- `href` - A data URL of type `application/vnd.ms-excel` to download the file.

Methods:
- `revoke` - Call this when finished to let the browser know not to keep the reference to the file any longer.

Behavior when dealing with malformed HTML tables is undefined.
