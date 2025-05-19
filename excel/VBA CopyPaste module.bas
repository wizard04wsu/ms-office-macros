Attribute VB_Name = "CopyPaste"

'Be aware that `Range.Copy` and `Range.PasteSpecial` use the clipboard. This can
' cause problems if you copy/cut anything else on your computer while they are
' being used here.

'Copy & paste values only (without using the clipboard). Returns the range that
' the values were copied to.
Public Function CopyPasteValues(fromRange As range, toCell As range) As range
	
	Dim rows As Long, columns As Long, toRange As range
	
	rows = fromRange.rows.Count
	columns = fromRange.columns.Count
	
	Set toRange = toCell.Worksheet.range(toCell, toCell.Offset(rows - 1, columns - 1))
	
	toRange.Value = fromRange.Value
	
	Set CopyPasteValues = toRange
	
End Function

'Copy & paste everything (values, formats, comments, et al.). This uses the
' clipboard.
Public Sub CopyPaste(fromRange As range, toRange As range)
	
	fromRange.Copy toRange
	
End Sub

'Use the Paste Special feature to copy & paste only specific attributes of a
' range. This uses the clipboard.
Public Sub CopyPasteSpecial(fromRange As range, toRange As range, _
 Optional pasteType As XlPasteType = xlPasteAll, _
 Optional pasteOperation As XlPasteSpecialOperation = xlPasteSpecialOperationNone, _
 Optional skipBlanks As Boolean = False, _
 Optional transpose As Boolean = False)
	
	fromRange.Copy
	Call toRange.PasteSpecial(pasteType, pasteOperation, skipBlanks, transpose)
	
End Sub
