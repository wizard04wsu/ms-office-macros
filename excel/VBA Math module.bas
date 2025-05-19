Attribute VB_Name = "Math"

'Based on https://www.techonthenet.com/excel/formulas/round_vba.php
'@param Double pValue				the value to round
'@param Integer pDecimalPlaces		how many decimal places to round to (if a Double is passed, it will be bankers-rounded to an Integer)
'@param Boolean pSymmetricRounding	if True, -1.5 would be rounded to -2 (instead of -1)
'@return Variant					returns either a Double or an Error
Public Function StandardRound(pValue As Double, Optional pDecimalPlaces As Integer = 0, Optional pSymmetricRounding As Boolean = False) As Variant
	
	Dim LValue As String
	Dim LPos As Integer
	Dim LNumDecimals As Long
	Dim LDecimalSymbol As String
	Dim QValue As Double
	
	' Return an error if the decimal places provided is negative
	If pDecimalPlaces < 0 Then
		StandardRound = CVErr(2001) '#VALUE!
		Exit Function
	End If
	
	' If your country uses a different symbol than the "." to denote a decimal
	' then change the following LDecimalSymbol variable to that character
	LDecimalSymbol = "."
	
	' Determine the number of decimal places in the value provided using
	' the length of the value and the position of the decimal symbol
	LValue = CStr(pValue)
	LPos = InStr(LValue, LDecimalSymbol)
	LNumDecimals = Len(LValue) - LPos
	
	' Round if the value provided has decimals and the number of decimals
	' is greater than the number of decimal places we are rounding to
	If (LPos > 0) And (LNumDecimals > 0) And (LNumDecimals > pDecimalPlaces) Then
		
		' Calculate the factor to add
		QValue = (1 / (10 ^ (LNumDecimals + 1)))
		
		If pSymmetricRounding Then
			' Symmetric rounding is commonly desired so if the value is
			' negative, make the factor negative
			' (Skipping the following 3 lines results in "Round Up" rounding)
			If (pValue < 0) Then
				QValue = -QValue
			End If
		End If
		
		' Add a 1 to the end of the value (For example, if pValue is 12.65
		' then we will use 12.651 when rounding)
		StandardRound = Round(pValue + QValue, pDecimalPlaces)
		
	' Otherwise return the original value
	Else
		StandardRound = pValue
	End If
	
End Function
