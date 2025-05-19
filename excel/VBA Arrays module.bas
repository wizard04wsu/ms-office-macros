Attribute VB_Name = "Arrays"

'Returns the number of items in the array.
Public Function GetLength(arr As Variant) As Long
	
	If Not IsArray(arr) Then
		Err.Raise 13, "GetLength", "Type mismatch: 'arr' must be an array"
	End If
	
	Dim upper As Long, lower As Long
	
	On Error Resume Next
	upper = UBound(arr)
	lower = LBound(arr)
	If Err.Number = 9 Then
		'error: Subscript out of range
		GoTo GetLength_empty_array
	ElseIf Err.Number <> 0 Then
		Dim savedErr: savedErr = Array(Err.Number, Err.Source, Err.Description)
		On Error GoTo 0
		Err.Raise savedErr(0), savedErr(1), savedErr(2)
	End If
	On Error GoTo 0
	
	If upper < lower Then
		'note: the array was declared and initialized like `Dim arr(): arr = Array()`
		GoTo GetLength_empty_array
	End If
	
	GetLength = UBound(arr) - LBound(arr) + 1
	
	Exit Function
	
GetLength_empty_array:
	'the array is empty
	
	GetLength = 0
	
End Function

'Returns the array's highest index. If the array is empty, Null is returned.
Public Function GetUBound(arr As Variant) As Variant
	
	If Not IsArray(arr) Then
		Err.Raise 13, "GetUBound", "Type mismatch: 'arr' must be an array"
	End If
	
	If GetLength(arr) = 0 Then
		'the array is empty
		
		GetUBound = Null
	Else
		GetUBound = UBound(arr)
	End If
	
End Function

'Returns the array's lowest index. If the array is empty, Null is returned.
Public Function GetLBound(arr As Variant) As Variant
	
	If Not IsArray(arr) Then
		Err.Raise 13, "GetLBound", "Type mismatch: 'arr' must be an array"
	End If
	
	If GetLength(arr) = 0 Then
		'the array is empty
		
		GetLBound = Null
	Else
		GetLBound = LBound(arr)
	End If
	
End Function

'Adds an item to the end of a dynamic array.
Public Sub Push(arr As Variant, item As Variant)
	
	If Not IsArray(arr) Then
		Err.Raise 13, "Push", "Type mismatch: 'arr' must be an array"
	End If
	
	Dim lower As Variant: lower = GetLBound(arr)
	Dim upper As Variant: upper = GetUBound(arr)
	
	If IsNull(lower) Then
		'the array is empty
		
		'add the new value
		arr = Array(item)
	Else
		'resize the array
		On Error Resume Next
		ReDim Preserve arr(lower To (upper + 1))
		If Err.Number = 10 Then
			'error: This array is fixed or temporarily locked
			Err.Description = "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
		If Err.Number <> 0 Then
			Dim savedErr: savedErr = Array(Err.Number, Err.Description)
			On Error GoTo 0
			Err.Raise savedErr(0), "Push", savedErr(1)
		End If
		On Error GoTo 0
		
		'add the new value
		assignAtIndex arr, upper + 1, item
	End If
	
End Sub

'Removes and returns the last item in a dynamic array, or returns Null if the array is already empty.
Public Function Pop(arr As Variant) As Variant
	
	If Not IsArray(arr) Then
		Err.Raise 13, "Pop", "Type mismatch: 'arr' must be an array"
	End If
	
	Dim length As Long: length = GetLength(arr)
	Dim lower As Variant: lower = GetLBound(arr)
	Dim upper As Variant: upper = GetUBound(arr)
	
	If length = 0 Then
		'the array is empty
		
		Pop = Null
		Exit Function
	End If
	
	If IsObject(arr(upper)) Then
		Set Pop = arr(upper)
	Else
		Pop = arr(upper)
	End If
	
	If length = 1 Then
		'that was the only item left
		
		'empty the array
		Erase arr
		
		If GetLength(arr) > 0 Then
			'the array is fixed-length
			Err.Raise 10, "Pop", "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
	Else
		'resize the array
		On Error Resume Next
		ReDim Preserve arr(lower To (upper - 1))
		If Err.Number = 10 Then
			'error: This array is fixed or temporarily locked
			Err.Description = "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
		If Err.Number <> 0 Then
			Dim savedErr: savedErr = Array(Err.Number, Err.Description)
			On Error GoTo 0
			Err.Raise savedErr(0), "Pop", savedErr(1)
		End If
		On Error GoTo 0
	End If
	
End Function

'Adds an item to the beginning of a dynamic array.
Public Sub Unshift(arr As Variant, item As Variant)
	
	If Not IsArray(arr) Then
		Err.Raise 13, "Unshift", "Type mismatch: 'arr' must be an array"
	End If
	
	Dim lower As Variant: lower = GetLBound(arr)
	Dim upper As Variant: upper = GetUBound(arr)
	Dim i As Long
	
	If IsNull(lower) Or IsNull(upper) Then
		'the array is empty
		
		'add the new value
		arr = Array(item)
	Else
		'resize the array
		ReDim Preserve arr(lower To (upper + 1))
		If Err.Number = 10 Then
			'error: This array is fixed or temporarily locked
			Err.Description = "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
		If Err.Number <> 0 Then
			Dim savedErr: savedErr = Array(Err.Number, Err.Description)
			On Error GoTo 0
			Err.Raise savedErr(0), "Unshift", savedErr(1)
		End If
		On Error GoTo 0
		
		'move everything toward the upper bound by one position
		For i = upper To lower Step -1
			assignAtIndex arr, i + 1, arr(i)
		Next
		
		'add the new value
		assignAtIndex arr, lower, item
	End If
	
End Sub

'Removes and returns the first item in a dynamic array, or returns Null if the array is already empty.
Public Function Shift(arr As Variant) As Variant
	
	If Not IsArray(arr) Then
		Err.Raise 13, "Shift", "Type mismatch: 'arr' must be an array"
	End If
	
	Dim length As Long: length = GetLength(arr)
	Dim lower As Variant: lower = GetLBound(arr)
	Dim upper As Variant: upper = GetUBound(arr)
	Dim i As Long
	
	If length = 0 Then
		'the array is empty
		
		Shift = Null
		Exit Function
	End If
	
	If IsObject(arr(lower)) Then
		Set Shift = arr(lower)
	Else
		Shift = arr(lower)
	End If
	
	If length = 1 Then
		'that was the only item left
		
		'empty the array
		Erase arr
		
		If GetLength(arr) > 0 Then
			'the array is fixed-length
			Err.Raise 10, "Shift", "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
	Else
		'move everything toward the lower bound by one position
		For i = lower + 1 To upper
			assignAtIndex arr, i - 1, arr(i)
		Next
		
		'resize the array
		ReDim Preserve arr(lower To (upper - 1))
		If Err.Number = 10 Then
			'error: This array is fixed or temporarily locked
			Err.Description = "This array is fixed or temporarily locked: 'arr' must be a dynamic array"
		End If
		If Err.Number <> 0 Then
			Dim savedErr: savedErr = Array(Err.Number, Err.Description)
			On Error GoTo 0
			Err.Raise savedErr(0), "Shift", savedErr(1)
		End If
		On Error GoTo 0
	End If
	
End Function

'Assigns a value to a position in an array.
Public Sub assignAtIndex(arr As Variant, index As Variant, item As Variant)
	
	If Not IsArray(arr) Then
		Err.Raise 13, "assignAtIndex", "Type mismatch: 'arr' must be an array"
	End If
	If index <> CLng(index) Then
		Err.Raise 13, "assignAtIndex", "Type mismatch: 'index' must be an integer"
	End If
	
	Dim length As Long: length = GetLength(arr)
	Dim lower As Variant: lower = GetLBound(arr)
	Dim upper As Variant: upper = GetUBound(arr)
	Dim invalidIndex As Boolean
	
	invalidIndex = IsNull(lower) Or IsNull(upper)   'the array is empty
	If Not invalidIndex Then invalidIndex = (index < lower Or index > upper)	'the index is out of range
	
	If invalidIndex Then
		Err.Raise 9, "assignAtIndex", "Subscript out of range: 'index' is outside of the array's lower and upper bounds"
	End If
	
	If IsObject(item) Then
		Set arr(index) = item
	Else
		arr(index) = item
	End If
	
End Sub
