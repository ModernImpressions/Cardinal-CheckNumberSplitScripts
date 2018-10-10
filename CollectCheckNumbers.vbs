' collect barcode values into a semicolon delimited string
' "1001;1001;1001;1002;1002;1003;etc"

Const REGEX_PATTERN = "^()"

'If Len(AllValues) > 0 Then
'	AllValues = AllValues & ";" & CurrentValue
'Else
'	AllValues = CurrentValue
'End If

If Len(AllValues) > 0 Then
	 'If IsValidRegExp(CurrentValue, REGEX_PATTERN) Then
		 'AllValues = AllValues & ";" & CurrentValue
	 'Else
		If CurrentValue = CurrentValue Or 0 Then
			AllValues = AllValues & ";" & CurrentValue
		End If	
	 'End If
Else
	 'If IsValidRegExp(CurrentValue, REGEX_PATTERN) Then
		AllValues = CurrentValue
	 'End If
End If

RRV = "AllValues=" & AllValues


' HELPER FUNCTIONS
' -----------------

'Test a Regular Expression pattern, Return True or False
Function IsValidRegExp(TestValue, TestExpression)
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Pattern = TestExpression
	objRegExp.Global = True
	If Not objRegExp.Test(TestValue) Then
		IsValidRegExp = False
	Else
		IsValidRegExp = True
	End If
End Function
