' Determine the split sequence for AllValues

Dim values : values = Split(AllValues, ";")
Dim sSplitting : sSplitting = "1"
Dim bDoSplit

For i = LBound(values) +1 To UBound(values)
	bDoSplit = False
	' compare current value with the previous value
	If Len(values(i)) > 0 Then
		If values(i) = values(i -1) Then
			bDoSplit = False
		Else
			bDoSplit = True
		End If
	End If
	' create the split sequence
	If Len(sSplitting) > 0 Then
		If bDoSplit Then
			sSplitting = sSplitting & ";"
		Else
			sSplitting = sSplitting & ","
		End If
	End If
	sSplitting = sSplitting & CStr(i - LBound(values) +1)
Next

RRV = "OutputPages=" & sSplitting
