Sub CheckSplit_OnLoad
	EKOManager.StatusMessage ("CheckSplit_Onload")
	
	' I'm going to be getting 2 documents at this stage, (1) original TIFF or PDF, and (2) the CSV
	' I have to find the CSV, find the split sequence inside it, then delete the CSV from my Knowledge Object (KO)
	' A KO contains 0 or more KD files (Knowledge Documents)
	
	EKOManager.StatusMessage ("KnowledgeObject.DocumentCount = " & KnowledgeObject.DocumentCount)
	
	For i = 0 To KnowledgeObject.DocumentCount
		Set KDocument = KnowledgeObject.GetDocument(i)
		FilePath = KDocument.FilePath
		FileExt = KDocument.GetFileExtension
		If (Ucase(FileExt) = "CSV") Then
			EKOManager.StatusMessage ("Found a CSV File")
			' pass the CSV to GetSplitSequence, get the result, create a new RRT in AS, then delete the CSV file
			splitSequence = GetSplitSequence("Check Number: [0-9]{6,}", FilePath)
			EKOManager.StatusMessage ("splitSequence = " & splitSequence)
			' Create new RRT
			Set Topic = KnowledgeContent.GetTopicInterface
			Topic.Replace "~USR::SplitSequence~", splitSequence
			
			' PB : 02MAY2013, potential bug in AS SDK on removing document, requiring index +1 as opposed to just index
			KnowledgeObject.RemoveDocument((i+1))
			EKOManager.StatusMessage ("Split Complete")
		End If
	Next
	
End Sub

Sub CheckSplit_OnUnload

End Sub

' Returns split sequence of pages based on changed values
Function GetSplitSequence(regexProfile, filename)
	Dim splitSequence : splitSequence = ""
	
	Dim arrLines : arrLines = AutoStoreLibrary_ReadTextFile_Unicode(filename)
	' match a regex against Check Number: [0-9]{6,}
	
	Set re = New RegExp
	re.Pattern = regexProfile
	
	Dim pageCounter : pageCounter = -1
	Dim lastCheckNumber : lastCheckNumer = ""
	For Each line In arrLines
		pageCounter = pageCounter + 1
		Set matches = re.Execute(line)
		If (matches.count > 0) Then
			If (matches(0) <> lastCheckNumber) Then
				'	msgbox pageCounter & ":" & matches(0)
				If (Len(splitSequence) > 0) Then
					splitSequence = splitSequence & ","
				End If
				splitSequence = splitSequence & (pageCounter -1)
			End If
			
			lastCheckNumber = matches(0)
		End If
	Next
	
	GetSplitSequence = splitSequence
End Function

Function AutoStoreLibrary_ReadTextFile_Unicode(file)
	Const ForReading = 1
	
	Dim arrFileLines()
	
	Dim ReadSuccess : ReadSuccess = False
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(file) Then
		Set objFile = objFSO.OpenTextFile(file, ForReading, True, -1)
		
		i = 0
		Do Until objFile.AtEndOfStream
			ReDim Preserve arrFileLines(i)
			arrFileLines(i) = objFile.ReadLine
			i = i + 1
		Loop
		
		objFile.Close
		ReadSuccess = True
	End If
	
	AutoStoreLibrary_ReadTextFile_Unicode = arrFileLines
End Function
