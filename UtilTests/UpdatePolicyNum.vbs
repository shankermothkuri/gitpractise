Set Arg = WScript.Arguments
	Msgbox "Test for Jenkins"
'UpdatePolicyNumber Arg(0) Arg(1) Arg(2)
Sub UpdatePolicyNumber(strFilePath,strTCID,strPolicyNum)			
    On Error resume next
		
	Set oExcel = CreateObject("Excel.Application")
	oExcel.DisplayAlerts = False
	Set oWB = oExcel.Workbooks.Open(strFilePath)
	
	Set oWSheet=oWB.Sheets("MasterControl")
	
	intRows=oWSheet.UsedRange.Rows.Count
	
	For i=2 to intRows
		If(Trim(oWSheet.Cells(i,1))=Trim(strTCID)) Then
			If (Len(Trim(strPolicyNum))=0) Then
				oWSheet.Cells(i,3).Value="Null"
			Else
				oWSheet.Cells(i,3).Value=strPolicyNum
			End If
			Exit For
		End if
	Next
	
	
	oWB.Save
	oExcel.Quit
	set oExcel = Nothing
	
End Sub
