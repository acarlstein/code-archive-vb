'Opening a new sheet:
Public Sub OpenNewSheet(Optional sheetName As String = vbNullString)
	Dim ws As Worksheet
	Set ws = Sheets.Add
	If (sheetName <> vbNullString) Then
		ws.Name = sheetName
	End If
	Cells(1, 1).Select
End Sub

'Knowing if a string is alphanumeric:
Function isAlphanumeric(str As String) As Boolean
	Dim i As Integer
	isAlphanumeric = True
	For i = 1 To Len(Trim(str))
		Select Case Mid$(Trim(str), i, 1)
			Case "A" To "Z", "a" To "z", "0" To "9"
			Case Else
				isAlphanumeric = False
				Exit For
		End Select
	Next i
End Function

'Display references + Remove missing references + Add references by GUID:
Private Const REFERENCE_ALREADY_IN_USE = 32813
Private Const REFERENCE_ADDED_SUCCESSFULLY = vbNullString
Private Const Reference_Word As String = "{00020905-0000-0000-C000-000000000046}"
Private Const Reference_MSComCtl2 As String = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}"
Private Const Reference_MSForms As String = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
Private Const Reference_Office As String = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
Private Const Reference_stdole As String = "{00020430-0000-0000-C000-000000000046}"
Private Const Reference_Excel As String = "{00020813-0000-0000-C000-000000000046}"
Private Const Reference_VBA As String = "{000204EF-0000-0000-C000-000000000046}"
Private theRef As Variant

'
' AddReferences Macro
'
Sub AddReferences()
	addReferenceByGUID (Reference_Word)
	addReferenceByGUID (Reference_MSComCtl2)
	addReferenceByGUID (Reference_MSForms)
	addReferenceByGUID (Reference_Office)
	addReferenceByGUID (Reference_stdole)
	addReferenceByGUID (Reference_Excel)
	addReferenceByGUID (Reference_VBA)
End Sub

Private Function addReferenceByGUID(strGUID As String)

	On Error Resume Next

	 removeMissingReferences

	'Clear any errors so that error trapping for GUID additions can be evaluated
	Err.Clear

	'Add the reference
	ThisWorkbook.VBProject.References.AddFromGuid _
	GUID:=strGUID, Major:=1, Minor:=0

	Select Case Err.Number
	Case Is = REFERENCE_ALREADY_IN_USE
		 Debug.Print strGUID & " Reference Already in use"
	Case Is = REFERENCE_ADDED_SUCCESSFULLY
		  Debug.Print strGUID & " Reference Added"
	Case Else
		MsgBox "A problem was encountered trying to" & vbNewLine _
		& "add or remove a reference (" & strGUID & ") in this file" & vbNewLine & "Please check the " _
		& "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
	End Select
	On Error GoTo 0

End Function

Private Sub removeMissingReferences()
	Dim i As Long
	For i = ThisWorkbook.VBProject.References.count To 1 Step -1
		Set theRef = ThisWorkbook.VBProject.References.item(i)
		If theRef.isbroken = True Then
			ThisWorkbook.VBProject.References.Remove theRef
		End If
	Next i
End Sub

Public Sub displayReferencesInUse()
	Dim theRef As Variant, i As Long
	For i = ThisWorkbook.VBProject.References.count To 1 Step -1
		Set theRef = ThisWorkbook.VBProject.References.item(i)
		Debug.Print "Reference:" & theRef.Name & " " & theRef.GUID
	Next i
End Sub

'Fill ComboBox (or ListBox) with sheet names + Add string item into ComboBox (or ListBox):
Public Sub fillComboBoxWithSheetNames(cmbBox As Variant, defaultItem As String)
	cmbBox.Clear
	Dim i
	For i = 1 To Sheets.count
		addItemIntoComboBox cmbBox, Sheets(i).Name, defaultItem
	Next i
End Sub

Public Sub addItemIntoComboBox(cmbBox As Variant, item As String, Optional defaultItem As String)
	cmbBox.AddItem item
	If defaultItem <> vbNullString Then
		If InStr(item, defaultItem) <> 0 Then
			cmbBox.Value = item
		End If
	End If
End Sub

'Add Columns:
Public Sub addColumns(sheetName As String, columnNames As String)

	Dim ColArray() As String
	Dim oWS As Worksheet
	Dim lastColumn As Long
	Dim j As Long

	On Error GoTo Disp_Error

	Set oWS = Sheets(sheetName)

	ColArray = Split(columnNames, ",")
	For j = LBound(ColArray) To UBound(ColArray)
		lastColumn = getLastColumn(oWS) + 1
		oWS.Columns(lastColumn).Insert
		Cells(1, lastColumn) = Trim(ColArray(j))
	Next

Disp_Error:
	If Err <> 0 Then
		MsgBox Err.Number & " - " & Err.Description, vbExclamation, "Error Adding Column"
		Resume Next
	End If

End Sub

Get last column + Get last row:

Public Function getLastColumn(oWS As Worksheet) As Long
	getLastColumn = oWS.Cells(1, Columns.count).End(xlToLeft).Column
End Function

Public Function getLastRow(oWS As Worksheet) As Long
	getLastRow = oWS.Cells(Rows.count, 1).End(xlUp).Row
End Function

'Find out is column exist:
Public Function doColumnExist(sheetName As String, searchColumnName As String)
	Dim oWS As Worksheet
	Set oWS = Worksheets(sheetName)
	Dim columnNamesRange As range
	Set columnNamesRange = oWS.range(oWS.Cells(1, 1), oWS.Cells(1, getLastColumn(oWS)))
	Dim i
	For i = 1 To getLastColumn(oWS)
		If searchColumnName = columnNamesRange.Columns(i).Text Then
			doColumnExist = True
			Exit Function
		End If
	Next i
	doColumnExist = False
End Function