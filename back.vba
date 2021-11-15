Sub Reset()
Dim iRow As Long
iRow = [Counta(Database!A:A)]
With UserForm1
.txtName.Value = ""
.txtSurname.Value = ""
.cmbSchool.Value = ""
.cmbGrade.Value = ""
.cmbGender.Value = ""
.txtId.Value = ""
.txtCellphone.Value = ""
.OptionY.Value = ""
.cmbNoEnrollments.Value = ""
.cmbYOFE.Value = ""
.cmbPYOFE.Value = ""
.cmbSocial.Value = ""
.txtUsername.Value = ""


Call Add_SearchColumn

ThisWorkbook.Sheets("Database").AutoFilterMode = False
ThisWorkbook.Sheets("SearchData").AutoFilterMode = False
ThisWorkbook.Sheets("SearchData").Cells.Clear

.ListDatabase.ColumnCount = 17
.ListDatabase.ColumnHeads = True

.ListDatabase.ColumnWidths = "30,100,100,120,120,120,90,120,120,120,90,90,90,90,90,90,90"

If iRow > 1 Then
.ListDatabase.RowSource = "Database!A2:Q" & iRow
Else
.ListDatabase.RowSource = "Database!A2:Q2"
End If
End With
End Sub

Sub submit()
Dim sh As Worksheet
Dim iRow As Long

Set sh = ThisWorkbook.Sheets("Database")
If UserForm1.txtRowNumber.Value = "" Then
iRow = [Counta(Database!A:A)] + 1
Else
iRow = UserForm1.txtRowNumber.Value
End If

With sh
.Cells(iRow, 1) = iRow - 1
.Cells(iRow, 2) = UserForm1.txtName.Value
.Cells(iRow, 3) = UserForm1.txtSurname.Value
.Cells(iRow, 4) = UserForm1.cmbSchool.Value
.Cells(iRow, 5) = UserForm1.cmbGrade.Value
.Cells(iRow, 6) = UserForm1.cmbGender.Value
.Cells(iRow, 7) = UserForm1.txtId.Value
.Cells(iRow, 8) = UserForm1.txtCellphone.Value
.Cells(iRow, 9) = IIf(UserForm1.OptionY.Value = True, "Yes", "No")
.Cells(iRow, 10) = UserForm1.cmbNoEnrollments.Value
.Cells(iRow, 11) = Application.UserName
.Cells(iRow, 12) = UserForm1.cmbYOFE.Value
.Cells(iRow, 13) = [Text(now(), "DD-MM-YYYY HH:MM:SS")]
.Cells(iRow, 15) = UserForm1.cmbPYOFE.Value
.Cells(iRow, 16) = UserForm1.cmbSocial.Value
.Cells(iRow, 17) = UserForm1.txtUsername.Value

If ([Text(now(), "YYYY")]) - .Cells(iRow, 12) > 5 Then
.Cells(iRow, 14) = "Completed School"

Set sh = ThisWorkbook.Sheets("Completed")
If UserForm1.txtRowNumber.Value = "" Then
iRow = [Counta(Database!A:A)] + 1
Else
iRow = UserForm1.txtRowNumber.Value
End If

With sh
.Cells(iRow, 1) = iRow - 1
.Cells(iRow, 2) = UserForm1.txtName.Value
.Cells(iRow, 3) = UserForm1.txtSurname.Value
.Cells(iRow, 4) = UserForm1.cmbSchool.Value
.Cells(iRow, 5) = UserForm1.cmbGrade.Value
.Cells(iRow, 6) = UserForm1.cmbGender.Value
.Cells(iRow, 7) = UserForm1.txtId.Value
.Cells(iRow, 8) = UserForm1.txtCellphone.Value
.Cells(iRow, 9) = UserForm1.OptionY.Value
.Cells(iRow, 10) = UserForm1.cmbNoEnrollments.Value
.Cells(iRow, 11) = Application.UserName
.Cells(iRow, 12) = UserForm1.cmbYOFE.Value
.Cells(iRow, 13) = [Text(now(), "DD-MM-YYYY HH:MM:SS")]
.Cells(iRow, 15) = UserForm1.cmbPYOFE.Value
.Cells(iRow, 16) = UserForm1.cmbSocial.Value
.Cells(iRow, 17) = UserForm1.txtUsername.Value
End With
Else
.Cells(iRow, 14) = "NotFinished School"
End If
End With
End Sub

Sub show_Form()

UserForm1.Show

End Sub

Function Selected_List() As Long
Dim i As Long
Selected_List = 0
For i = 0 To UserForm1.ListDatabase.ListCount - 1
If UserForm1.ListDatabase.Selected(i) = True Then
Selected_List = i + 1
Exit For
End If

Next i

End Function

Sub Add_SearchColumn()

UserForm1.EnableEvents = False

With UserForm1.ComboBox4

.Clear
.AddItem "All"
.AddItem "Name"
.AddItem "Surname"
.AddItem "School"
.AddItem "Grade"
.AddItem "Gender"
.AddItem "Date Of Birth"
.AddItem "Cellphone"
.AddItem "Returning"
.AddItem "No. OF Enrollements"
.AddItem "Submitted By"
.AddItem "School YOFE"
.AddItem "Student Status"
.AddItem "Program YOFE"

.Value = "All"

End With

UserForm1.EnableEvents = True

UserForm1.TextBox6.Value = ""
UserForm1.TextBox6.Enabled = False
UserForm1.cmdSearch.Enabled = False

End Sub


Sub SearchData()

Application.ScreenUpdating = False

Dim shDatabase As Worksheet
Dim shSearchData As Worksheet

Dim iColumn As Integer
Dim iDatabaseRow As Long
Dim iSearchRow As Long

Dim sColumn As String
Dim sValue As String

Set shDatabase = ThisWorkbook.Sheets("Database")
Set shSearchData = ThisWorkbook.Sheets("SearchData")

iDatabaseRow = ThisWorkbook.Sheets("Database").Range("A" & Application.Rows.Count).End(xlUp).Row

sColumn = UserForm1.ComboBox4.Value
sValue = UserForm1.TextBox6.Value

iColumn = Application.WorksheetFunction.Match(sColumn, shDatabase.Range("A1:O1"), 0)

If shDatabase.FilterMode = True Then
shDatabase.AutoFilterMode = False
End If

If UserForm1.ComboBox4.Value = "Surname" Then
shDatabase.Range("A1:O" & iDatabaseRow).AutoFilter Field:=iColumn, Criteria1:=sValue

Else

shDatabase.Range("A1:O" & iDatabaseRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"

End If

If Application.WorksheetFunction.Subtotal(3, shDatabase.Range("C:C")) >= 2 Then
shSearchData.Cells.Clear
shDatabase.AutoFilter.Range.Copy shSearchData.Range("A1")
Application.CutCopyMode = False
iSearchRow = shSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
UserForm1.ListDatabase.ColumnCount = 12
UserForm1.ListDatabase.ColumnWidths = "30,100,100,100,100,100,90,100,100,100,100,100"
End If
If iSearchRow > 1 Then
UserForm1.ListDatabase.RowSource = "SearchData!A2:O" & iSearchRow
MsgBox "Some Records were Found"

Else
MsgBox "No Record Found"
End If
End Sub

Private Sub cmbReset_Click()
Dim msgValue As VbMsgBoxResult
msgValue = MsgBox("Do you want to Reset form?", vbYesNo + vbInformation, "Confirmation")
If msgValue = vbNo Then Exit Sub
Call Reset
End Sub
