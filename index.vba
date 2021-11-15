Private Sub UserForm_Initialize()
txtId.SelStart = 0
 
With cmbSchool
.AddItem "Ndengetho High School"
.AddItem "Umlazi Comtech"
.AddItem "Lamontville High School"
.AddItem "Thokoza Mganga High School"
.AddItem "Qashana High School"
.AddItem "Ntee High School"
.AddItem "Amangwane High School"
.AddItem "Ohoye High School"
End With

With cmbGrade
.AddItem "8"
.AddItem "9"
.AddItem "10"
.AddItem "11"
End With

With cmbGender
.AddItem "Male"
.AddItem "Female"
End With

With cmbNoEnrollments
.AddItem "1"
.AddItem "2"
.AddItem "3"
.AddItem "4"
End With

With cmbYOFE
.AddItem "2015"
.AddItem "2016"
.AddItem "2017"
.AddItem "2018"
.AddItem "2019"
.AddItem "2020"
.AddItem "2021"
.AddItem "2022"
.AddItem "2023"
End With

With cmbPYOFE
.AddItem "2015"
.AddItem "2016"
.AddItem "2017"
.AddItem "2018"
.AddItem "2019"
.AddItem "2020"
.AddItem "2021"
.AddItem "2022"
.AddItem "2023"
End With

With cmbSocial
.AddItem "FACEBOOK"
.AddItem "LINKEDLN"
.AddItem "TWITTER"
.AddItem "INSTAGRAM"
.AddItem "NONE"
End With

Call Reset
End Sub


