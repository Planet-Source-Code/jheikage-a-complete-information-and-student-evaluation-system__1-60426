Attribute VB_Name = "Mainmod"
'Description    :   This functions are helper functions for connecting, record parsing
'                   searching etc. just take a look at them
Public IsTrue As Boolean 'for confirmation
Public MyLoginPass As String
Sub Main()
If App.PrevInstance Then
    MsgBox "Another instance of this program is running.", vbInformation, "MIS"
    End
End If
  Frmlogin.Show
End Sub

Public Function connect(user As String, password As String, server As String) As Boolean
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=Pc4;Initial Catalog=CSUMIS;Data Source=WINDOWS-7EC8A49
  On Error GoTo Erb
  With DE.Con
  If .State <> 0 Then .Close
    .ConnectionString = "Provider=MSDATASHAPE.1;data Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & password & ";Initial Catalog=CSUMIS;Data Source=" & server
    .CursorLocation = adUseClient
    .Open
  End With
  connect = True
  SetRs "Select * from names"
  Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error:" & Err.Number
    DE.Con.Close
    connect = False
End Function

Public Function SetRs(Sql As String)
Static x As Integer
x = x + 1
On Error GoTo Erb
'Dim rs As New Recordset
With DE.rsPIS
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open Sql, DE.Con, adOpenDynamic, adLockOptimistic
End With
x = 0
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error:" & Err.Number
    If FrmMainPIS.isadd = True Then DE.rsPIS.CancelBatch adAffectAllChapters
    DE.rsPIS.CancelUpdate
    If x > 3 Then Exit Function
    SetRs "Select * from names where idnum='XXX'"
End Function

Public Function LoadLV(lv As ListView, Sql As String)
On Error GoTo Erb
    Dim rs As New Recordset
    Set rs = Nothing
    'clear the lv
    lv.ListItems.Clear
    'set lists
    SetRs "Select * From Names"
    With rs
        .CursorLocation = adUseClient
        .Open Sql, DE.Con, adOpenDynamic, adLockOptimistic
        If .RecordCount > 0 Then
            Do Until .EOF
                If IsNull(.Fields("IDNUM").Value) Then
                    lv.ListItems.Add .AbsolutePosition, , "", 2, 2
                Else
                    lv.ListItems.Add .AbsolutePosition, , .Fields("IDNUM").Value, 2, 2
                End If
                lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("Lnam").Value
                lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Fnam").Value
                lv.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Mnam").Value
                'Updated code march 17, 2005
                If IsNull(.Fields("Curriculum").Value) Then     'use the tool tip text to store the curriculum
                    lv.ListItems(.AbsolutePosition).ToolTipText = ""
                Else
                    lv.ListItems(.AbsolutePosition).ToolTipText = .Fields("Curriculum").Value
                End If
                If IsNull(.Fields("course").Value) Then         'use tag to store course
                    lv.ListItems(.AbsolutePosition).Tag = ""
                Else
                    lv.ListItems(.AbsolutePosition).Tag = .Fields("course").Value
                End If
                'end of update
                .MoveNext
            Loop
        End If
        FrmMainPIS.LblDB.Caption = "Total Record Count: " & DE.rsPIS.RecordCount
        FrmMainPIS.LblQinfo.Caption = "Total Query Result: " & FrmMainPIS.LVStudents.ListItems.Count
    End With
    Exit Function
Erb:
    Set rs = Nothing
    MsgBox Err.Description, vbCritical, "Error:" & Err.Number
End Function

Public Function ReBind()
With FrmPIS
Set .txtrlntship.DataSource = DE
    .txtrlntship.DataMember = "PIS"
    .txtrlntship.DataField = "rlntship"
Set .txtIdnum.DataSource = DE
    .txtIdnum.DataMember = "PIS"
    .txtIdnum.DataField = "idnum"
Set .txtaddr.DataSource = DE
    .txtaddr.DataMember = "PIS"
    .txtaddr.DataField = "addr"
Set .txtAge.DataSource = DE
    .txtAge.DataMember = "PIS"
    .txtAge.DataField = "age"
Set .txtBday.DataSource = DE
    .txtBday.DataMember = "PIS"
    .txtBday.DataField = "bday"
Set .TxtBoarding.DataSource = DE
    .TxtBoarding.DataMember = "PIS"
    .TxtBoarding.DataField = "boarding"
Set .txtcoll.DataSource = DE
    .txtcoll.DataMember = "PIS"
    .txtcoll.DataField = "coll"
Set .TxtCollege.DataSource = DE
    .TxtCollege.DataMember = "PIS"
    .TxtCollege.DataField = "college"
Set .txtCourse.DataSource = DE
    .txtCourse.DataMember = "PIS"
    .txtCourse.DataField = "course"
Set .txtCschladd.DataSource = DE
    .txtCschladd.DataMember = "PIS"
    .txtCschladd.DataField = "Cschladd"
Set .txtfnam.DataSource = DE
    .txtfnam.DataMember = "PIS"
    .txtfnam.DataField = "fnam"
Set .TxtFoccup.DataSource = DE
    .TxtFoccup.DataMember = "PIS"
    .TxtFoccup.DataField = "foccup"
Set .txtGuardian.DataSource = DE
    .txtGuardian.DataMember = "PIS"
    .txtGuardian.DataField = "guardian"
Set .txtHAddress.DataSource = DE
    .txtHAddress.DataMember = "PIS"
    .txtHAddress.DataField = "haddress"
Set .txtIdnum.DataSource = DE
    .txtIdnum.DataMember = "PIS"
    .txtIdnum.DataField = "idnum"
Set .txtlnam.DataSource = DE
    .txtlnam.DataMember = "PIS"
    .txtlnam.DataField = "lnam"
Set .txtMajor.DataSource = DE
    .txtMajor.DataMember = "PIS"
    .txtMajor.DataField = "Major"
Set .txtMinor.DataSource = DE
    .txtMinor.DataMember = "PIS"
    .txtMinor.DataField = "minor"
Set .txtmnam.DataSource = DE
    .txtmnam.DataMember = "PIS"
    .txtmnam.DataField = "mnam"
Set .TxtMoccup.DataSource = DE
    .TxtMoccup.DataMember = "PIS"
    .TxtMoccup.DataField = "moccup"
Set .txtNationality.DataSource = DE
    .txtNationality.DataMember = "PIS"
    .txtNationality.DataField = "Nationality"

Set .txtnmefther.DataSource = DE
    .txtnmefther.DataMember = "PIS"
    .txtnmefther.DataField = "nmefther"
Set .txtnmelanlord.DataSource = DE
    .txtnmelanlord.DataMember = "PIS"
    .txtnmelanlord.DataField = "nmelanlord"
Set .txtnmemther.DataSource = DE
    .txtnmemther.DataMember = "PIS"
    .txtnmemther.DataField = "nmemther"
Set .TxtOccupation.DataSource = DE
    .TxtOccupation.DataMember = "PIS"
    .TxtOccupation.DataField = "Occupation"
Set .txtPlacebirth.DataSource = DE
    .txtPlacebirth.DataMember = "PIS"
    .txtPlacebirth.DataField = "Placebirth"
Set .txtRaddress.DataSource = DE
    .txtRaddress.DataMember = "PIS"
    .txtRaddress.DataField = "Raddress"
Set .txtrel.DataSource = DE
    .txtrel.DataMember = "PIS"
    .txtrel.DataField = "rel"
Set .txtrlntship.DataSource = DE
    .txtrlntship.DataMember = "PIS"
    .txtrlntship.DataField = "rlntship"
Set .txtscondary.DataSource = DE
    .txtscondary.DataMember = "PIS"
    .txtscondary.DataField = "scondary"
Set .TxtSex.DataSource = DE
    .TxtSex.DataMember = "PIS"
    .TxtSex.DataField = "sex"
Set .txtSschladd.DataSource = DE
    .txtSschladd.DataMember = "PIS"
    .txtSschladd.DataField = "sschladd"
Set .txtyrgrad.DataSource = DE
    .txtyrgrad.DataMember = "PIS"
    .txtyrgrad.DataField = "yrgrad"
Set .txtyrgrduated.DataSource = DE
    .txtyrgrduated.DataMember = "PIS"
    .txtyrgrduated.DataField = "yrgrduated"
Set .TxtCS.DataSource = DE
    .TxtCS.DataMember = "PIS"
    .TxtCS.DataField = "CS"
Set .TxtIncome.DataSource = DE
    .TxtIncome.DataMember = "PIS"
    .TxtIncome.DataField = "Income"
Set .txtSY.DataSource = DE
    .txtSY.DataMember = "PIS"
    .txtSY.DataField = "Curriculum"
End With
End Function

Function ttx(obj As Object)
    obj.Text = LTrim(RTrim(obj.Text))
End Function

Public Function Retrim()
With FrmPIS
    ttx .txtIdnum
    ttx .txtaddr
    ttx .txtAge
    ttx .txtBday
    ttx .TxtBoarding
    ttx .txtcoll
    ttx .TxtCollege
    ttx .txtCourse
    ttx .txtCschladd
    ttx .txtfnam
    ttx .TxtFoccup
    ttx .txtGuardian
    ttx .txtHAddress
    ttx .txtIdnum
    ttx .txtlnam
    ttx .txtMajor
    ttx .txtMinor
    ttx .txtmnam
    ttx .TxtMoccup
    ttx .txtNationality
    ttx .txtnmefther
    ttx .txtnmelanlord
    ttx .txtnmemther
    ttx .TxtOccupation
    ttx .txtPlacebirth
    ttx .txtRaddress
    ttx .txtrel
    ttx .txtrlntship
    ttx .txtscondary
    ttx .TxtSex
    ttx .txtSschladd
    ttx .txtyrgrad
    ttx .txtyrgrduated
    ttx .TxtCS
    ttx .txtrlntship
    ttx .TxtIncome
    ttx .txtSY
End With
End Function

Public Function clearText()
With FrmPIS
    .txtrlntship.Text = ""
    .txtIdnum.Text = ""
    .txtaddr.Text = ""
    .txtAge.Text = ""
    .txtBday.Text = ""
    .TxtBoarding.Text = ""
    .txtcoll.Text = ""
    .TxtCollege.Text = ""
    .txtCourse.Text = ""
    .txtCschladd.Text = ""
    .txtfnam.Text = ""
    .TxtFoccup.Text = ""
    .txtGuardian.Text = ""
    .txtHAddress.Text = ""
    .txtIdnum.Text = ""
    .txtlnam.Text = ""
    .txtMajor.Text = ""
    .txtMinor.Text = ""
    .txtmnam.Text = ""
    .TxtMoccup.Text = ""
    .txtNationality.Text = ""
    .txtnmefther.Text = ""
    .txtnmelanlord.Text = ""
    .txtnmemther.Text = ""
    .TxtOccupation.Text = ""
    .txtPlacebirth.Text = ""
    .txtRaddress.Text = ""
    .txtrel.Text = ""
    .txtrlntship.Text = ""
    .txtscondary.Text = ""
    .TxtSex.Text = ""
    .txtSschladd.Text = ""
    .txtyrgrad.Text = ""
    .txtyrgrduated.Text = ""
    .TxtCS.Text = ""
    .TxtIncome.Text = ""
    .txtSY.Text = ""
End With
End Function
