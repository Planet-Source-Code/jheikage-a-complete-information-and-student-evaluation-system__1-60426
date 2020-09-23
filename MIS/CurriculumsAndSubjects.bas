Attribute VB_Name = "CurriculumsAndSubjects"
'Description    :   This functions are helper functions.
'CurriculaHere

Function SetSelectedCur(Sql As String)

On Error GoTo Erb
    With DE.rsCurricula
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Sql, DE.Con
    End With
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    If DE.rsCurricula.State = 1 Then DE.rsCurricula.CancelUpdate Else DE.rsCurricula.Close
    SetSelectedCur "select * from curriculum where sy = '9999-9999'"
End Function

Function LoadCurtoLV(lv As ListView)
'clear the items
lv.ListItems.Clear
With DE.rsCurricula
    If .RecordCount > 0 Then
        Do Until .EOF
            lv.ListItems.Add .AbsolutePosition, , .Fields("YR").Value
            lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SC").Value
            If IsNull(.Fields("Description").Value) Then
                lv.ListItems(.AbsolutePosition).SubItems(2) = ""
            Else
                lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Description").Value
            End If
            lv.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Unts").Value
            lv.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Prerequisites").Value
            .MoveNext
        Loop
    End If
End With
End Function

Function loadSLToListLv(lv As ListView)
lv.ListItems.Clear
With DE.rsCurricula
    If .RecordCount > 0 Then
        Do Until .EOF
            
            lv.ListItems.Add .AbsolutePosition, , .Fields("SC").Value
            If IsNull(.Fields("Description").Value) Then
                lv.ListItems(.AbsolutePosition).SubItems(1) = ""
            Else
                lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("Description").Value
            End If
            lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Unts").Value
            .MoveNext
        Loop
    End If
End With
    
End Function

'End of CurriculaHere

'begin Subjects of student

Function SetSelectedStudSub(Sql As String)

On Error GoTo Erb
    With DE.rsSubjectsEnrolled
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Sql, DE.Con
    End With
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    If DE.rsSubjectsEnrolled.State = 1 Then DE.rsSubjectsEnrolled.CancelUpdate Else DE.rsSubjectsEnrolled.Close
    SetSelectedStudSub "select * from subjectsEnrolled where sy = '9999-9999'"
End Function

Function loadSLtoLv(lv As ListView)
lv.ListItems.Clear
With DE.rsSubjectsEnrolled
    If .RecordCount > 0 Then
        Do Until .EOF
            lv.ListItems.Add .AbsolutePosition, , .Fields("YR").Value
            lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("SC").Value
            lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Description").Value
            lv.ListItems(.AbsolutePosition).SubItems(3) = .Fields("Units").Value
            lv.ListItems(.AbsolutePosition).SubItems(4) = .Fields("Grade").Value
            lv.ListItems(.AbsolutePosition).SubItems(5) = .Fields("Remarks").Value
            .MoveNext
        Loop
    End If
End With
FrmStudentSubjectList.GroupSYSem
End Function
'end Subjects of student

'NEW! Update for SUbjects Offered.

Function SetSelectedSubOff(Sql As String)

On Error GoTo Erb
    With DE.rsSubjectOffered
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Sql, DE.Con
    End With
Exit Function
Erb:
    MsgBox Err.Description, vbCritical, "Error"
    If DE.rsSubjectOffered.State = 1 Then DE.rsSubjectOffered.CancelUpdate Else DE.rsSubjectOffered.Close
    SetSelectedSubOff "select * from subjects_Offered where sy = '9999-9999'"
End Function

Function loadSOToListLv(lv As ListView)
lv.ListItems.Clear
With DE.rsSubjectOffered
    If .RecordCount > 0 Then
        Do Until .EOF
            
            lv.ListItems.Add .AbsolutePosition, , .Fields("SC").Value
            If IsNull(.Fields("Description").Value) Then
                lv.ListItems(.AbsolutePosition).SubItems(1) = ""
            Else
                lv.ListItems(.AbsolutePosition).SubItems(1) = .Fields("Description").Value
            End If
            lv.ListItems(.AbsolutePosition).SubItems(2) = .Fields("Unts").Value
            .MoveNext
        Loop
    End If
End With
    
End Function
