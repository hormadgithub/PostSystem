Attribute VB_Name = "mdCenter"
    'ใช้เป็น Recordset ชั่วคราว
    Public rsTmp As ADODB.Recordset
    'ใช้แทนคำสั่ง SQL
    Public strCmdSQL  As String
    'ใช้แทน Control ใน Form
    Public Ctrl As Control
    
    'ใช้แทน Usr ที่ Logon เข้าระบบ
    Public UsrName As String * 3
    Public UsrDept As String * 3
    Public UsrLevel As Variant
    
    'ใช้เก็บค่าเมื่อ OPO มีการ Check ความถูกต้องแล้ว
    Public OPoCheck As Boolean
    
    'Procedure FindAllFields()
    'ใช้เก็บค่า Fields ทุก Fields ใน Record
    Public Field As Field
    Public Fld(0 To 50)

Public Sub AddList(tbName As String, fldName1 As String, ctrlList As Control, Condition As String)
'เป็น Procedure ที่ใช้ Add ค่า fldName1 ใน tbName เข้าไปใน ctrlList
'tbName = ชื่อ Table, fldName1 = Field ที่ต้องการดึง, ctrlList = Control ที่มี Method AddItem
        Set rsTmp = New ADODB.Recordset
        If Condition = "" Then
                strCmdSQL = "SELECT " & fldName1 & " FROM " & tbName
        Else
                strCmdSQL = "SELECT " & fldName1 & " FROM " & tbName & " " & Condition
        End If
        rsTmp.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        ctrlList.Clear
        With rsTmp
                If .BOF Or .EOF Then
                        Exit Sub
                End If
                .MoveFirst
                Do Until .EOF
                        ctrlList.AddItem .Fields(fldName1)
                        .MoveNext
                Loop
                .Close
        End With
        Set rsTmp = Nothing
End Sub

Public Sub Add2ListView(tbName As String, fldName1 As String, fldName2 As String, ctrlList As Control)
'เป็น Procedure ที่ใช้ Add ค่า fldName1 และ fldName2 หรือเพียง fldName1 อย่างเดียวใน tbName เข้าไปใน ctrlList
'tbName = ชื่อ Table, fldName1 = Field ที่ต้องการดึง, ctrlList = Control ที่มี Method AddItem
    Set rsTmp = New ADODB.Recordset
    strCmdSQL = "SELECT " & fldName1 & "," & fldName2 & " FROM " & tbName
    rsTmp.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
    ctrlList.Clear
    With rsTmp
        .MoveFirst
        Do Until .EOF
            ctrlList.AddItem .Fields(fldName1) & " - " & .Fields(fldName2)
            .MoveNext
        Loop
        .Close
    End With
    Set rsTmp = Nothing
End Sub

Public Sub Add_Record()
    With rsActive
        .AddNew
    End With
        
   MovetxtCaption = "กำลังทำการเพิ่มข้อมูลแฟ้ม " & tbActive & "..."

    With frmActive
        .lblMovetxt.Left = .Width
        .tmMovetxt.Enabled = True
        .lblMovetxt.Caption = MovetxtCaption   'เป็นตัวแปรใน Module Main
        .Caption = frmActive.Caption & "(" & mode & ")"
        .Show vbModal
    End With

End Sub

Public Sub Cancel_Record()
    With rsActive
        .CancelBatch
    End With
End Sub

Public Sub chkTypeOfControl()
    For Each Ctrl In frmActive.Controls
        If Trim$(Ctrl.Tag) <> "" Then
            Set Ctrl.DataSource = rsActive
            'MsgBox Ctrl.Name
            Ctrl.DataField = Ctrl.Tag
        End If
    Next
End Sub

Public Sub Edit_Record()
    On Error GoTo Err_Edit
    If rsActive.AbsolutePosition > 0 Then 'แก้ไขในกรณีที่ไม่ได้เลือก Record
        Select Case tbActive
        Case "OPO"
                'กรณีที่ OPO ยังไม่ได้ Check
                If OPoCheck = False Then
                    MovetxtCaption = "Between Processing"
                Else
                'กรณีที่ OPO  Check แล้ว
                    MovetxtCaption = "Can not edit because it checked complete"
                End If
    
        Case Else
                MovetxtCaption = "Can not edit because it checked complete"
        End Select
    End If

                
    'สั่งตัวหนังสือวิ่ง
    With frmActive
        .lblMovetxt.Left = .Width
        .tmMovetxt.Enabled = True
        .lblMovetxt.Caption = MovetxtCaption 'เป็นตัวแปรใน Module Main
        .Caption = frmActive.Caption & "(" & mode & ")"
        .Show vbModal
    End With
    Exit Sub

Err_Edit:
   MsgBox "Please select record to edit.", vbInformation, "Not  found"
   MsgBox Err.Description
End Sub

Public Function FindField(tbName As String, fldName1 As String, fldName2 As String, strText As String) As String
'ใช้ค้นหาค่าในตาราง tbName โดยค้นหาค่า ctrlText จาก fldName2 แล้วดึงค่า fldName1 กลับมา
'tbName = ชื่อ Table, fldName1 = Field ที่ต้องการดึง, fldName2 = Field ที่ใช้อ้างเป็นเงื่อนไข, ctrlText = ค่าที่ค้นหา
    Set rsTmp = New ADODB.Recordset
    strCmdSQL = "SELECT * FROM " & tbName & " WHERE " & fldName2 & "=" & "'" & strText & "'"
    With rsTmp
        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        If .EOF Or .BOF Or IsNull(.Fields(fldName1)) = True Then Exit Function
        FindField = Trim$(.Fields(fldName1))
    End With
    Set rsTmp = Nothing
End Function

Public Function FindField_2condition(tbName As String, fldName As String, fldName1 As String, fldName2 As String, strText1 As String, strText2 As String) As String
'ใช้ค้นหาค่าในตาราง tbName โดยค้นหาค่า strText จาก fldName2 แล้วดึงค่า fldName1 กลับมา
'tbName = ชื่อ Table, fldName1 = Field ที่ต้องการดึง, fldName2 = Field ที่ใช้อ้างเป็นเงื่อนไข, ctrlText = ค่าที่ค้นหา
    Set rsTmp = New ADODB.Recordset
    strCmdSQL = "SELECT * FROM " & tbName & " WHERE " & fldName1 & "=" & "'" & strText1 & "' AND " & fldName2 & "=" & "'" & strText2 & "'"
    With rsTmp
        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        If .EOF Or .BOF Or IsNull(.Fields(fldName1)) = True Then Exit Function
        FindField_2condition = Trim$(.Fields(fldName))
    End With
    Set rsTmp = Nothing
End Function

Public Sub FindAllField(tbName As String, fldName As String, strText As String)
'ใช้ดึงค่าจาก Table มาทุก Field มาเก็บในตัวแปรชื่อ Fld(0 to 50)
'ค้นหาจากตาราง tbName โดยค้นหาค่า ctrlText จาก fldName2 แล้วดึงค่า fldName1 กลับมา
'tbName = ชื่อ Table, fldName1 = Field ที่ต้องการดึง, fldName2 = Field ที่ใช้อ้างเป็นเงื่อนไข, ctrlText = ค่าที่ค้นหา
    Dim ColNo As Byte
    Dim ColName As String
    
    Set rsTmp = New ADODB.Recordset
    strCmdSQL = "SELECT * FROM " & tbName & " WHERE " & fldName & " = " & "'" & strText & "'"
    With rsTmp
        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        For Each Field In .Fields
                If ColNo = .Fields.Count - 1 Then Exit Sub
                Fld(ColNo) = Trim(Field.Value)
                ColNo = ColNo + 1
        Next
    End With
    Set rsTmp = Nothing
End Sub

Public Sub FindList(ctrlText As Control, tbName As String, fldName1 As String, ctrlList As Control)
'ใช้ค้นหาค่า Listindex ของ ComboBox
'เป็น Procedure ที่ใช้ Search ค่า ListIndex, โดยค้นหาจาก fldName1 ที่มีค่า = ctrlText ที่มี Properties "Text"
    Dim strCriteria As String
    If Len(ctrlText.Text) > 0 Then
        Set rsTmp = New ADODB.Recordset
        strCmdSQL = "SELECT * FROM " & tbName & " ORDER BY " & fldName1
        rsTmp.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        With rsTmp
            strCriteria = fldName1 & " = '" & CStr(ctrlText.Text) & "'"
            .Find strCriteria, 0, adSearchForward, adBookmarkFirst
            If .BOF Or .EOF Then
                Exit Sub
            End If
            ctrlList.ListIndex = .AbsolutePosition - 1
        End With
    ElseIf Len(ctrlText.Text) = 0 Then
        ctrlList.ListIndex = -1
    End If
End Sub

Public Sub FindList_2Condition(tbName As String, fldWhere As String, ctrlWhere As Control, FldOrder As String, fldFind As String, ctrlFind As Control, ctrlList As Control)
'ใช้ค้นหาค่า Listindex ของ ComboBox
'เป็น Procedure ที่ใช้ Search ค่า ListIndex, โดยค้นหาจาก fldWhere ที่มีค่า = ctrlText ที่มี Properties "Text"
    Dim strCriteria As String
    If Len(ctrlWhere.Text) <> "" Then
        Set rsTmp = New ADODB.Recordset
        strCmdSQL = "SELECT * FROM " & tbName & " WHERE " & fldWhere & " = '" & Trim$(ctrlWhere.Text) & "' ORDER BY " & FldOrder
        rsTmp.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        With rsTmp
            strCriteria = fldFind & " = '" & Trim$(ctrlFind.Text) & "'"
            .Find strCriteria, 0, adSearchForward, adBookmarkFirst
            If .BOF Or .EOF Then
                Exit Sub
            End If
            ctrlList.ListIndex = .AbsolutePosition - 1
        End With
    ElseIf Len(ctrlText.Text) = 0 Then
        ctrlList.ListIndex = -1
    End If
End Sub

Public Sub Refresh_Record()
With rsActive
    .Requery
End With
    Call FindBfRecord
    With frmdgMain
        Set .dgMain.DataSource = rsActive
            i = 0
        cntcol = 0
        For Each Fields In rsActive.Fields
            'ไม่ให้แสดงบาง Fields ที่ต้องการ
            If InStr(1, HideField, UCase(Trim(rsActive.Fields(i).Name))) <> 0 And Mid(UserActive, 2, 3) <> "MIS" Then
                .dgMain.Columns.Remove cntcol
            ElseIf Right(UCase(Trim(rsActive.Fields(i).Name)), 4) = "FLAG" Then
                        .dgMain.Columns.Remove cntcol
                      Else
                            .dgMain.Columns(cntcol).Caption = Find_FldDesc(rsActive.Fields(i).Name)
                           cntcol = cntcol + 1 'เก็บจำนวน Colum ที่มีอยู่ทั้งหมด
              End If
             i = i + 1
        Next
      '.dgMain.SetFocus
      End With
 End Sub

Public Sub Save_Record()
    On Error GoTo ErrDuplicateKey
    
    Select Case UCase(tbActive) 'กรณีใช้หลาย field เป็น Key
        Case "OPO"
            If Trim(frmActive.txtRunNO.Text) = "***NEW***" Then       'เป็น Auto running no
                Set rsTemp = New Recordset
                With rsTemp
                    .Open ("select *  from Runno where RNtbname = '" & tbActive & "'"), dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                    If Not .EOF Then
                        frmActive.txtRunNO.Text = rsTemp!RNtbcode & rsTemp!RNyear & rsTemp!RNmonth & Format(rsTemp!RNrunno, rsTemp!RNformat)
                        .Close
                        .Open (" Update Runno set  RNrunno=RNrunno+1 where RNtbname = '" & tbActive & "'")
                    Else
                        MsgBox "This Table Not Use Running No.", vbCritical, "Error"
                        Call Cancel_Rec
                        Exit Sub
                    End If
                End With
                Set rsTemp = Nothing
            End If
    End Select
    
    With rsActive
        .Fields!LastUpdate = Now
        .Fields!LastUser = Mid(UserActive, 5, 10)
        .Update
        Exit Sub
        
ErrDuplicateKey:
        .CancelBatch
        MsgBox Err.Description
        'MsgBox "This product was purchased you must change your product!", vbCritical, "Warning!"
    End With
End Sub

Public Sub Delete_Record()
    Dim strMsgBox As String
    Dim bteAnswer As Byte
    
    On Error GoTo Err_Delete
    If rsActive.RecordCount = 0 Then
        MsgBox "Not found record for delete. !", vbInformation + vbOKCancel, "Not found"
        Exit Sub
    End If

    With rsActive
        Select Case tbActive
            
            'กรณีที่จะลบใบ OPO
            Case "OPO"
                strMsgBox = "Are you sure you want to delete record  'OPO NO. " & .Fields(0).Value & "' and details ?"
                bteAnswer = MsgBox(strMsgBox, vbCritical + vbOKCancel, "Confirm Record Delete")
                Select Case bteAnswer
                    Case vbOK
                        .Delete
                        Exit Sub
                    Case vbCancel
                        Exit Sub
                End Select
            
            'กรณีที่จะลบใบ OPODT
            Case "OPODT"
                strMsgBox = "Are you sure you want to delete record 'OPODT Item No. " & .Fields(2).Value & "' ?"
                bteAnswer = MsgBox(strMsgBox, vbCritical + vbOKCancel, "Confirm Record Delete")
                Select Case bteAnswer
                    Case vbOK
                        .Delete
                        Call frmOPODT.UpdateTableOPO
                        Exit Sub
                    Case vbCancel
                        Exit Sub
                End Select
            End Select
    End With

Err_Delete:
   MsgBox "Please select record to delete. !", vbInformation, "Not  found"
   MsgBox Err.Description
    
End Sub

Public Sub Chk_UserActive()
    UsrLevel = Left(UserActive, 1)
    UsrDept = Mid$(UserActive, 2, 3)
    UsrName = Mid$(UserActive, 5, 3)
End Sub

Public Sub Frm_Protect() 'ป้องกันไม่ให้แก้ไข frmOPO
        For Each Ctrl In frmActive
            If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Then
                Ctrl.Enabled = False
                If Ctrl.Name = "cmdExtRem" Or Ctrl.Name = "cmdIntRem" Then Ctrl.Enabled = True
            End If
        Next
End Sub

Public Sub Frm_UnProtect() 'ยกเลิกการป้องกันไม่ให้แก้ไข frmOPO
        For Each Ctrl In frmActive
            If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Then
                Ctrl.Enabled = True
            End If
        Next
End Sub

