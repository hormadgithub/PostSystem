Attribute VB_Name = "mdCenter"
    '���� Recordset ���Ǥ���
    Public rsTmp As ADODB.Recordset
    '��᷹����� SQL
    Public strCmdSQL  As String
    '��᷹ Control � Form
    Public Ctrl As Control
    
    '��᷹ Usr ��� Logon ����к�
    Public UsrName As String * 3
    Public UsrDept As String * 3
    Public UsrLevel As Variant
    
    '���纤������� OPO �ա�� Check �����١��ͧ����
    Public OPoCheck As Boolean
    
    'Procedure FindAllFields()
    '���纤�� Fields �ء Fields � Record
    Public Field As Field
    Public Fld(0 To 50)

Public Sub AddList(tbName As String, fldName1 As String, ctrlList As Control, Condition As String)
'�� Procedure ����� Add ��� fldName1 � tbName ����� ctrlList
'tbName = ���� Table, fldName1 = Field ����ͧ��ô֧, ctrlList = Control ����� Method AddItem
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
'�� Procedure ����� Add ��� fldName1 ��� fldName2 ������§ fldName1 ���ҧ����� tbName ����� ctrlList
'tbName = ���� Table, fldName1 = Field ����ͧ��ô֧, ctrlList = Control ����� Method AddItem
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
        
   MovetxtCaption = "���ѧ�ӡ��������������� " & tbActive & "..."

    With frmActive
        .lblMovetxt.Left = .Width
        .tmMovetxt.Enabled = True
        .lblMovetxt.Caption = MovetxtCaption   '�繵����� Module Main
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
    If rsActive.AbsolutePosition > 0 Then '���㹡óշ����������͡ Record
        Select Case tbActive
        Case "OPO"
                '�óշ�� OPO �ѧ����� Check
                If OPoCheck = False Then
                    MovetxtCaption = "Between Processing"
                Else
                '�óշ�� OPO  Check ����
                    MovetxtCaption = "Can not edit because it checked complete"
                End If
    
        Case Else
                MovetxtCaption = "Can not edit because it checked complete"
        End Select
    End If

                
    '��觵��˹ѧ������
    With frmActive
        .lblMovetxt.Left = .Width
        .tmMovetxt.Enabled = True
        .lblMovetxt.Caption = MovetxtCaption '�繵����� Module Main
        .Caption = frmActive.Caption & "(" & mode & ")"
        .Show vbModal
    End With
    Exit Sub

Err_Edit:
   MsgBox "Please select record to edit.", vbInformation, "Not  found"
   MsgBox Err.Description
End Sub

Public Function FindField(tbName As String, fldName1 As String, fldName2 As String, strText As String) As String
'����Ҥ��㹵��ҧ tbName �¤��Ҥ�� ctrlText �ҡ fldName2 ���Ǵ֧��� fldName1 ��Ѻ��
'tbName = ���� Table, fldName1 = Field ����ͧ��ô֧, fldName2 = Field �������ҧ�����͹�, ctrlText = ��ҷ�����
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
'����Ҥ��㹵��ҧ tbName �¤��Ҥ�� strText �ҡ fldName2 ���Ǵ֧��� fldName1 ��Ѻ��
'tbName = ���� Table, fldName1 = Field ����ͧ��ô֧, fldName2 = Field �������ҧ�����͹�, ctrlText = ��ҷ�����
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
'��֧��Ҩҡ Table �ҷء Field ����㹵���ê��� Fld(0 to 50)
'���Ҩҡ���ҧ tbName �¤��Ҥ�� ctrlText �ҡ fldName2 ���Ǵ֧��� fldName1 ��Ѻ��
'tbName = ���� Table, fldName1 = Field ����ͧ��ô֧, fldName2 = Field �������ҧ�����͹�, ctrlText = ��ҷ�����
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
'����Ҥ�� Listindex �ͧ ComboBox
'�� Procedure ����� Search ��� ListIndex, �¤��Ҩҡ fldName1 ����դ�� = ctrlText ����� Properties "Text"
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
'����Ҥ�� Listindex �ͧ ComboBox
'�� Procedure ����� Search ��� ListIndex, �¤��Ҩҡ fldWhere ����դ�� = ctrlText ����� Properties "Text"
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
            '�������ʴ��ҧ Fields ����ͧ���
            If InStr(1, HideField, UCase(Trim(rsActive.Fields(i).Name))) <> 0 And Mid(UserActive, 2, 3) <> "MIS" Then
                .dgMain.Columns.Remove cntcol
            ElseIf Right(UCase(Trim(rsActive.Fields(i).Name)), 4) = "FLAG" Then
                        .dgMain.Columns.Remove cntcol
                      Else
                            .dgMain.Columns(cntcol).Caption = Find_FldDesc(rsActive.Fields(i).Name)
                           cntcol = cntcol + 1 '�纨ӹǹ Colum ��������������
              End If
             i = i + 1
        Next
      '.dgMain.SetFocus
      End With
 End Sub

Public Sub Save_Record()
    On Error GoTo ErrDuplicateKey
    
    Select Case UCase(tbActive) '�ó������� field �� Key
        Case "OPO"
            If Trim(frmActive.txtRunNO.Text) = "***NEW***" Then       '�� Auto running no
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
            
            '�óշ���ź� OPO
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
            
            '�óշ���ź� OPODT
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

Public Sub Frm_Protect() '��ͧ�ѹ��������� frmOPO
        For Each Ctrl In frmActive
            If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Then
                Ctrl.Enabled = False
                If Ctrl.Name = "cmdExtRem" Or Ctrl.Name = "cmdIntRem" Then Ctrl.Enabled = True
            End If
        Next
End Sub

Public Sub Frm_UnProtect() '¡��ԡ��û�ͧ�ѹ��������� frmOPO
        For Each Ctrl In frmActive
            If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Then
                Ctrl.Enabled = True
            End If
        Next
End Sub

