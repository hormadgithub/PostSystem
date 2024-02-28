VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLookup 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lookup"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOrder 
      Height          =   315
      ItemData        =   "frmLookup.frx":030A
      Left            =   90
      List            =   "frmLookup.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1485
      Width           =   1455
   End
   Begin VB.ComboBox cmbSort 
      Height          =   315
      ItemData        =   "frmLookup.frx":032E
      Left            =   90
      List            =   "frmLookup.frx":0330
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   855
      Width           =   2580
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   345
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3645
      Width           =   1155
   End
   Begin VB.ComboBox cmbCndt 
      Height          =   315
      ItemData        =   "frmLookup.frx":0332
      Left            =   105
      List            =   "frmLookup.frx":034E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2205
      Width           =   1005
   End
   Begin VB.TextBox txtCndt 
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   2970
      Width           =   2430
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4275
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4770
      Width           =   1110
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   4725
      TabIndex        =   0
      Top             =   4185
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   4065
      Left            =   2835
      TabIndex        =   13
      Top             =   45
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   7170
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lookup"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5550
      Left            =   2700
      TabIndex        =   12
      Top             =   -225
      Width           =   105
   End
   Begin MSComctlLib.StatusBar stbLookup 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5220
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   1
            Object.Width           =   5821
            Text            =   "Lookup Table :"
            TextSave        =   "NUM"
            Key             =   "lookuptable"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
            Text            =   "Record No :"
            TextSave        =   "Record No :"
            Key             =   "record"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   16
      Top             =   1245
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operator :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   1935
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      X1              =   90
      X2              =   2565
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operan Value :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   2700
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for record in table "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   285
      TabIndex        =   9
      Top             =   165
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   45
      Picture         =   "frmLookup.frx":0377
      Top             =   135
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order By:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   615
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search in results :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3015
      TabIndex        =   3
      Top             =   4275
      Width           =   1710
   End
   Begin VB.Image Image2 
      Height          =   7200
      Left            =   0
      Picture         =   "frmLookup.frx":0571
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10995
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrFld As String 'Field ที่ต้องการแสดงตอน Lookup
Dim i As Integer
Public strTBName As String


Private Sub cmbOrder_Click()
Call cmbSort_Click
End Sub

Private Sub cmbSort_Click()
Dim FldOrder As String
Dim SelOrder As String * 5
           If cmbOrder.ListIndex < 0 Then Exit Sub
           SelOrder = IIf(cmbOrder.ListIndex = 0, " ASC", " DESC")
          FldOrder = cmbSort.Text
          If rsLookup.RecordCount = 0 Or FldOrder = "" Then Exit Sub
          rsLookup.Sort = FldOrder & SelOrder
          'Call FindBfRecord
End Sub

Private Sub cmdCancel_Click()
          CancelLookup = True
          rsLookup.Close
          Set rsLookup = Nothing
          Unload Me
End Sub

Private Sub cmdOk_Click()
         Dim i As Integer
         Dim j As Integer
         Dim sClmnName As String
         
         If rsLookup.AbsolutePosition < 0 Then
                    MsgBox "กรุณาเลือกข้อมูลที่ต้องการก่อน Click OK ", vbInformation, "Please Click Data."
                   Call cmdRefresh_Click
                    Exit Sub
         End If
        
        If Trim(LookupColumn) = "" Then
          
                  Select Case UCase(ActiveLookup)
                  Case "VW_PART_CRODT"
                            LookupRetVal = Trim(rsLookup.fields(3))
                        
                  Case "CRODT", "EXPRESS_STMAS", "EXPRESSST_STMAS"
                            LookupRetVal = Trim(rsLookup.fields(1))
                            
                  Case "LOCATION"
                            LookupRetVal = Trim(rsLookup.fields(1))
                  
                  Case "PARTCATEGORY"
                            LookupRetVal = Trim(rsLookup.fields(2))
                        
                  Case "PARTSUBCATE"
                            LookupRetVal = Trim(rsLookup.fields(3))
                        
                  Case Else
                            LookupRetVal = Trim(rsLookup.fields(0))     'Field ที่เป็น Key ของ Table
                        
                  End Select
         Else
                  i = 1
                  LookupRetVal = ""
                  j = Len(LookupColumn)
                  Do Until i = 0
                           i = InStr(i, LookupColumn, ",")
                           If i = 0 Then
                                    If Trim(LookupRetVal) = "" Then
                                             LookupRetVal = Trim(rsLookup.fields(LookupColumn))
                                    Else
                                             LookupRetVal = LookupRetVal & ", " & Trim(rsLookup.fields(LookupColumn))
                                    End If
                           Else
                                    sClmnName = Left(LookupColumn, i - 1)
                                    LookupColumn = Mid(LookupColumn, i + 1, j - i)
                                    If Trim(LookupRetVal) = "" Then
                                             LookupRetVal = Trim(rsLookup.fields(sClmnName))
                                    Else
                                             LookupRetVal = LookupRetVal & ", " & Trim(rsLookup.fields(LookupColumn))
                                    End If
                           End If
                  Loop
                  
         End If
        
          CancelLookup = False
          rsLookup.Close
          Set rsLookup = Nothing
          Unload Me
 End Sub

Private Sub cmdRefresh_Click()
Dim FldName As String
          FldName = Trim(Find_FldCode(cmbSort.Text))
          txtSearch.Text = ""
          With rsLookup
                    If FieldTypeNumeric(rsLookup, FldName) Then
                                 If cmbCndt.Text = "Like" Then cmbCndt.ListIndex = 2 'ถ้าเลือก Like แล้ว Field เป็น Numeric ให้เปลี่ยนเป็น = แทน
                                 If cmbCndt.Text = "Not Like" Then cmbCndt.ListIndex = 3 'ถ้าเลือก Not Like แล้ว Field เป็น Numeric ให้เปลี่ยนเป็น <> แทน
                                .Close
                                If Trim(StrJoinTable) <> "" Then
                                        strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "," & StrJoinTable & "   where " & FldName & cmbCndt.Text & " " & CDouble(Trim(txtCndt.Text)) & "  and " & LookupCondition & "  order by " & FldName
                                        .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                               Else
                                        strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "  where " & FldName & cmbCndt.Text & "  " & CDouble(Trim(txtCndt.Text)) & " and " & LookupCondition & "order by " & FldName
                                        .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                            End If
                    Else 'ไม่ใช่ Numeric
                            .Close
                            If Trim(StrJoinTable) <> "" Then
                                If cmbCndt.Text = "Like" Or cmbCndt.Text = "Not Like" Then
                                          strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "," & StrJoinTable & "   where " & FldName & "  " & cmbCndt.Text & "  '" & Trim(txtCndt.Text) & "%' and " & LookupCondition & "  order by " & FldName
                                          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                                Else
                                          strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "," & StrJoinTable & "  where " & FldName & "  " & cmbCndt.Text & "'" & Trim(txtCndt.Text) & "' and " & LookupCondition & "order by " & FldName
                                          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                                End If
                           Else
                                If cmbCndt.Text = "Like" Then
                                          strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "   where " & FldName & "  " & cmbCndt.Text & "  '" & Trim(txtCndt.Text) & "%' and " & LookupCondition & "  order by " & FldName
                                          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                                Else
                                          strCmdSQL = "select " & StrFld & "  from " & ActiveLookup & "  where " & FldName & "  " & cmbCndt.Text & "'" & Trim(txtCndt.Text) & "' and " & LookupCondition & "order by " & FldName
                                          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                                End If
                        End If
                    End If
          End With
          Set dgMain.DataSource = rsLookup
        'กำหนด  Colum Header  ให้กับ Data Grid มี cmbSort
        Call Assign_HeaderGrid(rsLookup, dgMain)
          
          If Trim(LookupTitle) = "" Then
                    dgMain.Caption = ActiveLookup & "  Lookup"
          Else
                    dgMain.Caption = LookupTitle
          End If
End Sub

Private Sub dgMain_DblClick()
          Call cmdOk_Click
End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
        cmbSort.Text = dgMain.Columns(ColIndex).Caption
        Call cmbSort_Click
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        stbLookup.Panels("record").Text = "Record No :" & rsLookup.AbsolutePosition & "/" & rsLookup.RecordCount
End Sub

Private Sub Form_Activate()
If Trim(txtSearch.Text) <> "" Then
        If InStr(1, UCase(ActiveLookup), "QUOTATIONDT") <> 0 Then
            strTBName = "Quotationdt"
        ElseIf InStr(1, UCase(ActiveLookup), "COMPOSDT") <> 0 Or InStr(1, UCase(ActiveLookup), "CMPDT") <> 0 Then
            strTBName = "Composdt"
        ElseIf InStr(1, UCase(ActiveLookup), "CPBDT") <> 0 Then
            strTBName = "CPBDT"
        ElseIf InStr(1, UCase(ActiveLookup), "SCRAP_DAMAGEDT") <> 0 Then
            strTBName = "SCRAP_DAMAGEDT"
        Else
            strTBName = ""
        End If
     '19/05/2010 santi ถ้ามีมากกว่า 1 ตัวเช่น Item ทำให้เกิด Error เนื่องจากโปรมแกรมจะค้นหาโดยเทียบกับ TbActive ซึ่งไม่ตรงกับ Table ที่ต้องการค้นหา
     'If FieldTypeNumeric(rsLookup, cmbSort.Text) Then  'ถ้าเป็นตัวเลขไม่ต้องมี Qute ค่อม
    If FieldTypeNumeric(rsLookup, cmbSort.Text) Then  'ถ้าเป็นตัวเลขไม่ต้องมี Qute ค่อม
        rsLookup.Find cmbSort.Text & " = " & txtSearch.Text
     Else
        If UCase(ActiveLookup) <> "VW_MOVEMENT_PARTSERIALNODT" And Trim(cmbSort.Text) <> "" Then
            rsLookup.Find Find_FldCode(cmbSort.Text) & " like '" & txtSearch.Text & "%'"
        End If
    End If
End If
dgMain.SetFocus
'frmMain.stbfrmmain.Panels("record").Text = "Record No :" & rsLookup.AbsolutePosition & "/" & rsLookup.RecordCount
'กำหนดการแรียงลำดับ
      '13/03/2012 กำหนดการแรียงลำดับ
        If LookupOrderType <> "" Then
            If UCase$(LookupOrderType) = "ASC" Then
                cmbOrder.ListIndex = 0
            Else
                cmbOrder.ListIndex = 1
            End If
         Else
            cmbOrder.ListIndex = 0    'เรียงจากน้อยไปมาก
'            Select Case UCase(ActiveLookup)
'                         Case "VW_LOOKUP_OPORCVFORSHIPNOTE", "VW_LOOKUP_PACKING", "VW_LOOKUP_OPO", "VW_LOOKUP_QUOTATION"
'                                    cmbOrder.ListIndex = 1        'เรียงจากมากไปน้อย
'                         Case "PRICEREQUEST", "PRICEREQUESTDT", "VW_TRANSPORTS_VEHICLE" ', "VW_LOOKUP_OPODTFORPACKING"
'                                    cmbOrder.ListIndex = 0        'เรียงจากน้อยไปมาก
'                         Case Else 'ตรวจสอบว่าปกติ Table เรียงมากน้อยไปมากหรือไม่
'                                    If (Count_Record("AllTable", "TBName='" & ActiveLookup & "' And TBOrderByAsc = 'Y'") <> 0) Or UCase(ActiveLookup) = "PARTTUBE" Then
'                                        cmbOrder.ListIndex = 0    'เรียงจากน้อยไปมาก
'                                    Else
'                                        If InStr(1, UCase(dgMain.Columns(0).Caption), "ITEM") <> 0 Then
'                                                cmbOrder.ListIndex = 0    'เรียงจากน้อยไปมาก
'                                        Else
'                                             cmbOrder.ListIndex = 1    'เรียงจากมากไปน้อย
'                                        End If
'                                    End If
'            End Select
    End If
    txtSearch.SetFocus
End Sub

Private Sub Form_Load()
Err_Desc = "" 'กำหนดให้เป็นว่างก่อนเพื่อใช้ตรวจสอบทีหลังได้
        
       Dim cmbIndex As Integer
        Dim FldName As String
        If LookupCondition = "" Then LookupCondition = "1=1"
        StrFld = "" 'กำหนดให้เป็นว่างไว้ก่อน
        StrJoinTable = ""
'        If Left(UCase(ActiveLookup), 2) <> "VW" And Left(UCase(ActiveLookup), 4) <> "VIEW" Then 'แสดงว่าไม่ได้มาจาก View
'                'เอา TBShowField จาก   AllTable  มาใส่ให้กับตัวแปร strFld สำหรับแสดงตอน Brows
'                Set rsLookup = New Adodb.Recordset
'                With rsLookup
'                        strCmdSQL = "SELECT  *  FROM Alltable  WHERE TBName= '" & Trim(UCase(ActiveLookup)) & "'"
'                        .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
'                        If Not .BOF Or Not .EOF Then
'                                StrFld = rsLookup!TBShowField
'                                cmbIndex = CInt(rsLookup!TBcmbOrderIndex)
'                        Else
'                                StrFld = ""
'                        End If
'                        .Close
'                 End With
'                 Set rsLookup = Nothing
'        End If
        If Trim$(StrFld) = "" Then StrFld = "*"
        Set rsLookup = New Adodb.Recordset
        With rsLookup
                strCmdSQL = "SELECT " & StrFld & " FROM " & ActiveLookup & " WHERE " & LookupCondition
                '18-08-2023 ต้องการ Order ตามต้องการ
                If UCase(ActiveLookup) = "VW_MOVEMENT_PARTSERIALNODT" Then
                    strCmdSQL = strCmdSQL & " ORDER BY DocDate, DocType"
                End If
                .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
               ' .Sort = Trim(rsLookup.fields(0).Name)
        End With
        Set dgMain.DataSource = rsLookup
        'กำหนด  Colum Header  ให้กับ Data Grid มี cmbSort
        Call Assign_HeaderGrid_cmb(rsLookup, dgMain, cmbSort)

        txtSearch.Text = LookupRetVal

'        If Left(UCase(ActiveLookup), 2) <> "VW" And Left(UCase(ActiveLookup), 4) <> "VIEW" Then  'แสดงว่าไม่ได้มาจาก View
'           If UCase(ActiveLookup) = "PARTTUBE" Then
'              cmbSort.ListIndex = 2 'Real_OH
'           Else
'            cmbSort.ListIndex = cmbIndex 'ได้ค่ามาจากการกำหนดที่ All Table
'           End If
'        Else
'            '30-10-2018 ต้องการให้เรียงลำดับตาม View
'           If UCase(ActiveLookup) <> "VW_MOVEMENT_PARTSERIALNODT" Then cmbSort.ListIndex = 0
'        End If
        cmbCndt.ListIndex = 0

        If Trim(LookupTitle) = "" Then
            dgMain.Caption = ActiveLookup & "  Lookup"
        Else
            dgMain.Caption = LookupTitle
        End If
                
        '13/03/2012 ถ้ามีการกำหนด Field ในการเรียงลำดับ
'        If LookupOrderField <> "" Then
'           Dim strOrderField As String
'           strOrderField = Find_FldDesc(LookupOrderField)
'           For i = 0 To cmbSort.ListCount
'                     cmbSort.ListIndex = i
'                  If UCase(strOrderField) = UCase(cmbSort.Text) Then Exit For
'           Next
'        End If
        

    cmbSort.ListIndex = 0

      '13/03/2012 กำหนดการแรียงลำดับ
        If LookupOrderType <> "" Then
            If UCase$(LookupOrderType) = "ASC" Then
                cmbOrder.ListIndex = 0
            Else
                cmbOrder.ListIndex = 1
            End If
        End If
        
        stbLookup.Panels("lookuptable").Text = "Lookup Table :" & ActiveLookup
        Me.Caption = dgMain.Caption
        stbLookup.Panels("record").Text = "Record No :" & rsLookup.AbsolutePosition & "/" & rsLookup.RecordCount
        cmdOk.Enabled = (rsLookup.RecordCount <> 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
      'ป้องกันค่าที่ไม่เคลียร์ออก
       LookupColumn = ""
       LookupOrderField = ""
       LookupOrderType = ""
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim SavRec As Long
Dim FldName As String


On Error GoTo err_rec:
    With rsLookup
    FldName = Trim(cmbSort.Text)   'Trim(Find_FldCode(cmbSort.Text, ActiveLookup))
    If txtSearch.Text <> "" And .RecordCount <> 0 Then
          SavRec = .AbsolutePosition
          .MoveFirst
          'สำหรับ ตัวเลขทั้งหมดไม่ต้องใส่ Qoute
           If FieldTypeNumeric(rsLookup, FldName) Then
                .Find FldName & " like " & Trim(txtSearch.Text)
            Else 'ต้องใส่ Quote
                 .Find FldName & " like '" & Trim(txtSearch.Text) & "%'"
           End If
         If .EOF Then .AbsolutePosition = SavRec
    End If
    End With
    '29/01/2010
    If KeyCode = vbKeyReturn Then
       Call cmdOk_Click
    End If
Exit Sub
err_rec:
           MsgBox "Please click grid before search", vbInformation, "Click grid"
End Sub

