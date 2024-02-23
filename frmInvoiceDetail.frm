VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoiceDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Detail  Maintenance"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmInvoiceDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTop1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   3
      Top             =   585
      Width           =   6825
      Begin VB.Timer tmMovetxt 
         Interval        =   20
         Left            =   3465
         Top             =   0
      End
      Begin VB.Label lblMovetxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   135
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4305
      Left            =   -15
      TabIndex        =   0
      Top             =   930
      Width           =   6825
      Begin VB.TextBox txtTrackno 
         Height          =   345
         Left            =   1185
         TabIndex        =   26
         Tag             =   "trackno"
         Top             =   3870
         Width           =   4320
      End
      Begin VB.TextBox txtserviceprice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Tag             =   "serviceprice"
         Text            =   "0"
         Top             =   3435
         Width           =   825
      End
      Begin VB.TextBox txtunitprice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "unitprice"
         Text            =   "0"
         Top             =   2985
         Width           =   825
      End
      Begin VB.TextBox txtweight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1215
         TabIndex        =   19
         Tag             =   "weight"
         Text            =   "0"
         Top             =   2550
         Width           =   825
      End
      Begin VB.TextBox txtInvdt_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1215
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "Invdt_Item"
         Top             =   1215
         Width           =   450
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   405
         TabIndex        =   12
         Top             =   1530
         Width           =   5325
         Begin VB.TextBox txtSendType 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   16
            Tag             =   "SendType"
            Top             =   225
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.OptionButton optREG 
            Caption         =   "ลทบ"
            Height          =   255
            Left            =   810
            TabIndex        =   15
            Top             =   585
            Width           =   1110
         End
         Begin VB.OptionButton optEMS 
            Caption         =   "EMS"
            Height          =   255
            Left            =   795
            TabIndex        =   14
            Top             =   255
            Width           =   1110
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ส่งแบบ :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.TextBox txtInvDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   405
         Width           =   1740
      End
      Begin VB.TextBox txtfld1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "Invno"
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Track No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   375
         TabIndex        =   25
         Top             =   3945
         Width           =   750
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ค่าบริการ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   24
         Top             =   3510
         Width           =   675
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ค่าส่ง:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   705
         TabIndex        =   22
         Top             =   3060
         Width           =   420
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "น้ำหนัก:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   20
         Top             =   2625
         Width           =   555
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   17
         Top             =   1305
         Width           =   435
      End
      Begin VB.Label lblCSName 
         AutoSize        =   -1  'True
         Caption         =   "??"
         Height          =   195
         Left            =   1260
         TabIndex        =   11
         Top             =   885
         Width           =   150
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2955
         TabIndex        =   10
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ลูกค้า.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   615
         TabIndex        =   8
         Top             =   870
         Width           =   510
      End
      Begin VB.Label lblCSCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   405
         Width           =   885
      End
   End
   Begin VB.Frame fraBottom 
      Height          =   735
      Left            =   -45
      TabIndex        =   1
      Top             =   5115
      Width           =   6825
      Begin MSComctlLib.Toolbar tbrSave 
         Height          =   570
         Left            =   5535
         TabIndex        =   2
         Top             =   135
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImgEnabled"
         DisabledImageList=   "ImgDisabled"
         HotImageList    =   "imgHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "บันทึกข้อมูล"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cancel"
               Object.ToolTipText     =   "ยกเลิกการเพิ่ม-แก้ไข"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgDisabled 
         Left            =   1440
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":14BE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgEnabled 
         Left            =   135
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":1D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":20B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":298C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgHot 
         Left            =   810
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":3266
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":3580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoiceDetail.frx":389A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "... Invoice Detail ..."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   -345
      TabIndex        =   5
      Top             =   135
      Width           =   7515
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   0
      Picture         =   "frmInvoiceDetail.frx":4174
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "frmInvoiceDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form นี้เป็นแบบ Header Detail
Const Vtable = "Customer" 'ชื่อต้องตรงกับชื่อ Table จริง
Private blnLoadcomplete As Boolean

Private Sub Save()
If Not IsNumeric(txtweight.Text) Then
  MsgBox "กรุณาทำการป้อน น้ำหนัก ก่อนทำการ Save นะครับ.", vbCritical, "กรุณาตรวจสอบข้อมูลอีกครั้ง"
  txtweight.SetFocus
  Exit Sub
End If

If MsgBox("Are you Sure Save this record ?", vbInformation + vbOKCancel, "Save Record") = vbCancel Then Exit Sub
Call Save_Rec
Unload Me
End Sub

'
'
'Private Sub cmdLCust_Click()
'Call CustomerLookup("Y")
'End Sub
'
'Private Sub CustomerLookup(CallLookup As String)
'    LookupCondition = ""
'    ActiveLookup = "CUSTOMER" 'มีการกรอง ShareDept อยู่แล้ว
'    LookupRetVal = Trim(txtFld2.Text)
'
'    If CallLookup = "Y" Then 'เป็นตัวที่บอกว่าต้องการเรียกใช้ Looup หรือไม่
'            frmLookup.Show vbModal 'ถ้าไม่พบค่อยเรียก Lookup แล้วส่งค่า LookupRetVal กลับมาให้
'            If CancelLookup Then
'               txtfld1.SetFocus
'               Exit Sub
'            End If
'  End If
'  If LookupCondition = "" Then
'    FindCondition = "CSCODE = '" & Trim(LookupRetVal) & "'"
' Else
'    FindCondition = LookupCondition & " and CSCODE = '" & Trim(LookupRetVal) & "'"
' End If
'  Set rsTemp = New Adodb.Recordset
'  strCmdSQL = "select  *  from " & ActiveLookup & "  where " & FindCondition
'  rsTemp.Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
'  If Not rsTemp.EOF Then
'         txtfld1 = Trim(rsTemp!CsCode)
'         lblCSName = rsTemp!CSthiName
'  Else
'         lblCSName = ""
'  End If
'  rsTemp.Close
'  Set rsTemp = Nothing
'
'  End Sub



Private Sub Form_Load()
Dim sCond As String
Dim strCsCode As String
Err_Desc = "" 'กำหนดให้เป็นว่างก่อนเพื่อใช้ตรวจสอบทีหลังได้  Set dbActive = dbSQL

  blnLoadcomplete = False
  Call Define_Field_Tag(Me, rsActive) 'ส่ง Form and Adodb.Recordset เพื่อกำหนด Data Source  maxlenght
sCond = "Invno='" & Trim(rsBrowse!invno) & "'"
 Set rsTemp = New Adodb.Recordset
 With rsTemp
        strCmdSQL = "select *  From Invoice  where  " & sCond
        .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
             txtfld1.Text = rsTemp!invno
             txtInvDate.Text = rsTemp!invdate
             strCsCode = rsTemp!CsCode
       .Close
  End With
  Set rsTemp = Nothing

 lblCSName.Caption = strCsCode & "-" & Find_Ret_Val("Customer", "CsName", "CSCode='" & strCsCode & "'")

 If mode = "ADD" Then
    txtInvdt_Item.Text = max_Item("Invoicedetail", "invdt_item", sCond)
    txtSendType.Text = "EMS"
    txtTrackno.Text = ""
End If
optEMS.Value = IIf(Trim(txtSendType.Text) = "EMS", True, False)
optREG.Value = IIf(Trim(txtSendType.Text) = "REG", True, False)

  blnLoadcomplete = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Cancel_Rec
End Sub


Private Sub optEMS_Click()
txtSendType.Text = "EMS"
End Sub

Private Sub optREG_Click()
txtSendType.Text = "REG"
End Sub

Private Sub tbrSave_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
        Case "SAVE"
            Call Save
        Case "CANCEL"
            Unload Me
End Select
End Sub

Private Sub tmMovetxt_Timer()
With lblMovetxt
   If .Left <= -(lblMovetxt.Width) Then .Left = Me.Width
 .Left = .Left - 10
End With
End Sub



Private Sub txtSendType_Change()
Call calRateAndservice
End Sub

Private Sub txtweight_Validate(Cancel As Boolean)
Call calRateAndservice
End Sub

Private Sub calRateAndservice()
        If Not blnLoadcomplete Then Exit Sub
        If Not IsNumeric(txtweight.Text) Then txtweight.Text = "0"
        txtunitprice.Text = fn_GetRate(Trim(txtSendType.Text), CDouble(txtweight.Text))
        txtserviceprice.Text = fn_GetService(Trim(txtSendType.Text), CDouble(txtweight.Text))
End Sub
