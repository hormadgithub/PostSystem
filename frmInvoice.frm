VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Maintenance"
   ClientHeight    =   4935
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
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
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
      Height          =   3375
      Left            =   -30
      TabIndex        =   0
      Top             =   915
      Width           =   6825
      Begin VB.TextBox txtCSCode 
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
         Left            =   1260
         TabIndex        =   13
         Tag             =   "CSCode"
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox txtInvDate 
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
         Left            =   5565
         TabIndex        =   9
         Tag             =   "InvDate"
         Top             =   375
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtfld1 
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
         Left            =   1260
         TabIndex        =   6
         Tag             =   "Invno"
         Top             =   390
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpInvDate 
         Height          =   315
         Left            =   4020
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   405
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483624
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   124911617
         CurrentDate     =   36949
      End
      Begin VB.Label lblCSName 
         AutoSize        =   -1  'True
         Caption         =   "??"
         Height          =   195
         Left            =   1275
         TabIndex        =   12
         Top             =   1245
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
         TabIndex        =   11
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
      Left            =   0
      TabIndex        =   1
      Top             =   4185
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
               Picture         =   "frmInvoice.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":14BE
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
               Picture         =   "frmInvoice.frx":1D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":20B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":298C
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
               Picture         =   "frmInvoice.frx":3266
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":3580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInvoice.frx":389A
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
      Caption         =   "... Invoice ..."
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
      Left            =   -240
      TabIndex        =   5
      Top             =   135
      Width           =   7305
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   0
      Picture         =   "frmInvoice.frx":4174
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form นี้เป็นแบบ Header Detail
Const Vtable = "Customer" 'ชื่อต้องตรงกับชื่อ Table จริง

Private Sub Save()
If Trim(lblCSName.Caption) = "" Then
  MsgBox "กรุณาป้อนรหัสลูกค้าให้ถูกต้องด้วยนะครับ.", vbCritical, "กรุณาตรวจสอบข้อมูลอีกครั้ง"
  txtCSCode.SetFocus
  Exit Sub
End If

If MsgBox("Are you Sure Save this record ?", vbInformation + vbOKCancel, "Save Record") = vbCancel Then Exit Sub
Call Save_Rec
Unload Me
End Sub



Private Sub cmdLCust_Click()
Call CustomerLookup("Y")
End Sub

Private Sub CustomerLookup(CallLookup As String)
    LookupCondition = ""
    ActiveLookup = "CUSTOMER" 'มีการกรอง ShareDept อยู่แล้ว
    LookupRetVal = Trim(txtCSCode.Text)

    If CallLookup = "Y" Then 'เป็นตัวที่บอกว่าต้องการเรียกใช้ Looup หรือไม่
            frmLookup.Show vbModal 'ถ้าไม่พบค่อยเรียก Lookup แล้วส่งค่า LookupRetVal กลับมาให้
            If CancelLookup Then
               txtfld1.SetFocus
               Exit Sub
            End If
  End If
  If LookupCondition = "" Then
    FindCondition = "CSCODE = '" & Trim(LookupRetVal) & "'"
 Else
    FindCondition = LookupCondition & " and CSCODE = '" & Trim(LookupRetVal) & "'"
 End If
  Set rsTemp = New Adodb.Recordset
  strCmdSQL = "select  *  from " & ActiveLookup & "  where " & FindCondition
  rsTemp.Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not rsTemp.EOF Then
         txtfld1 = Trim(rsTemp!CsCode)
         lblCSName = rsTemp!CSthiName
  Else
         lblCSName = ""
  End If
  rsTemp.Close
  Set rsTemp = Nothing
  
  End Sub


Private Sub dtpInvDate_Change()
txtInvDate.Text = CStr(dtpInvDate.Value)
End Sub

Private Sub Form_Load()
Dim sCond As String
Err_Desc = "" 'กำหนดให้เป็นว่างก่อนเพื่อใช้ตรวจสอบทีหลังได้  Set dbActive = dbSQL
  Call Define_Field_Tag(Me, rsActive) 'ส่ง Form and Adodb.Recordset เพื่อกำหนด Data Source  maxlenght
 If mode = "ADD" Then
    txtInvDate.Text = CurrentDate
    lblCSName.Caption = ""
End If
Call Assign_DateToCtrl(txtInvDate, dtpInvDate)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Cancel_Rec
End Sub


Private Sub OLE1_Updated(Code As Integer)

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

Private Sub txtWKact_date_Validate(Cancel As Boolean)
Cancel = Assign_DateToCtrl(txtWKact_date, dtpWKact_date)
End Sub

Private Sub txtCSCode_Change()
lblCSName.Caption = Find_Ret_Val("Customer", "CsName", "CSCode='" & Trim(txtCSCode.Text) & "'")
End Sub
