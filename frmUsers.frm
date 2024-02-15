VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Maintenance"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
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
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6720
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
      Left            =   -90
      TabIndex        =   0
      Top             =   930
      Width           =   6825
      Begin VB.TextBox txtfld1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         TabIndex        =   7
         Tag             =   "UserId"
         Top             =   435
         Width           =   1710
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1350
         TabIndex        =   6
         Tag             =   "Password"
         Top             =   990
         Width           =   1710
      End
      Begin VB.Label lblUserID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User ID :"
         Height          =   195
         Left            =   570
         TabIndex        =   9
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   1095
         Width           =   780
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
               Picture         =   "frmUsers.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":14BE
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
               Picture         =   "frmUsers.frx":1D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":20B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":298C
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
               Picture         =   "frmUsers.frx":3266
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":3580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":389A
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
      Caption         =   "... Users ..."
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
      Left            =   -360
      TabIndex        =   5
      Top             =   135
      Width           =   7545
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   0
      Picture         =   "frmUsers.frx":4174
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form นี้เป็นแบบ Header Detail
Const Vtable = "Customer" 'ชื่อต้องตรงกับชื่อ Table จริง

Private Sub Save()
If MsgBox("Are you Sure Save this record ?", vbInformation + vbOKCancel, "Save Record") = vbCancel Then Exit Sub
Call Save_Rec
Unload Me
End Sub

Private Sub Form_Activate()
  txtfld1.Enabled = True
 If mode = "ADD" Then
   txtfld1.SetFocus
   'SendKeys "{Home}+{End}"
Else
  txtPassword.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim sCond As String
Err_Desc = "" 'กำหนดให้เป็นว่างก่อนเพื่อใช้ตรวจสอบทีหลังได้  Set dbActive = dbSQL
  Call Define_Field_Tag(Me, rsActive) 'ส่ง Form and Adodb.Recordset เพื่อกำหนด Data Source  maxlenght
  If mode = "ADD" Then
     txtfld1.Text = ""
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Cancel_Rec
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

