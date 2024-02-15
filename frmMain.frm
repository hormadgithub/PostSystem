VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "อารดา เซอร์วิส"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Invoice"
      Height          =   2280
      Left            =   -15
      TabIndex        =   7
      Top             =   5985
      Width           =   16035
      Begin VB.TextBox txtInvno 
         Height          =   360
         Left            =   1155
         TabIndex        =   10
         Tag             =   "INVNO"
         Top             =   495
         Width           =   2355
      End
      Begin VB.TextBox txtCSCode 
         Height          =   360
         Left            =   1140
         TabIndex        =   9
         Tag             =   "CSCode"
         Top             =   945
         Width           =   2355
      End
      Begin VB.TextBox txtInvDate 
         Height          =   360
         Left            =   4755
         TabIndex        =   8
         Tag             =   "InvDate"
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label lblInvno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice No.:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cust Code:"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   4230
         TabIndex        =   11
         Top             =   585
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ลบรายการ"
      Height          =   480
      Left            =   10680
      TabIndex        =   6
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "แก้ไข"
      Height          =   480
      Left            =   4305
      TabIndex        =   5
      Top             =   8505
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "เพิ่มรายการ"
      Height          =   480
      Left            =   2970
      TabIndex        =   4
      Top             =   8505
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   480
      Left            =   6975
      TabIndex        =   3
      Top             =   8505
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "อัพเดต"
      Height          =   480
      Left            =   5610
      TabIndex        =   2
      Top             =   8505
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgInvoice 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   615
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   4524
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
            LCID            =   2057
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
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton lblExit 
      Caption         =   "ออกจากโปรแกรม"
      Height          =   540
      Left            =   14295
      TabIndex        =   0
      Top             =   8460
      Width           =   1665
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2565
      Left            =   -15
      TabIndex        =   15
      Top             =   3180
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   4524
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
            LCID            =   2057
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
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "อันดา เซอร์วิส"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   23.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   -15
      TabIndex        =   14
      Top             =   0
      Width           =   16050
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private rsBrowse As adodb.Recordset
Private itmFound As ListItem   ' FoundItem variable.
Private strPositionAppr As String
Private blnVendorUseHotOrder As Boolean
Private blnVendorSetCost4Digit As Boolean
Private conn As New ADODB.Connection

Private Sub cmdAdd_Click()
 Set rsActive = New ADODB.Recordset
With rsActive
        strCmdSQL = "select  *  From invoice  where  1<>1"
       .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockOptimistic, adCmdText
        .AddNew
End With
Call Define_Field_Tag(Me, rsActive)
txtInvDate.Text = CurrentDate
txtInvno.SetFocus
End Sub

Private Sub cmdCancel_Click()
         With rsActive
                  .CancelBatch
                  .CancelUpdate
         End With
End Sub

Private Sub cmdDelete_Click()
Call Delete_Record("INVOICE", "Invno='" & rsActive!INVNo & "'")
Call RefreshData
End Sub

Private Sub cmdUpdate_Click()
rsActive.UpdateBatch
rsActive.Update
Call RefreshData
End Sub

Private Sub dgInvoice_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call GetData(rsBrowse!INVNo)
End Sub

Private Sub Form_Load()
 'dbActive.Close


 Set rsBrowse = New ADODB.Recordset
With rsBrowse
            strCmdSQL = "select  *  From invoice "
          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockOptimistic, adCmdText
End With
Set dgInvoice.DataSource = rsBrowse

End Sub

Private Sub RefreshData()
rsBrowse.Requery

End Sub

Private Sub lblExit_Click()
Unload Me
End Sub

Private Sub GetData(strInvno As String)
 Set rsActive = New ADODB.Recordset
With rsActive
        If strInvno <> "" Then
            strCmdSQL = "select  *  From invoice  where  invno='" & strInvno & "'"
        Else
            strCmdSQL = "select  *  From invoice "
        End If
        .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockOptimistic, adCmdText
End With
Call Define_Field_Tag(Me, rsActive)
End Sub
