VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPostSystem 
   BackColor       =   &H009FBF9F&
   BorderStyle     =   0  'None
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleMode       =   0  'User
   ScaleWidth      =   12771.64
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBrowse 
      BackColor       =   &H009FBF9F&
      BorderStyle     =   0  'None
      Height          =   4005
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   12675
      Begin Crystal.CrystalReport crptWKChecklist 
         Left            =   6465
         Top             =   3285
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Crystal.CrystalReport crptInvoice 
         Left            =   6840
         Top             =   3285
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.TextBox txtFindValue 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   645
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3375
         Width           =   2100
      End
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   3060
         Left            =   -30
         TabIndex        =   6
         Top             =   15
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   5398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761087
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Label lblCntRecord 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   5280
         TabIndex        =   15
         Top             =   3330
         Width           =   450
      End
      Begin VB.Label lblTbActive_Des 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         Height          =   195
         Left            =   8550
         TabIndex        =   14
         Top             =   3555
         Width           =   180
      End
      Begin VB.Label lblTbActive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12240
         TabIndex        =   13
         Top             =   3375
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label txtTbActive 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟ้มที่เลือก >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7395
         TabIndex        =   12
         Top             =   3555
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหา:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   3390
         Width           =   480
      End
      Begin VB.Label lblFirst 
         AutoSize        =   -1  'True
         BackColor       =   &H009FBF9F&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         TabIndex        =   10
         Top             =   3285
         Width           =   315
      End
      Begin VB.Label lblLast 
         AutoSize        =   -1  'True
         BackColor       =   &H009FBF9F&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   9
         Top             =   3285
         Width           =   315
      End
      Begin VB.Label lblPrevious 
         AutoSize        =   -1  'True
         BackColor       =   &H009FBF9F&
         Caption         =   "<|"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3780
         TabIndex        =   8
         Top             =   3285
         Width           =   255
      End
      Begin VB.Label lblNext 
         AutoSize        =   -1  'True
         BackColor       =   &H009FBF9F&
         Caption         =   "|>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4275
         TabIndex        =   7
         Top             =   3285
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar tbrExit 
      Height          =   570
      Left            =   11910
      TabIndex        =   3
      Top             =   4695
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "imlEnabled"
      DisabledImageList=   "imlDisabled"
      HotImageList    =   "imlHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrSave 
      Height          =   570
      Left            =   5490
      TabIndex        =   2
      Top             =   3930
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "imlEnabled"
      DisabledImageList=   "imlDisabled"
      HotImageList    =   "imlHot"
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
   Begin MSComctlLib.Toolbar tbrCommand 
      Height          =   705
      Left            =   -15
      TabIndex        =   1
      Top             =   4620
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   1244
      ButtonWidth     =   1217
      ButtonHeight    =   1244
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "iml_E"
      DisabledImageList=   "iml_D"
      HotImageList    =   "iml_H"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "ADD"
            ImageKey        =   "ADD"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "EDIT"
            ImageKey        =   "EDIT"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "DELETE"
            ImageKey        =   "DELETE"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "REFRESH"
            ImageKey        =   "REFRESH"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Criteria"
            Key             =   "CRITERIA"
            ImageKey        =   "CRITERIA"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PRINT"
            Key             =   "PRINT"
            Object.ToolTipText     =   "PRINT"
            ImageKey        =   "PRINT"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList iml_E 
         Left            =   6975
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   25
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":0000
               Key             =   "ADD"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":0400
               Key             =   "EDIT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":082A
               Key             =   "DELETE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":0C93
               Key             =   "REFRESH"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":10A1
               Key             =   "CRITERIA"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":1DD1
               Key             =   "PRINT"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList iml_H 
         Left            =   7380
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   25
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":3053
               Key             =   "ADD"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":342D
               Key             =   "EDIT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":3898
               Key             =   "DELETE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":3D4E
               Key             =   "REFRESH"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":41DF
               Key             =   "CRITERIA"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":4F7F
               Key             =   "PRINT"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList iml_D 
         Left            =   7965
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   25
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":5859
               Key             =   "ADD"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":5B7E
               Key             =   "EDIT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":5F26
               Key             =   "DELETE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":62BF
               Key             =   "REFRESH"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":6666
               Key             =   "CRITERIA"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":7219
               Key             =   "PRINT"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlHot 
         Left            =   8730
         Top             =   45
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
               Picture         =   "frmPostSystem.frx":7AF3
               Key             =   "SAVE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":80DF
               Key             =   "CLOSE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":8676
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlEnabled 
         Left            =   9000
         Top             =   45
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
               Picture         =   "frmPostSystem.frx":8D88
               Key             =   "SAVE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":92B6
               Key             =   "CLOSE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":9721
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlDisabled 
         Left            =   9405
         Top             =   45
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
               Picture         =   "frmPostSystem.frx":9DB3
               Key             =   "SAVE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":A326
               Key             =   "CLOSE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPostSystem.frx":A7F2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMain 
      Appearance      =   0  'Flat
      BackColor       =   &H009FBF9F&
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   -45
      TabIndex        =   0
      Top             =   5205
      Width           =   12660
      Begin TabDlg.SSTab stbDetail1 
         Height          =   3420
         Left            =   60
         TabIndex        =   17
         Top             =   150
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   6033
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   -2147483626
         ForeColor       =   8388608
         TabCaption(0)   =   "รายการ"
         TabPicture(0)   =   "frmPostSystem.frx":AF04
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "dgDetail1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "ลูกค้า"
         TabPicture(1)   =   "frmPostSystem.frx":AF20
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "dgDetail2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "---EMS---"
         TabPicture(2)   =   "frmPostSystem.frx":AF3C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgDetail3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "ลงทะเบียน"
         TabPicture(3)   =   "frmPostSystem.frx":AF58
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgDetail4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "ผู้ใช้(Users)"
         TabPicture(4)   =   "frmPostSystem.frx":AF74
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "dgDetail5"
         Tab(4).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgDetail1 
            Height          =   3000
            Left            =   -74985
            TabIndex        =   18
            Top             =   315
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   5292
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
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
         Begin MSDataGridLib.DataGrid dgDetail2 
            Height          =   3045
            Left            =   0
            TabIndex        =   19
            Top             =   330
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   5371
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12640511
            HeadLines       =   1
            RowHeight       =   15
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
         Begin MSDataGridLib.DataGrid dgDetail3 
            Height          =   3045
            Left            =   -74985
            TabIndex        =   20
            Top             =   330
            Width           =   12510
            _ExtentX        =   22066
            _ExtentY        =   5371
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12632319
            HeadLines       =   1
            RowHeight       =   15
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
         Begin MSDataGridLib.DataGrid dgDetail4 
            Height          =   3045
            Left            =   -74970
            TabIndex        =   21
            Top             =   315
            Width           =   12570
            _ExtentX        =   22172
            _ExtentY        =   5371
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12648384
            HeadLines       =   1
            RowHeight       =   15
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
         Begin MSDataGridLib.DataGrid dgDetail5 
            Height          =   3045
            Left            =   -75000
            TabIndex        =   22
            Top             =   330
            Width           =   12570
            _ExtentX        =   22172
            _ExtentY        =   5371
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12640511
            HeadLines       =   1
            RowHeight       =   15
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
      End
   End
   Begin VB.Label lblTitle 
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
      Height          =   585
      Left            =   0
      TabIndex        =   16
      Top             =   15
      Width           =   12975
   End
End
Attribute VB_Name = "frmPostSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnMouseRightClick As Boolean

Private Sub Form_Load()
        Set frmActive = Me
        lblTbActive = "INVOICE"
        lblTbActive_Des = "แฟ้มใบเสร็จ"
      
        Call InitializeLoadForm
        
        With frmActive
                Call Define_MouseIcon_HandPoint(tbrExit)
                Set rsBrowse1 = New Adodb.Recordset
                Set rsBrowse2 = New Adodb.Recordset
                Set rsBrowse3 = New Adodb.Recordset
                Set rsBrowse4 = New Adodb.Recordset
                Set rsBrowse5 = New Adodb.Recordset
                Set rsBrowse6 = New Adodb.Recordset
                
                 Set rsBrowse = New Adodb.Recordset
                With rsBrowse
                            strCmdSQL = "select  *  From invoice "
                          .Open strCmdSQL, dbActive, adOpenForwardOnly, adLockOptimistic, adCmdText
                End With
                Set dgMain.DataSource = rsBrowse

                'Call BrowseDetail1(False)
                
                'กำหนดความยาวของ Object ต่างๆ
                stbDetail1.Width = dgMain.Width
                stbDetail1.Tab = 0
                stbDetail1.Left = 0
                
                dgDetail1.Width = stbDetail1.Width - 120
                dgDetail2.Width = stbDetail1.Width - 120
                dgDetail3.Width = stbDetail1.Width - 120
                dgDetail4.Width = stbDetail1.Width - 120
                dgDetail5.Width = stbDetail1.Width - 120
'                dgDetail6.Width = stbDetail1.Width - 120
                
                dgDetail1.Height = stbDetail1.Height - 300
                dgDetail2.Height = stbDetail1.Height - 300
                dgDetail3.Height = stbDetail1.Height - 300
                dgDetail4.Height = stbDetail1.Height - 300
                dgDetail5.Height = stbDetail1.Height - 300
'                dgDetail6.Height = stbDetail1.Height - 300
                
                
        End With
'        cmbWkT.ListIndex = 0
        Call dgMain_HeadClick(0)
        Call Set_Button(rsBrowse)
        Call RowColChg
        Call BrowseDetail1(True)
        lblCntRecord.Caption = "Record No :" & rsBrowse.AbsolutePosition & "/" & rsBrowse.RecordCount
End Sub

'Private Sub SetOptionStatus(blnStatus As Boolean)
'cmbWkT.Enabled = blnStatus
'End Sub


Private Sub dgMain_GotFocus()
lblTbActive = "INVOICE"
'SetOptionStatus (True)
tbrCommand.Buttons("CRITERIA").Enabled = True
tbrCommand.Buttons("PRINT").Enabled = True
End Sub

Private Sub dgDetail1_GotFocus()
lblTbActive = "INVOICEDETAIL"
gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
Call Set_Button(rsBrowse1)
tbrCommand.Buttons("CRITERIA").Enabled = False
tbrCommand.Buttons("PRINT").Enabled = False
''SetOptionStatus (False)
End Sub

Private Sub dgDetail2_GotFocus()
lblTbActive = "CUSTOMER"
gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
Call Set_Button(rsBrowse2)
tbrCommand.Buttons("CRITERIA").Enabled = False
tbrCommand.Buttons("PRINT").Enabled = False
'SetOptionStatus (False)
End Sub

Private Sub dgDetail3_GotFocus()
lblTbActive = "EMSRATE"
gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
Call Set_Button(rsBrowse3)
tbrCommand.Buttons("CRITERIA").Enabled = False
'tbrCommand.Buttons("PRINT").Enabled = False
tbrCommand.Buttons("PRINT").Enabled = True
'SetOptionStatus (False)
End Sub

Private Sub dgDetail4_GotFocus()
lblTbActive = "REGRATE"
gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
Call Set_Button(rsBrowse4)
tbrCommand.Buttons("CRITERIA").Enabled = False
tbrCommand.Buttons("PRINT").Enabled = False
'SetOptionStatus (False)
End Sub

Private Sub dgDetail5_GotFocus()
lblTbActive = "USERS"
gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
Call Set_Button(rsBrowse5)
tbrCommand.Buttons("CRITERIA").Enabled = False
tbrCommand.Buttons("PRINT").Enabled = False
'SetOptionStatus (False)
End Sub

'Private Sub dgDetail6_GotFocus()
'lblTbActive = "REGSERVICE"
'gsKeyHeader = DGetKeyToSelect("INVOICE", rsBrowse)
'Call Set_Button(rsBrowse6)
'tbrCommand.Buttons("CRITERIA").Enabled = False
'tbrCommand.Buttons("PRINT").Enabled = False
''SetOptionStatus (False)
'End Sub


Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "INVOICE"
Call DataGrid_HeaderClick(ColIndex, rsBrowse)
 End Sub

Private Sub dgDetail1_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "INVOICEDETAIL"
Call DataGrid_HeaderClick(ColIndex, rsBrowse1)
End Sub

Private Sub dgDetail2_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "CUSTOMER"
Call DataGrid_HeaderClick(ColIndex, rsBrowse2)
End Sub

Private Sub dgDetail3_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "EMSRATE"
Call DataGrid_HeaderClick(ColIndex, rsBrowse3)
End Sub

Private Sub dgDetail4_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "REGRATE"
Call DataGrid_HeaderClick(ColIndex, rsBrowse4)
End Sub

Private Sub dgDetail5_HeadClick(ByVal ColIndex As Integer)
'ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
lblTbActive = "USERS"
Call DataGrid_HeaderClick(ColIndex, rsBrowse5)
End Sub

'Private Sub dgDetail6_HeadClick(ByVal ColIndex As Integer)
''ทำการเรียงลำดับตาม Header ที่ click สลับกันระหว่าง มากไปน้อย และ น้อยไปมาก
'lblTbActive = "REGSERVICE"
'Call DataGrid_HeaderClick(ColIndex, rsBrowse6)
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsBrowse = Nothing
Set rsBrowse1 = Nothing
Set rsBrowse2 = Nothing
Set rsBrowse3 = Nothing
Set rsBrowse4 = Nothing
Set rsBrowse5 = Nothing
Set rsBrowse6 = Nothing
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lblTBActive_Change()
        tbActive = UCase(Trim(lblTbActive))
        Select Case tbActive
                      Case "INVOICE"
                                    lblTbActive_Des = "แฟ้มใบเสร็จ"
                      Case "INVOICEDETAIL"
                                    lblTbActive_Des = "รายการส่ง"
                      Case "CUSTOMER"
                                    lblTbActive_Des = "ลูกค้า"
                      Case "EMSRATE"
                                    lblTbActive_Des = "อัตราค่าส่ง และ ค่าบริการ แบบ EMS"
                      Case "REGRATE"
                                    lblTbActive_Des = "อัตราค่าส่ง และ ค่าบริการ แบบ ลทบ."
                      Case "USERS"
                                    lblTbActive_Des = "ผู้ใช้งาน(Users)"
        End Select
        '08/026/2011 เพื่อไม่ให้เกิด Error เนื่องจากจำ Column เดิม
        ColName = ""
        txtFindValue = ""
End Sub


Private Sub optAll_Click()
Call GetData("")
End Sub

'Private Sub GetData(strWKType As String)
'  Err_Desc = "" 'ทำการ Clear ให้เป็นว่างก่อน
'
' Set rsBrowse = New Adodb.Recordset
' If Trim(strWKType) <> "" Then
'    strCmdSQL = Find_strCmdSQLForBrowse(tbActive, "WKType='" & strWKType & "'")
'  Else
'    strCmdSQL = Find_strCmdSQLForBrowse(tbActive)
'  End If
'  rsBrowse.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'
'lblCntRecord.Caption = "Record No :" & rsBrowse.AbsolutePosition & "/" & rsBrowse.RecordCount
'
'  'กำหนด DataSource  และ Tag
'  Set dgMain.DataSource = rsBrowse
'  Call Define_Field_Tag(frmActive, rsBrowse)
'   'กำหนด  Colum Header  ให้กับ Data Grid ไม่มี cmborder
' 'Call Assign_HeaderGrid(rsBrowse, dgMain)
' Call Enable_tbrCommand
' Call Set_Button(rsBrowse)
'End Sub


'Private Sub optExternal_Click()
'Call GetData("E")
'End Sub

Private Sub optInternal_Click()
Call GetData("I")
End Sub

Private Sub stbDetail1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case stbDetail1.Tab
             Case 0
                        dgDetail1.SetFocus
             Case 1
                        dgDetail2.SetFocus
             Case 2
                        dgDetail3.SetFocus
             Case 3
                        dgDetail4.SetFocus
             Case 4
                        dgDetail5.SetFocus
'             Case 5
'                        dgDetail6.SetFocus
End Select
End Sub

Private Sub tbrExit_ButtonClick(ByVal Button As MSComctlLib.Button)
Unload Me
strCondition = ""
End Sub

Private Sub tbrSave_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "SAVE"
                        Call Save_Rec
                        Call Define_Field_Tag(Me, rsBrowse)
                        If bAddRecord Then 'ถ้า Add ได้สำเร็จ
                            SelectCondition = DGetKeyToSelect(tbActive, rsActive)
                        Else
                            SelectCondition = DGetKeyToSelect(tbActive, rsBrowse)
                        End If
                        sCriteria = SelectCondition
                        rsBrowse.Requery
                        Call Multi_Find(rsBrowse, sCriteria)
                        Call Refresh_Record(rsBrowse)
        'กรณีที่ยกเลิกคำสั่ง
        Case "CANCEL"
                        Call Cancel_Rec
                        Call Define_Field_Tag(Me, rsBrowse)
        End Select
        Call Enable_tbrCommand
End Sub



Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        Call Set_Button(rsBrowse)
        Call RowColChg
        Call BrowseDetail1(True)
        lblCntRecord.Caption = "Record No :" & rsBrowse.AbsolutePosition & "/" & rsBrowse.RecordCount
End Sub

Private Sub BrowseDetail1(blnClose As Boolean)
On Error Resume Next
        With frmActive
                  If Not rsBrowse.EOF Then
                      strCondition = " INVNo='" & Trim(rsBrowse!invno) & "'"
                  Else
                      strCondition = "1<>1"
                  End If
                  'BrowseDetail1
                  strCmdSQL = Find_strCmdSQLForBrowse("INVOICEDETAIL")
                  If blnClose Then rsBrowse1.Close
                  rsBrowse1.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                   'กำหนด DataSource  และ Tag
                   Set .dgDetail1.DataSource = rsBrowse1
                  'Call Assign_HeaderGrid(rsBrowse1, .dgDetail1)
                  
                  
                  'BrowseDetail2
                  strCmdSQL = Find_strCmdSQLForBrowse("CUSTOMER")
                  If blnClose Then rsBrowse2.Close
                  rsBrowse2.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                   'กำหนด DataSource  และ Tag
                   Set .dgDetail2.DataSource = rsBrowse2
                  'Call Assign_HeaderGrid(rsBrowse2, .dgDetail2)
                  
                  'BrowseDetail3
                  strCmdSQL = Find_strCmdSQLForBrowse("EMSRATE")
                  If blnClose Then rsBrowse3.Close
                  rsBrowse3.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                   'กำหนด DataSource  และ Tag
                   Set .dgDetail3.DataSource = rsBrowse3
                  'Call Assign_HeaderGrid(rsBrowse3, .dgDetail3)
                  
                  'BrowseDetail4
                  strCmdSQL = Find_strCmdSQLForBrowse("REGRATE")
                  If blnClose Then rsBrowse4.Close
                  rsBrowse4.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                   'กำหนด DataSource  และ Tag
                   Set .dgDetail4.DataSource = rsBrowse4
                  'Call Assign_HeaderGrid(rsBrowse4, .dgDetail4)
                  
                  'BrowseDetail5
                  strCmdSQL = Find_strCmdSQLForBrowse("USERS")
                  If blnClose Then rsBrowse5.Close
                  rsBrowse5.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                   'กำหนด DataSource  และ Tag
                   Set .dgDetail5.DataSource = rsBrowse5
                  'Call Assign_HeaderGrid(rsBrowse5, .dgDetail5)
                  
                  
      End With
End Sub


'==================================

Public Sub tbrCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSchItemOver7 As String
Dim strHasContract As String
Dim ConnectionString As String
        
        mode = Button.Key
        Select Case Trim(UCase(mode))
                Case "PRINT"
                        If lblTbActive = "INVOICE" Then
                                    Screen.MousePointer = vbHourglass
                                    With crptInvoice
                                           ' .Connect = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=alluser;PWD=alluser;Initial Catalog=" & CurrentDB & ";Data Source=" & CurrentServer
                                            '.Connect = "DSN = DILFx;UID = Nittaya;PWD=123;DSQ=PostSystem"
                                            '.LogOnServer "PDSODBC.DLL", "PostSystem", "PostSystem", "Nittaya", "123"
                                           ' .Connect = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=PostSystem;UID=Nittaya;PWD=1233"

                                            .Connect = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost:3300;Database=Postdb;user=root;Password=;"
                                                                 

                                            .ReportFileName = RptPath & "rptInvoice.rpt"
                                            .ParameterFields(1) = "เลขที่ Invoice;" & rsBrowse!invno & ";true"
                                            .PrintReport
                                    End With
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                       End If
                             
                Case "CRITERIA"
                              If tbActive <> "INVOICE" Then
                                 MsgBox "กรุณาเลือกที่แฟ้มหลักก่อนเลือก Criteria นะครับ.", vbCritical, "เลือกแฟ้มหลัก."
                              Else
                                  ' frmCriteria.Show vbModal
                              End If
                            Exit Sub
        End Select
        
        
        Select Case UCase(tbActive)
                      Case "INVOICE"
                                 Set frmActive = frmInvoice
                                 Call ActivateCommand_OldForm(mode, rsBrowse, Me)
                                If UCase(mode) = "DELETE" Then
                                    Call BrowseDetail1(True)
                                End If
                      Case "INVOICEDETAIL"
                                 Set frmActive = frmInvoiceDetail
                                 Call ActivateCommand_OldForm(mode, rsBrowse1, Me)

                      Case "CUSTOMER"
                                 Set frmActive = frmCustomer
                                 Call ActivateCommand_OldForm(mode, rsBrowse2, Me)
                      Case "EMSRATE"
                                 Set frmActive = frmEmsRate
                                 Call ActivateCommand_OldForm(mode, rsBrowse3, Me)
                      Case "REGRATE"
                                 Set frmActive = frmRegRate
                                 Call ActivateCommand_OldForm(mode, rsBrowse4, Me)
                      Case "USERS"
                                 Set frmActive = frmUsers
                                 Call ActivateCommand_OldForm(mode, rsBrowse5, Me)
                End Select



End Sub

Private Sub txtFindValue_Change()
        If Trim(txtFindValue) = "" Then Exit Sub
        Select Case tbActive
                      Case "INVOICE"
                                    Call SearchDataInGrid(rsBrowse)
                                    
                      Case "INVOICEDETAIL"
                                    Call SearchDataInGrid(rsBrowse1)
                                    
                       Case "CUSTOMER"
                                    Call SearchDataInGrid(rsBrowse2)
                                    
                      Case "EMSRATE"
                                    Call SearchDataInGrid(rsBrowse3)
                                    
                      Case "REGRATE"
                                    Call SearchDataInGrid(rsBrowse4)

                      Case "USERS"
                                    Call SearchDataInGrid(rsBrowse5)

                                    
         End Select

End Sub


Private Sub lblFirst_Click()
        Select Case tbActive
                      Case "INVOICE"
                                    Call First_Rec(rsBrowse)
                                    
                      Case "INVOICEDETAIL"
                                    Call First_Rec(rsBrowse1)
                                    
                       Case "CUSTOMER"
                                    Call First_Rec(rsBrowse2)
                                    
                      Case "EMSRATE"
                                    Call First_Rec(rsBrowse3)
                                    
                      Case "REGRATE"
                                    Call First_Rec(rsBrowse4)

                      Case "USRES"
                                    Call First_Rec(rsBrowse5)
         End Select

                                    
End Sub

Private Sub lblLast_Click()
        Select Case tbActive
                      Case "INVOICE"
                                    Call Last_Rec(rsBrowse)
                                    
                      Case "INVOICEDETAIL"
                                    Call Last_Rec(rsBrowse1)
                                    
                       Case "CUSTOMER"
                                    Call Last_Rec(rsBrowse2)
                                    
                      Case "EMSRATE"
                                    Call Last_Rec(rsBrowse3)
                                    
                      Case "REGRATE"
                                    Call Last_Rec(rsBrowse4)
                                    
                      Case "USRES"
                                    Call Last_Rec(rsBrowse5)
                                    
         End Select

End Sub

Private Sub lblNext_Click()
        Select Case tbActive
                      Case "INVOICE"
                                    Call Next_Rec(rsBrowse)
                                    
                      Case "INVOICEDETAIL"
                                    Call Next_Rec(rsBrowse1)
                                    
                       Case "CUSTOMER"
                                    Call Next_Rec(rsBrowse2)
                                    
                      Case "EMSRATE"
                                    Call Next_Rec(rsBrowse3)
                                    
                      Case "REGRATE"
                                    Call Next_Rec(rsBrowse4)
                                    
                      Case "USERS"
                                    Call Next_Rec(rsBrowse5)
                                    

         End Select

End Sub

Private Sub lblPrevious_Click()
        Select Case tbActive
                      Case "INVOICE"
                                    Call Previous_Rec(rsBrowse)
                                    
                      Case "INVOICEDETAIL"
                                    Call Previous_Rec(rsBrowse1)
                                    
                       Case "CUSTOMER"
                                    Call Previous_Rec(rsBrowse2)
                                    
                      Case "EMSRATE"
                                    Call Previous_Rec(rsBrowse3)
                                    
                      Case "REGRATE"
                                    Call Previous_Rec(rsBrowse4)
                                    
                      Case "USERS"
                                    Call Previous_Rec(rsBrowse5)
                                    
         End Select

End Sub




