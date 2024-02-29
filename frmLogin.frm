VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   285
      Left            =   1530
      TabIndex        =   4
      Top             =   1665
      Width           =   1170
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   990
      Width           =   1710
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1320
      TabIndex        =   2
      Top             =   435
      Width           =   1710
   End
   Begin VB.Image imgHandPoint 
      Height          =   480
      Left            =   3345
      Picture         =   "frmLogin.frx":094A
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Password :"
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   1095
      Width           =   780
   End
   Begin VB.Label lblUserID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "User ID :"
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   540
      Width           =   630
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New Adodb.Connection


Private Sub cmdLogin_Click()
'call Test_StoreProc
Dim strResult As String
    Set cmdSp = New Command
    With cmdSp
              .ActiveConnection = dbActive
              
            .CommandType = adCmdStoredProc
            .CommandText = "check_login"

            .Parameters.Append cmdSp.CreateParameter("struserid", adChar, adParamInput, 10)
            .Parameters("struserid").Value = Trim(txtUserID.Text)


            .Parameters.Append cmdSp.CreateParameter("strpassword", adChar, adParamInput, 15)
            .Parameters("strpassword").Value = Trim(txtPassword.Text)

            .Parameters.Append cmdSp.CreateParameter("strResult", adVarChar, adParamOutput, 50)
            .Execute

               strResult = cmdSp.Parameters("strResult").Value
     End With

     Set cmdSp = Nothing
    If UCase(Trim(strResult)) = "FOUND" Then
                UsrSTFCode = Trim(txtUserID.Text)
               frmPostSystem.Show vbModal
        Else
                MsgBox strResult, vbCritical, "Try again"
    End If


'If Count_Record("users", "userid='" & Trim(txtUserID.Text) & "'") = 0 Then
'     MsgBox "User Not Found", vbCritical, "Try again"
'Else
'     If Count_Record("users", "userid='" & Trim(txtUserID.Text) & "' and password='" & Trim(txtPassword.Text) & "'") <> 0 Then
'            UsrSTFCode = Trim(txtUserID.Text)
'           frmPostSystem.Show vbModal
'    Else
'            MsgBox "Wrong Password", vbCritical, "Try again"
'    End If
'End If
End Sub

Private Sub Form_Load()

   txtUserID.Text = "admin"
   txtPassword.Text = "admin"
    
End Sub
