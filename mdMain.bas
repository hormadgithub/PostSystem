Attribute VB_Name = "mdMain"

'Sub Main()
'On Error Resume Next
'
'          CurrentServer = "localhost"
'          Set dbActive = New ADODB.Connection
'          Set dbSQL = New ADODB.Connection
'
'        strConnectionDB = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=PostSystem;user=Nittaya;Password=123;"
'        dbSQL.ConnectionString = strConnectionDB
'        dbSQL.CursorLocation = adUseClient
'        dbSQL.Open
'        Set dbActive = dbSQL
'
'
'' -----------------------------
'' Change format date.
'        Call CH_DateFormat
'        '�ӡ���Դ Messenger Service  ��� Start Auto
''        Call EnableMessengerService
'        MaskDate = "../../...."
'       CurrentDate = CStr(Format(Date, Fdate))
'        Set rsActive = New ADODB.Recordset
'        frmLogin.Show
'        'frmMain.Show
'        Set rsActive = Nothing
'End Sub



'14-01-2022 ¡��ԡ Server ���
'Public Const CurrentServer = "MARIA"
'14-01-2022 �������ҹ Server ����
Public Const CurrentServer = "localhost"

Public strConnectionDB  As String
Public strDefaultConnectionDB  As String

'========== Service System ===================
Public rsSaveBrowse As Adodb.Recordset
Public strMtvKey As String '�ç�Ѻ MtvKey �� Menutv

'����ǡѺ Detail Credit Limit �ͧ �١���
Public CurCRAmt As Double  '�ӹǹ�Թ���������㹻Ѩ�غѹ�ͧ�١�����¹��
Public SumCurCRAmt As Double  '�ӹǹ�Թ���������㹻Ѩ�غѹ�ͧ�١��ҡ�������
Public CRLMAmt As Double '�ӹǹ Cremit Line �ͧ�١��Ҥ������ Table
Public SUMCRLMAmt As Double '�ӹǹ Cremit Line ����ͧ�١��ҡ�������
Public ARAMT As Double '�ӹǹ�Թ����ҧ���к� Dos �ͧ�١�����¹��(DoAmt ���ҧ����)
Public SUMARAmt As Double '�ӹǹ�Թ����ҧ���к� Dos �ͧ�١��ҡ�������
Public OverCRLM As Double '�ӹǹ�Թ����Թ�ͧ�١�����¹��
Public AdvChqAmt As Double '���Թ��������������ǧ˹�����ͧ�١�����¹��
Public SumAdvChqAmt As Double '���Թ��������������ǧ˹�����ͧ�١��ҡ�������

Public DocNoBfDel1 As String, DocNoBfDel2 As String  '���Ţ����͡��á�͹�ӡ�� Delete �����������ѧ Delete ����
''����ǡѺ Customer Status
Public CSRemark As String
Public BlackList As String * 1 '�١��ҷ���ջѭ��
Public CSblkLstrem As String ' �����˵ػѭ��
Public CSTerms As Integer '��Ҩҡ Customer
Public CountBOR As Integer '�ӹǹ�����ͧ�١�����¹��
Public StdAvgDisc As Double '��ǹŴ�ҵ�Ұҹ����Ἱ������ҡѹ


Public CurrentCriteria As String
Public sCriteria As String
Public StrOrder As String
Public sKey As String
Public SelectCondition As String
Public ColName As String
Public strSort As String '���§�ӴѺ ASC OR DESC
Public savFrmActive As Form

Public savTBName As String
Public savColname  As String

'
'''*** 2. LocaleInfo
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Const LOCALE_SSHORTDATE = &H1F
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


'========== Service System ===================

'== ���Ǩ�ͺ��觷������㹡�������ǡѹ ===========
Public strGroupSalesTB As String
Public strGroupTarget As String
Public strGroupT20Cust As String
Public strGroupFunnel As String

Public blnGroupSalesTB As Boolean
Public blnGroupTarget As Boolean
Public blnGroupT20Cust As Boolean
Public blnGroupFunnel As Boolean
'============================================


Public blnAllowRowColChange As Boolean

Public strRcvCtrlCondition As String
Public bSalesDept As Boolean '�͹ Logon �������Ἱ� Sales �������
Public MouseX, MouseY, i
Public Const FormTop = 1100
Public Const sCondSalesDeptNotBranch = "DPsale = 'Y' And DPCode<>'CHO' and  DPCode<>'RAY'"


'������Ѻ�纡����Ѻ Item
Public Save_Item As Integer '�������͹���º��º㹡����Ѻ Itemno
Public Edit_Item As Integer '�������͹���º��º㹡����Ѻ Itemno
Public blnSwap_Itemno As Boolean

'
''Level For Report Menu
Public Enum ERptMnuLevel '������ѡ
                        Menu = 0
                        SubMenu = 1
                        Opt = 2
             End Enum
'Color For Report Menu
Public Enum EMnuColor '������ѡ
                        Default = vbBlack '�͹�á��
                        Active = vbWhite  '�͹Mouse Over
                        Choice = vbYellow '�͹���͡����
             End Enum

Public Enum ESubMnuColor '��������
                        Default = vbWhite '�͹�á��
                        Active = &HF3DEAD '&HFF8080 '�͹Mouse Over
                        Choice = vbYellow '�͹���͡����
             End Enum

Public Enum EOptColor 'Option ���������͡
                        Default = vbWhite '�͹�á��
                        Active = &HF3DEAD    '&HFF8080 '�͹Mouse Over
                        Choice = vbYellow '�͹���͡����
             End Enum
Public Const FrameRptForeColor = vbBlue

'=============================================================
        '�� DataBase ����� Current Connection
        Public CurrentDB As String
        Public blnDataTest As Boolean
        Public OldDB As String
        
        '������� Database AccountDB �������
        Public AccountDB  As Boolean
         
        '������� Database AccountDB_Old ������� ���������Ѻ�٢����Żա�͹˹��
        Public AccountDB_Old  As Boolean
         
        '��᷹����� SQL
        Public strCmdSQL As String
        Public strTempolary As String
       
    
        '��᷹ Control � Form
        Public Ctrl As Control
    
        '��᷹ Usr ��� Logon ����к�
        Public UsrLevel As String * 1   'Level
        Public UsrDept As String * 3    'Department
        Public UsrBranch As String * 1    'Branch
        Public UsrPwd As String         'Password
        Public UsrSCode As String * 6   'Sale Code
        Public UsrPSTCode As String * 2         'Position Code
        Public UsrSTFCode As String * 5         'Position Code
        Public UsrName As String                        '���ͷ������� Staff Code (�������ͷ����㹡�� Login)
        Public AreaCondition As String                  '�纾�鹷��ͧ Sales ����鹷�����
        Public UsrIsMarketing As String * 1        '�����������¢����������
        Public bSalesActive As Boolean   '����� Sales Active ��������
        Public sConditionRefreshRecord  As String   '���͹�㹡�� Refresh Record
        ''# Group Of Pneumax User
        Public bGP_Mis As Boolean
        Public bAdminLogon As Boolean   '������Ѻ�����ٷ���͹
        Public bGP_Project As Boolean
        
        Public bGP_SalesRep As Boolean
        Public bGP_Support As Boolean
        Public bGP_SupportMgr As Boolean
        Public bGP_SalesUmgr As Boolean
        Public bGP_SalesMgr As Boolean
        Public bGP_Director As Boolean
        Public bGP_Manager As Boolean
        Public bGP_Sales As Boolean
        
        Public bGP_Import As Boolean
        Public bGP_MatMgr As Boolean
        Public bGP_Sec As Boolean

'===============================================================

'�����͹�㹡�����͡���������Ѻ Sales �� CRO ,QTT
Public SalesCondition As String
Public bExitSub  As Boolean '������Ѻ�͡�����ѧ��Ѻ�Ҩҡ������¡�� StoreProc ���� ��� Exit Sub �������

Public StrBrowseField As String, StrJoinTable As String, StrTBCondition As String   '�繵���������纤�� Field �ҡ��� AllTable

 Private Sale_Stack(100, 2) As String '��� Function GetSaleCondition
 Private top_stack As Integer
 Private SaleCondition, GetBoss As String

Public fraTopColor As ColorConstants    '�բͧ Frame ���ǹ��Ǣͧ Form
'Public dbFox As ADODB.Connection '������� Foxpro
Public dbSQL As Adodb.Connection '������� SQL
Public dbActive As Adodb.Connection  '����ҵ�ͧ��� Connect �Ѻ Database ����˹
Public dbTemp As Adodb.Connection  '��㹡�õԴ��͢��� DataBase

'Public frmPrint As String
Public odbcSQL As Adodb.Connection '������� SQL ��ҹ ODBC
Public fields As Adodb.Field

Public SysLogin As String '����������ӡ�����͡�͹ Login (S-Service P-Pneumax)
Public cmdSp As Adodb.Command
Public SaleDept As String '��Ἱ��ͧ Sale ������
Public DirGrpDept As String '��Ἱ���� Director ����鹴�������


'Public MenuLevel, MenuName As String '��Ǻ������٢ͧ User
Public HideField As String '������Ѻ�Ǻ��� Field ����ͧ�������ʴ�
Public CntCol As Integer '��Ѻ�ӹǹ Colum �ͧ Data Grid ��͹ź Colum �ͧ Data Grid

Public rsBrowse As Adodb.Recordset '���� Record Set �͹ Browse
Public rsBrowse1 As Adodb.Recordset '���� Record Set �͹ Browse Detail1
Public rsBrowse2 As Adodb.Recordset '���� Record Set �͹ Browse Detail2
Public rsBrowse3 As Adodb.Recordset '���� Record Set �͹ Browse Detail3
Public rsBrowse4 As Adodb.Recordset '���� Record Set �͹ Browse Detail4
Public rsBrowse5 As Adodb.Recordset '���� Record Set �͹ Browse Detail5
Public rsBrowse6 As Adodb.Recordset '���� Record Set �͹ Browse Detail6

Public rsActive As Adodb.Recordset '���繵�ǡ�ҧ㹡�äǺ��� Adodb.Recordset ������
Public rsTemp  As Adodb.Recordset '���� Record Set Temp
Public rsFunction  As Adodb.Recordset '���� Record Set Temp � Function
Public rsClone As Adodb.Recordset
Public frmActive As Form
Public RecBfRefresh As String   '���˹� record ��͹�ӡ�� Add ,Edit
Public HaveDetail As Boolean '��͡����ա�û�͹ Detail ���������ѧ

Public blnDocNoInRcvCtrlHaveRef  As Boolean '��͡����ա����ҧ�ҡ�����������������

Public tbActive As String
Public tbActual As String 'Table ����������ԧ����� Table ����� View
Public tbActualPrevious As String 'Table ��͹˹��

Public fldOrder As String
Public CurrentTable As String
Public bBrowseState As Boolean
Public dgMainCaption As String
Public strCurrentNodeClick As String 'Node �Ѩ�غѹ��� Click
'�� Caption �ͧ lblMovetxt �������觵͹���ͧ���п����
Public MovetxtCaption As String

'�� Field ����� Key ��������ҵ͹ Refresh
Public FldRefresh As String
Public KeyActive As String
Public KeyValue(9) As String

''# ��᷹ ValBfClick �ͧ Activate Node Select
Public gsKeyHeader As String

Public strCondition As String  '������Ѻ�����͹���ѧ Where 㹡�� Check DupKey
                                                          'Find_Rec ��� Refresh
Public strRptCondition As String  '������Ѻ�����͹���ѧ Where 㹡�� Check DupKey
                                                          'Find_Rec ��� RefreshRpt_Rec �ͧ Report ��ҧ�
Public HeaderCondition As String '�� Condition �ͧ Header

Public FindCondition As String  '������Ѻ�����͹���ѧ Where ��͹���¡�� Lookup
Public LookupCondition As String  '������Ѻ�����͹���ѧ Where � Lookup
Public LookupColumn As String  '������Ѻ�����͹���ѧ Where � Lookup
Public countCondition As String '������Ѻ�Ѻ�ӹǹ Record
Public MaskDate As String, CurrentDate As String
Public SelectNumRec   As String '�ӹǹ Record ����ͧ�������ʴ�

'Constant Variable
Public RptPath As String   '�� Path �ͧ Report �͹���¡�����
Public PicPath As String, PhotoPath As String, SignaturePath As String, ImgPartPath As String, CustRemPath As String     '�� Path �ͧ�ٻ�Ҿ�͹�ʴ��͡˹�Ҩ� ,�ٻ��ѡ�ҹ,����繵� ����ӴѺ
Public CustMapPath As String   '�� Path �ͧἹ���ͧ�١���
Public CustPOPath As String   '�� Path �ͧ PO �١���
Public FInteger As String, FReal As String, FMoney As String, Fdate As String, FDateTime As String    '��˹� Format �ͧ����Ţ �ѹ���
Public FCurrency As String '�ȹ���  4 ��ѡ
Public FExchgRate As String '�ȹ���  5  ��ѡ
Public VatRate As String
Public LpoFactor  As Double '��Ǥٳ����Ѻ�Թ��ҫ����Ң���
Public INCService As Double '��Ǥٳ����Ѻ��� Incentive �Ѻ Service
Public MultiFactor As Double '��Ǥٳ����ѺPartNid ���Ѵ����繷�͹ �
'Part Type ������������ ListPrice
Public EditPartUPrice As String
'Part Type �������ˢ���Թ�Ҥҵ����
Public OverPartUPrice As String
'Part Type ������������� PartNo
Public EditPartNo As String
'Part Type ������������ PartDescription
Public EditPartDesc As String
'Part Type ����ͧ�ӡ�� Check On Hand
Public PartChkOnhand As String
'�����ʴ� Error  Message ��� User ���
Public ShowErrMsg As String
'���������ö ����ͧ���
Public dblDepreciation As Double
'================================
Public intCnt  As Integer

' Declear Variable For Lookup
Public ActiveLookup As String
Public rsFind As Adodb.Recordset
Public rsCtrlRec As Adodb.Recordset
Public rsLookup As Adodb.Recordset
Public rsFromTB As Adodb.Recordset
'Return value from LookUp Form
Public LookupRetVal As String
Public iLookup(10) As String

Public CancelValidate As Boolean '������Ѻ��Ǩ�ͺ��Ҩ�����ҹ Field ����դ�����١��ͧ�������
Public CancelLookup As Boolean '���Ǩ�ͺ��� ���� Cancel �͹ Lookup
Public LookupTitle As String

'Public qttTbActive As String
Public mode As String '������繡�� ADD  ���� EDIT
Public bAddRecord As Boolean
Public strEvents As String '������繡�� ADD  ���� EDIT
Public SaveMode As String '�������������Ѻ���� Mode ��� �͹��˹������ Mode ���

'�纤�������˵ط���觤׹�ҡ Form Remark
Public RemarkValue As String
'��Ǩ�ͺ������� Record �ͧ Header �� Cro,Quotation ��Ҽ�ҹ��õ�Ǩ�ͺ�������öź���������
Public CheckDel As Boolean

'����������´����觷��ӧҹ���� ERROR
Public Err_Desc  As String


'### Sep 2, 2006 ###
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1

Public JobCtrlCircuitPath As String

'31-05-2017 �ӡ���Դ˹�� Form ��� Auto
Public blnOpenFormAuto As Boolean


Sub Main()
On Error Resume Next
            RptPath = "D:\PostSystem\Report\"
          Set dbActive = New Adodb.Connection
          Set dbSQL = New Adodb.Connection

        strConnectionDB = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=Postdb;user=root;Password=;"

        dbSQL.ConnectionString = strConnectionDB
        dbSQL.CursorLocation = adUseClient
        dbSQL.Open
        Set dbActive = dbSQL


' -----------------------------
' Change format date.
        Call CH_DateFormat
        '�ӡ���Դ Messenger Service  ��� Start Auto
'        Call EnableMessengerService
        MaskDate = "../../...."
       CurrentDate = CStr(Format(Date, Fdate))
        Set rsActive = New Adodb.Recordset
        frmLogin.Show
        'frmMain.Show
        Set rsActive = Nothing
End Sub

Public Function fn_GetRate(strSendType As String, intWeight As Integer) As Integer
Dim intRate As Integer
        strCondition = "minweight<=" & Trim(intWeight) & " and  maxweight>=" & Trim(intWeight)
        If strSendType = "EMS" Then
            intRate = Find_Ret_Num("EMSRATE", "emsprice", strCondition)
        Else
            intRate = Find_Ret_Num("REGRATE", "regprice", strCondition)
        End If
        fn_GetRate = intRate
End Function

Public Function fn_GetService(strSendType As String, intWeight As Integer) As Integer
Dim intService As Integer
        strCondition = "minweight<=" & Trim(intWeight) & " and  maxweight>=" & Trim(intWeight)
        If strSendType = "EMS" Then
            intService = Find_Ret_Num("EMSRATE", "emsservice", strCondition)
        Else
            intService = Find_Ret_Num("REGRATE", "regservice", strCondition)
        End If
        fn_GetService = intService
End Function

'
'Public Sub AssingConstantVar()
''��˹� Path �ҡ Tabel ConstVar ���¡��͹ Login ������� Convar �ͧ Database ���١���
' Set rsFunction = New Adodb.Recordset
'  With rsFunction
'        strCmdSQL = "select  * from ConstVar "
'  .Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
'  Do While Not .EOF
'        Select Case Trim(rsFunction!CVCode)
'                 Case "CUSTMAP"
'                            CustMapPath = rsFunction!CVValue
'                Case "CUSTPO"
'                            CustPOPath = rsFunction!CVValue
'                Case "IMGPARTPATH"
'                            ImgPartPath = rsFunction!CVValue
'                Case "PHOTOPATH"
'                            PhotoPath = rsFunction!CVValue
'                Case "SIGNATUREPATH"
'                            SignaturePath = rsFunction!CVValue
'                Case "PICTURE"
'                            PicPath = rsFunction!CVValue
'                Case "REPORT"
'                            RptPath = rsFunction!CVValue
'                Case "CUSTREMARK"
'                            CustRemPath = rsFunction!CVValue
'                Case "FORMATINTEGER"
'                            FInteger = rsFunction!CVValue
'                Case "FORMATREAL"
'                            FReal = rsFunction!CVValue
'                Case "FORMATMONEY"
'                            FMoney = rsFunction!CVValue
'                Case "FORMATCURRENCY"
'                            FCurrency = rsFunction!CVValue
'                Case "FORMATEXCHGRATE"
'                            FExchgRate = rsFunction!CVValue
'                Case "FORMATDATE"
'                            Fdate = rsFunction!CVValue
'                Case "FORMATDATETIME"
'                            FDateTime = rsFunction!CVValue
'                Case "VATRATE"
'                            VatRate = rsFunction!CVValue
'                Case "LPOFACTOR"
'                            LpoFactor = CDbl(rsFunction!CVValue)
'                Case "INCSERVICE"
'                            INCService = CDbl(rsFunction!CVValue)
'                Case "MULTIFACTOR"
'                            MultiFactor = CDbl(rsFunction!CVValue)
'                Case "EDITPARTUPRICE"
'                           EditPartUPrice = rsFunction!CVValue
'                Case "OVERPARTUPRICE"
'                            OverPartUPrice = rsFunction!CVValue
'                Case "EDITPARTNO"
'                            EditPartNo = rsFunction!CVValue
'                Case "EDITPARTDESC"
'                            EditPartDesc = rsFunction!CVValue
'                Case "PARTCHKONHAND"
'                            PartChkOnhand = rsFunction!CVValue
'                Case "SHOWERRMSG"
'                            ShowErrMsg = rsFunction!CVValue
'                Case "DEPRECIATION"
'                            dblDepreciation = rsFunction!CVValue
'                  '�繡�����ͧ�������оǡ���������������º��º����ѧ
'                Case "STRGROUPSALESTB"
'                            strGroupSalesTB = rsFunction!CVValue
'                Case "STRGROUPTARGET"
'                            strGroupTarget = rsFunction!CVValue
'                Case "STRGROUPT20CUST"
'                            strGroupT20Cust = rsFunction!CVValue
'                Case "STRGROUPFUNNEL"
'                            strGroupFunnel = rsFunction!CVValue
'        End Select
'        .MoveNext
'  Loop
'  .Close
'  End With
'  Set rsFunction = Nothing
'If Left(UCase(CurrentDB), 9) = "ACCOUNTDB" Then
'        RptPath = RptPath + Trim(UCase(CurrentDB)) + "\"
'ElseIf UCase(CurrentDB) = "DATATEST" Then
'                 RptPath = RptPath + "DATATEST\"
'End If
'End Sub

'������Ѻ��ô֧ Form �������
Public Sub Add_Rec()
         With rsActive
                  .AddNew
         End With
' '�Ԩ�óҡ�͹��Ҩ��� Runno �������
' '������� Running ����͹ 3
'  If CurrentDate >= CDate("01/03/2006") Then
'         With frmActive
'                 '������� Table ����� Runno ���� ��� ����� AccountDB ���ӡ�á�˹� DocNo �ͧ��
'                 If Find_Ret_Val("RUNNO", "RNtbName", "RNtbname='" & UCase(tbActive) & "'") <> "" And Not AccountDB Then
'                         Set rsFunction = New Adodb.Recordset
'                            With rsFunction
'                                      Select Case UCase(tbActive) '㹡óշ���¡Ἱ�
'                                                   Case "WKORDER", "CONTRACT"
'                                                                strCmdSQL = ("select *  from Runno where RNtbname = '" & tbActive & "' and DPCode='" & UsrDept & "'")
'                                                    Case Else
'                                                                strCmdSQL = ("select *  from Runno where RNtbname = '" & tbActive & "'")
'                                      End Select
'                                .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'                                If Not .EOF Then
'                                    If Month(Now) <> CInt(rsFunction!RNmonth) Then
'                                       MsgBox "����������¹ Running No. �ͧ��͹������ǹФ�Ѻ.", vbInformation, "Change Running No."
'                                    End If
'                                    .Close
'                                 Else
'                                    MsgBox "This Table Not Use Running No OR You Dept Not Have RunNo.", vbCritical, "Error"
'                                    Call Cancel_Rec
'                                    Exit Sub
'                                 End If
'                             End With
'                             Set rsFunction = Nothing
'                End If
'    End With
' End If
   mode = "ADD"
   MovetxtCaption = "���ѧ�ӡ��������������� " & tbActive & "..."
   
        With frmActive
                .lblMovetxt.Left = .Width
                .tmMovetxt.Enabled = True
                .lblMovetxt.Caption = MovetxtCaption   '�繵����� Module Main
                '.lblMovetxt.ForeColor = vbBlue
                .Caption = frmActive.Caption & "(" & mode & ")"
                .Show vbModal
        End With
        Unload frmActive
End Sub

Public Sub Edit_Rec(ByRef rsSend As Adodb.Recordset)
         On Error GoTo Err_Edit
         If rsActive.AbsolutePosition > 0 Then '���㹡óշ����������͡ Record
                  mode = "EDIT"
                  MovetxtCaption = "���ѧ�ӡ����䢢�������� " & tbActive & "..."
                  With frmActive
                           If AccountDB Then
                                    .txtfld1.Enabled = True
                                    If Not (TypeOf .txtfld1 Is MaskEdBox) Then .txtfld1.Locked = False
                           Else
                                    .txtfld1.Enabled = False
                           End If
                           .lblMovetxt.Left = .Width
                           .tmMovetxt.Enabled = True
                           .lblMovetxt.Caption = MovetxtCaption '�繵����� Module Main
                           .Caption = frmActive.Caption & "(" & mode & ")"
                           .Show vbModal
                  End With
                  Unload frmActive
                  
         Else
                  MsgBox "Not found record for edit.!" & Err.Description, vbOKOnly + vbCritical, "Not found"
                  
         End If
         
Exit Sub '����������价Ӥ���觢�ҧ��ҧ

Err_Edit:
         MsgBox "Please select record for edit.!", vbOKOnly + vbCritical, "Not  found"
         MsgBox Err.Description
         Call Last_Rec(rsSend)
         
End Sub

Public Sub Delete_Rec(ByRef rsSend As Adodb.Recordset)
        Dim sSQL As String
        Dim sCond As String
        Dim dblSumAmt1 As Double
        Dim dblSumAmt2 As Double
        Dim dblSumAmt3 As Double
        
        On Error GoTo Err_Delete
        If rsActive.AbsolutePosition > 0 Then '���㹡óշ����������͡ Record
                If rsActive.RecordCount = 0 Then
                        MsgBox "Not found record for delete. !", vbOKOnly + vbCritical, "Not found"
                        Exit Sub
                End If
                mode = "DELETE"
                Select Case tbActive
                            Case "WKISSUEOILDT"
                                            sCond = "WIONo=  '" & rsActive!wiono & "'"
                End Select
                               
               With rsActive
                        .Delete  'ź�����Ţͧ rsActive ��� Main
                End With 'End with rsactive
                Select Case tbActive
                            Case "WKISSUEOILDT"
                                        dblSumAmt1 = Format(Sum_Item("WKISSUEOILDT", "WKOilAmt", sCond), FReal)
                                        dblSumAmt2 = Format(Sum_Item("WKISSUEOILDT", "WKExpressAmt", sCond), FReal)
                                        dblSumAmt3 = Format(Sum_Item("WKISSUEOilDT", "WKAllowanceAmt", sCond), FReal)
                                        ' �ӡ�� update �ӹǹ�Թ���� Detail ���Ѻ Header
                                        dbSQL.Execute "Update WKIssueOil  Set TotOilAmtInDatail= " & dblSumAmt1 & " , TotExpressAmtInDatail = " & dblSumAmt2 & " , TotAllowanceAmtInDatail = " & dblSumAmt3 & " Where " & sCond
                End Select
       
       End If
      mode = ""
    Exit Sub
Err_Delete:
        MsgBox "Please select record  for delete. !", vbOKOnly + vbCritical, "Not  found"
        MsgBox Err.Description
        Call Last_Rec(rsSend)
        RecBfRefresh = rsActive.fields(FldRefresh)
        Exit Sub
End Sub


'��Ǩ�ͺ�Է���㹡�� �Է��㹡���� ,ADD EDIT ���� DELETE
Public Function Check_Permission(strMtvKey As String) As Boolean
         Dim sCond As String
         Dim i As Integer
         
         Check_Permission = False
        sCond = "MtvKey = '" & strMtvKey & "' And MtvLevel Like '%" & UsrLevel & "%' And MtvShareDept Like '%" & UsrDept & "%' "
        i = DCount_Record("MenuTV", sCond)
        If i > 0 Then
            Check_Permission = True
        Else
            MsgBox "�س������Է�������ٹ��Ф�Ѻ.", vbCritical, "Not Permission."
        End If
End Function


Public Function Check_Add(ByRef rsSend As Adodb.Recordset) As Boolean
         Dim sCond As String
         Dim i As Integer
         
         Check_Add = True

'         '��Ǩ�ͺ ���������ö�ӡ�� Add ��
'         sCond = "TBName = '" & tbActive & "'"
'        i = DCount_Record("MenuTV", sCond)
'         If i > 0 Then
'                Check_Add = True
'                  Select Case UCase(tbActive)
'                                 Case "WKREQUEST", "WKPROBLEM", "WKCHKLIST", "WKRPLIST"
'                                              Check_Add = Count_Record("WKOrder", "WKNo='" & rsBrowse!WKNo & "' And ( rtrim(WKStatus)='' OR rtrim(WKStatus)='O' OR rtrim(WKStatus)='W' ) and ManHourCost=0") <> 0
'                                 Case "WKSERVICE", "WKACTIVITY"
'                                              Check_Add = Count_Record("WKOrder", "WKNo='" & rsBrowse!WKNo & "' And ( rtrim(WKStatus)='' OR rtrim(WKStatus)='O') OR rtrim(WKStatus)='W' ") <> 0
'
'                                Case "CONTRACTSUBDETAIL"
'                                              Check_Add = (Trim$(rsBrowse!SMGRApprCode) = "")
'
'                  End Select
'            End If
End Function

Public Function Check_Edit(ByRef rsSend As Adodb.Recordset) As Boolean
         Dim sCond As String
         Dim i As Integer
         Check_Edit = False
          If rsSend.BOF Or rsSend.EOF Then Exit Function
         Check_Edit = True


'
'         sCond = "TBName = '" & tbActive & "' And MtvEditLevel Like '%" & UsrLevel & "%' And MtvEditDept Like '%" & UsrDept & "%' "
'         i = DCount_Record("MenuTV", sCond)
'         If i > 0 Then Check_Edit = True
End Function

Public Function Check_Delete(ByRef rsSend As Adodb.Recordset) As Boolean
         Dim sCond As String
         Dim i As Integer
         Dim intCount As Integer
         Dim strReturn As String
         
         Check_Delete = False
         If (rsSend.BOF Or rsSend.EOF) Then Exit Function
         Check_Delete = True

'
'         sCond = "TBName = '" & tbActive & "' And MtvDeleteLevel Like '%" & UsrLevel & "%' And MtvDeleteDept Like '%" & UsrDept & "%' "
'         i = DCount_Record("MenuTV", sCond)
'         If i > 0 Then
'                  Check_Delete = True
'                  Select Case UCase(tbActive)
'                                Case "WKISSUEOIL" '��� Account �Ŵ Check �������ӡ����䢨��������öź��
'                                              sCond = "WIONO='" & rsBrowse!wiono & "'"
'                                              Check_Delete = Trim(Find_Ret_Val("WKISSUEOIL", "WIOCHECK", sCond)) = "" And Count_Record("WKIssueOil", sCond & "  And WIOWaitEdit=''") <> 0
'
'                                Case "WKISSUEOILDT"  '��һŴ Check ��������ö�����¡����
'                                              sCond = "WIONO='" & rsBrowse!wiono & "'"
'                                              Check_Delete = Trim(Find_Ret_Val("WKISSUEOIL", "WIOCHECK", sCond)) = ""
'
'
'                                Case "CONTRACTSUBDETAIL"
'                                              Check_Delete = Trim$(rsBrowse!SMGRApprCode) = ""
'
'
'                                Case "SVTYPE"
'                                              Check_Delete = Count_Record("WKORDER", "SVTCode='" & Trim$(rsBrowse!SVTcode) & "'") = 0
'                                Case "MADE"
'                                              Check_Delete = Count_Record("MODEL", "MadeCode='" & Trim$(rsBrowse!madecode) & "'") = 0
'                                Case "MODEL"
'                                              Check_Delete = Count_Record("MODELDT", "MadeCode='" & Trim$(rsBrowse1!madecode) & "' and ModCode='" & Trim$(rsBrowse1!ModCode) & "'") = 0 And _
'                                                                              Count_Record("CHECKLIST", "MadeCode='" & Trim$(rsBrowse1!madecode) & "' and ModCode='" & Trim$(rsBrowse1!ModCode) & "'") = 0 And _
'                                                                              Count_Record("REPAIRLIST", "MadeCode='" & Trim$(rsBrowse1!madecode) & "' and ModCode='" & Trim$(rsBrowse1!ModCode) & "'") = 0
'
'                                Case "CONTRACT", "CONTRACTDT", "CLAIMTB"
'                                              Check_Delete = Count_Record("WKOrder", "CTNO='" & Trim(rsBrowse!CTNO) & "'") = 0 And Trim$(rsBrowse!CTCancel) = "" And Trim$(rsBrowse!CTComplete) = "" And Trim$(rsBrowse!CTCheckCode) = "" And Trim$(rsBrowse!CustNotReqCTCode) = ""
'
'                                Case "SCHEDULETIME"
'                                              Check_Delete = IsNull(rsBrowse2!schedate)
'
'                                   '���������ź����ա�õ���ԡ����
'                                 Case "WKORDER", "WKREQUEST", "WKPROBLEM", "WKCHKLIST", "WKRPLIST"
'                                              Check_Delete = Count_Record("WKOrder", "WKNo='" & rsBrowse!WKNo & "' And ( rtrim(WKStatus)='' OR rtrim(WKStatus)='O' OR rtrim(WKStatus)='W') and WKCheck=''  and WIONO=''") <> 0
'                                              '30/09/2010 �������öź�������١����
'                                               If UCase(tbActive) = "WKORDER" Then
'                                                    Check_Delete = Check_Delete And Count_Record("WKOrder", "MainWKNo='" & rsBrowse!WKNo & "' And WKNo<>'" & rsBrowse!WKNo & "'") = 0
'                                               End If
'
'                                 Case "WKSERVICE", "WKACTIVITY"
'                                              Check_Delete = Count_Record("WKOrder", "WKNo='" & rsBrowse!WKNo & "' And ( rtrim(WKStatus)='' OR rtrim(WKStatus)='O' OR rtrim(WKStatus)='W') ") <> 0
'
'                  End Select
'         End If
End Function


'������Ѻ��Ǩ�ͺ�ѹ�������� Null ���������������ҡѺ Control � Form CROTODR
Public Function Assign_DateToCtrl(refCtrl As Control, dtp As DTPicker) As Boolean
        Dim strTmp As String
        Dim intTmp As String
        Dim strDate As String
        'Convert �.�. to �.�.
        
        If Trim(refCtrl) <> "" And Not IsDate(refCtrl) Then
                GoTo Error_Date
        End If
        If Trim(refCtrl) <> "" Then
                strTmp = Left(Right(refCtrl, 4), 2)
                If Left(strTmp, 1) = "/" Then
                        GoTo Error_Date
                End If
                If strTmp = "25" Or strTmp = "24" Then
                        If Len(Trim(refCtrl)) < 8 Then
                                GoTo Error_Date
                        End If
                        If strTmp <> "/" Or strTmp <> "0" Then
                                intTmp = CInt(Right(refCtrl, 4)) - 543
                                strTmp = CStr(intTmp)
                                refCtrl = Left(refCtrl, Len(refCtrl) - 4) & strTmp
                        End If
                End If
        End If
        
        'Assign date to DTPicker
        strDate = IIf(IsDate(refCtrl), Format(refCtrl, Fdate), "")
        If strDate = "" Then
                dtp.Value = CurrentDate
                refCtrl = strDate
                'MsgBox "Error date: Pleas verify data again.", vbInformation, "Error Data"
        Else
                dtp.Value = strDate
        End If
        Assign_DateToCtrl = False
        Exit Function
        
Error_Date:
        MsgBox "Error date: Pleas verify data again.", vbInformation, "Error Data"
        Assign_DateToCtrl = True

End Function

Public Sub Delete_Record(TableName As String, Condition As String)
If Condition = "" Then
        strCmdSQL = "DELETE " & TableName & ""
Else
        strCmdSQL = "DELETE " & TableName & " WHERE " & Condition
End If
        dbActive.Execute strCmdSQL
End Sub

Public Function Define_Runno(TBname As String) As String
Dim DefineCond As String
              DefineCond = " RNtbname = '" & TBname & "'"
                Set rsFunction = New Adodb.Recordset
                With rsFunction
                        strCmdSQL = ("select *  from Runno where " & DefineCond)
                        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                        If Not .EOF Then
                                Define_Runno = Trim(rsFunction!RNtbcode) & rsFunction!RNyear & rsFunction!RNmonth & Format(rsFunction!RNrunno, rsFunction!RNformat)
                                .Close
                                strCmdSQL = (" Update Runno set  RNrunno=RNrunno+1 where " & DefineCond)
                                .Open strCmdSQL
                        Else
                                Define_Runno = ""
                        End If
                End With
                Set rsFunction = Nothing
End Function

Public Sub Save_Rec()
'Dim sTBname As String
'Dim sDPcode As String
'Dim BaktxtFld1 As String
'Dim SelectTB As String
'Dim UpdRunnoCond As String
'        '������� Table ����� Runno ���� ��� ����� AccountDB ���ӡ�á�˹�
'        If Find_Ret_Val("RUNNO", "RNtbName", "RNtbname='" & UCase(tbActive) & "'") <> "" Then
'               '�ӡ�ä���Table ����ͧ��� Running ������ Runno ������������
'                If UCase(Trim(frmActive.txtfld1.Text)) = "***NEW***" Then       '�� Auto running no
'                        If AccountDB Then '����� AccountDB ����ö��˹���� txtFld1 ��
'                                  frmActive.txtfld1.Text = "NEW DOCNO"
'                                  GoTo Next_StepSave
'                         End If
'
'                        BaktxtFld1 = Trim(frmActive.txtfld1.Text)
'                        SelectTB = Trim(tbActive)
'
'                        Select Case UCase(tbActive)
'                        Case "WKORDER", "CONTRACT" '�ӡ�� Update Ἱ���Ἱ��ѹ
'                                    UpdRunnoCond = " RNtbname = '" & SelectTB & "' And DpCode='" & UsrDept & "'"
'
'                        Case Else
'                                    UpdRunnoCond = " RNtbname = '" & SelectTB & "'"
'
'                        End Select
'
'                           Set rsFunction = New Adodb.Recordset
'                           With rsFunction
'                                    strCmdSQL = ("select *  from Runno where " & UpdRunnoCond)
'                                    .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'                                    If Not .EOF Then
'                                             Select Case UCase(tbActive)
'                                             Case "PRICEREQUEST"
'                                                      frmActive.txtfld1.Text = Trim(rsFunction!RNtbcode) & rsFunction!RNyear & "/" & _
'                                                                                                   rsFunction!RNmonth & "-" & Format(rsFunction!RNrunno, rsFunction!RNformat)
'
'                                             Case Else
'                                                      frmActive.txtfld1.Text = Trim(rsFunction!RNtbcode) & rsFunction!RNyear & rsFunction!RNmonth & Format(rsFunction!RNrunno, rsFunction!RNformat)
'
'                                             End Select
'
'                                             .Close
'                                             strCmdSQL = (" Update Runno set  RNrunno=RNrunno+1 where " & UpdRunnoCond)
'                                             .Open strCmdSQL
'
'                                    Else
'                                             MsgBox "This Table Not Use Running No.", vbCritical, "Error"
'                                             Call Cancel_Rec
'                                             Exit Sub
'
'                                    End If
'
'                           End With
'                           Set rsFunction = Nothing
'                  End If
'
'        Else '�����辺
''                If frmActive.txtfld1.Text = "" Then    '��Ǩ�ͺ��� Field ��ѡ��ҧ�������
''                        MsgBox frmActive.txtfld1.Tag & " not empty. !", vbCritical, "Empty primary key"
''                        Exit Sub
''                End If
'        End If
'Next_StepSave:
        rsActive.fields!LastUpdate = Now
        rsActive.fields!LastUser = UsrSTFCode
      
         On Error GoTo Err_Save
         rsActive.Update
         bAddRecord = True
Exit Sub

Err_Save:
'        '�ӡ��Ŵ Running NO. �ʴ���� Erron �͹ New
'        If Trim(BaktxtFld1) = "***NEW***" Then  '��� UpdRunnoCond
'                dbActive.Execute "Update Runno set  RNrunno=RNrunno-1 where " & UpdRunnoCond
'        End If
        If Err.Number = -2147217900 Then
                MsgBox "�����ŷ��س��ͧ��úѹ�֡���������º��������!", vbCritical, "Already record "
        Else
                MsgBox Err.Description
        End If
        Call Cancel_Rec
End Sub


Public Function FieldTypeNumeric(rsFldType As Adodb.Recordset, colno As Variant) As Boolean
If Trim(colno) <> "" Then
    FieldTypeNumeric = (rsFldType.fields(colno).Type = adNumeric) _
                                             Or (rsFldType.fields(colno).Type = adSmallInt) _
                                             Or (rsFldType.fields(colno).Type = adUnsignedSmallInt) _
                                             Or (rsFldType.fields(colno).Type = adTinyInt) _
                                             Or (rsFldType.fields(colno).Type = adUnsignedTinyInt) _
                                             Or (rsFldType.fields(colno).Type = adBigInt) _
                                             Or (rsFldType.fields(colno).Type = adUnsignedBigInt) _
                                             Or (rsFldType.fields(colno).Type = adInteger) _
                                             Or (rsFldType.fields(colno).Type = adUnsignedInt) _
                                             Or (rsFldType.fields(colno).Type = adSingle) _
                                             Or (rsFldType.fields(colno).Type = adCurrency)
Else
  FieldTypeNumeric = False
End If
End Function


Public Sub Refresh_Record(rsSend As Adodb.Recordset)  '��ͧ������ӡ�� Requery ��������
        Dim i As Integer
        On Error Resume Next
        With frmActive
        
        With frmActive
                Set .dgMain.DataSource = rsSend
                i = 0
                CntCol = 0
                For Each fields In rsSend.fields
                        '�������ʴ��ҧ Fields ����ͧ���
                        If InStr(1, HideField, UCase(Trim(rsSend.fields(i).Name))) <> 0 Then
                                .dgMain.Columns.Remove CntCol
                        ElseIf Right(UCase(Trim(rsSend.fields(i).Name)), 4) = "FLAG" Then
                                .dgMain.Columns.Remove CntCol
                        Else
                                .dgMain.Columns(CntCol).Caption = Find_FldDesc(rsSend.fields(i).Name)
                                '��� Field Type �� ����Ţ���Դ���
                                If FieldTypeNumeric(rsSend, i) Then .dgMain.Columns(CntCol).Alignment = dbgRight
                                
                                '�������ҧ Colum
                                If Find_FldColWidth(rsSend.fields(i).Name) <> 0 Then .dgMain.Columns(CntCol).Width = Find_FldColWidth(rsSend.fields(i).Name)
                                CntCol = CntCol + 1 '�纨ӹǹ Colum ��������������
                        End If
                        i = i + 1
                Next
        End With
        End With
 End Sub


Public Sub Cancel_Rec()
         With rsActive
                  .CancelBatch
                  .CancelUpdate
         End With
End Sub

Public Sub RowColChg()
  On Error Resume Next
  With rsBrowse
'       frmmain.stbfrmmain.Panels("record").Text = "Record No :" & .AbsolutePosition & "/" & .RecordCount & "  "
       If Err.Number <> 0 Then Call Last_Rec(rsBrowse)
   End With
End Sub


Public Function Upper_Char(KeyAscii As Integer) As Integer
Select Case KeyAscii
 Case 65 To 91, 97 To 123   'A-Z,a-z
       If (KeyAscii >= 97 And KeyAscii <= 123) Then '����繵����硷�����繵���˭�
           Upper_Char = KeyAscii - 32
        End If
Case Else
       Upper_Char = KeyAscii
End Select
End Function

Public Function Chk_Alpha(KeyAscii As Integer) As Integer
Select Case KeyAscii
 Case 65 To 91, 97 To 123, 8, 13, 27 'A-Z,a-z
           Chk_Alpha = KeyAscii
        
   Case Else
         MsgBox "Input text Only.", vbInformation, "Error"
         KeyAscii = 10
        Chk_Alpha = KeyAscii
End Select
End Function

Public Function Chk_Digit(KeyAscii As Integer) As Integer
        Select Case KeyAscii
                Case 48 To 57, 8, 13, 27, 46 '0-9 BackSpace Enter Escape Period
                        Chk_Digit = KeyAscii
                        
                Case Else
                        MsgBox "Input Numeric Only.", vbInformation, "Error"
                        KeyAscii = 10
                        Chk_Digit = KeyAscii
        End Select
End Function

Public Function Chk_Dupkey(dbfName As String, Condition As String) As Boolean
        Set rsFunction = New Adodb.Recordset
        strCmdSQL = "select  *  from " & dbfName & " where " & Condition
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText

 If rsFunction.EOF Then
        
        Chk_Dupkey = True
  Else
     MsgBox "This record is already. ", vbInformation + vbOKOnly, "Already record"
     Chk_Dupkey = False
  End If
  Set rsFunction = Nothing
End Function

Public Function Find_Rec(dbfName As String, Condition As String) As Boolean
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  *  from " & dbfName & " where " & Condition
  rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
 If Not rsFunction.EOF Then
     Find_Rec = True
  Else
     Find_Rec = False
  End If
  Set rsFunction = Nothing
End Function


Public Function Substr(strVal As String, PLeft As Integer, PRight As Integer) As String
Substr = Right(Left(strVal, PLeft), PRight) '������Ѻ��Ҥ�������ҧ String
End Function


Public Function max_Item(dbfName As String, FldName As String, Condition) As String
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  max( " & FldName & " ) as maxItem  from " & dbfName & " where " & Condition
  rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
 If Not IsNumeric(rsFunction!maxItem) Then
     max_Item = "1"
  Else
     max_Item = CStr(rsFunction!maxItem + 1)
  End If
  Set rsFunction = Nothing
End Function


Public Function Sum_Item(dbfName As String, FldName As String, Condition As String) As String
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  sum( " & FldName & " ) as sumItem  from " & dbfName & " where " & Condition
  rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
 If Not IsNumeric(rsFunction!sumItem) Then
     Sum_Item = 0
  Else
     Sum_Item = CStr(rsFunction!sumItem)
  End If
  rsFunction.Close
  Set rsFunction = Nothing
End Function


Public Function Count_Rec_Inc1(dbfName As String, Condition) As String
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  count( * ) as Cnt_Rec  from " & dbfName & " where " & Condition
  rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
  Count_Rec_Inc1 = Trim(CStr(rsFunction!Cnt_rec + 1))
  Set rsFunction = Nothing
End Function

Public Function FindListindex(RefValue As String, _
                                                TBname As String, _
                                                FindField As String, _
                                                Optional Condition As String, _
                                                Optional OrderField As String, _
                                                Optional OrderType As String, _
                                                Optional Distinct As Boolean) As Integer
        '����Ҥ�� Listindex �ͧ ComboBox
        '�� Procedure ����� Search ��� ListIndex, �¤��Ҩҡ FindField ����դ�� = RefCtrl ����� Properties "Text"
        Dim strCriteria As String
        Dim strTmp As String
        
        strTmp = ""
        If Distinct = True Then strTmp = " DISTINCT "
        If Trim(OrderType) = "" Then OrderType = "ASC"
                
        strCmdSQL = "SELECT " & strTmp & FindField & " FROM " & TBname
        If Trim(Condition) <> "" Then strCmdSQL = strCmdSQL & " WHERE " & Condition
        If Trim(OrderField) <> "" Then strCmdSQL = strCmdSQL & " ORDER BY " & OrderField & " " & OrderType
                
        Set rsFunction = New Adodb.Recordset
        rsFunction.Open strCmdSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
        With rsFunction
                strCriteria = FindField & " = '" & CStr(RefValue) & "'"
                .Find strCriteria, 0, adSearchForward, adBookmarkFirst
                If .BOF Or .EOF Then
                        FindListindex = -1
                        Exit Function
                End If
                FindListindex = .AbsolutePosition - 1
        End With
                
End Function

Public Function Find_FldDesc(fldCode As String) As String
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  flddesc  from fldhdrtb where fldcode = '" & fldCode & "'"
  rsFunction.Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
  If Not rsFunction.EOF Then
     Find_FldDesc = Trim$(CStr(rsFunction!flddesc))
  Else
     Find_FldDesc = Trim$(CStr(fldCode))
  End If
  Set rsFunction = Nothing
End Function

'��˹��������ҧ�ͧ Colum � DataGrid
Public Function Find_FldColWidth(fldCode As String) As Integer
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  fldcolwidth  from fldhdrtb where fldcode = '" & fldCode & "'"
  rsFunction.Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
  If Not rsFunction.EOF Then
     Find_FldColWidth = rsFunction!fldColWidth * 100
  Else
     Find_FldColWidth = 0
  End If
  Set rsFunction = Nothing
End Function

Public Function Find_PositionInTB(TableName As String, Condition As String, FieldOrder As String, Criteria As String) As Double
          Dim dblPosition As Double
          Set rsFunction = New Adodb.Recordset
          With rsFunction
                    If Condition = "" Then
                              strCmdSQL = "SELECT " & FieldOrder & " FROM " & TableName & " ORDER BY " & FieldOrder
                    Else
                              strCmdSQL = "SELECT " & FieldOrder & " FROM " & TableName & " WHERE " & Condition & " ORDER BY " & FieldOrder
                    End If
                    .Open strCmdSQL, dbActive, adOpenDynamic, adLockReadOnly, adCmdText
                    .Find Criteria, 0, adSearchForward, 1
                    If Not .BOF Or Not .EOF Then
                              Find_PositionInTB = .AbsolutePosition
                    Else
                              Find_PositionInTB = 0
                    End If
          End With
          Set rsFunction = Nothing
End Function

Public Function Find_FldCode(flddesc As String) As String
  '19/02/2010 ���ͧ�ҡ�Ҩ�ա�õ�� ���ͫ�ӡѹ���ҧ Table �ѹ �� Item
  Dim strTBName As String
  Dim cntrec As Integer
 strTBName = tbActive
 cntrec = Count_Record("FldHdrTB", "flddesc = '" & flddesc & "'")
If cntrec >= 2 Then '��ҫ�ӡѹ��ͧ�к� Table
    Find_FldCode = Find_Ret_Val("FldHdrTB", "fldcode", "flddesc = '" & flddesc & "' and TBName='" & strTBName & "'")
ElseIf cntrec = 1 Then
      Find_FldCode = Find_Ret_Val("FldHdrTB", "fldcode", "flddesc = '" & flddesc & "'")
Else
   Find_FldCode = Trim$(CStr(flddesc))
End If


'  Set rsFunction = New Adodb.Recordset
'  strCmdSQL = "select  fldcode  from fldhdrtb where flddesc = '" & flddesc & "'"
'  rsFunction.Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
'  If Not rsFunction.EOF Then
'     Find_FldCode = Trim$(CStr(rsFunction!fldCode))
'  Else
'     Find_FldCode = Trim$(CStr(flddesc))
'  End If
'  Set rsFunction = Nothing
End Function

Public Function Find_Ret_Val(dbfName As String, FldReturn As String, Condition As String) As String
        Dim i As Integer
        Set rsFunction = New Adodb.Recordset
        strCmdSQL = "select  " & FldReturn & " from " & dbfName & " where " & Condition
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        i = InStr(1, FldReturn, ".")
        If i > 0 Then
                FldReturn = Right(FldReturn, Len(FldReturn) - i)
        End If
        If Not rsFunction.EOF And Not IsNull(rsFunction.fields(0)) Then
                Find_Ret_Val = Trim(rsFunction.fields(0))
        Else
                Find_Ret_Val = ""
        End If
        Set rsFunction = Nothing
End Function


Public Function Find_Ret_NewRec(dbfName As String, FldReturn As String, Position As Integer, Condition As String) As String
        Dim i As Integer
        '���觤���Ţ��� �ͧ Record ����
        Set rsFunction = New Adodb.Recordset
        If Condition = "" Then
                strCmdSQL = "SELECT " & FldReturn & " FROM " & dbfName & " ORDER BY " & FldReturn
        Else
                strCmdSQL = "SELECT " & FldReturn & " FROM " & dbfName & " WHERE " & Condition & " ORDER BY " & FldReturn
        End If
        
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        If Not rsFunction.EOF Then
                If Not IsNull(rsFunction(0).Value) And Trim(rsFunction(0).Value) <> "" Then
                        rsFunction.MoveLast
                        If Position = 0 Then
                                Find_Ret_NewRec = Trim(rsFunction(0).Value) + 1
                                Exit Function
                        Else
                                Find_Ret_NewRec = Right(Trim(rsFunction(0).Value), Position) + 1
                        End If
                        i = Len(Find_Ret_NewRec)
                        Do Until i = Position
                                If i > Position Then
                                        Find_Ret_NewRec = ""
                                        GoTo Refresh_Recordset
                                End If
                                Find_Ret_NewRec = "0" & CStr(Find_Ret_NewRec)
                                i = i + 1
                        Loop
                End If
        Else
                If Position = 0 Then
                        Find_Ret_NewRec = "1"
                Else
                        Find_Ret_NewRec = "1"
                        i = Len(Find_Ret_NewRec)
                        Do Until i = Position
                                Find_Ret_NewRec = "0" & CStr(Find_Ret_NewRec)
                                i = i + 1
                        Loop
                End If
        End If
        
Refresh_Recordset:
        rsFunction.Close
        Set rsFunction = Nothing
End Function

Public Function Find_Ret_Num(dbfName As String, FldReturn As String, Condition As String) As String
  Set rsFunction = New Adodb.Recordset
  strCmdSQL = "select  " & FldReturn & " from " & dbfName & " where " & Condition
  rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
  If Not rsFunction.EOF Then
     Find_Ret_Num = CStr(rsFunction(Trim(FldReturn)))
  Else
    Find_Ret_Num = "0"
  End If
  Set rsFunction = Nothing
End Function

Public Function Chk_Sort(SortStr As String) As String
If SortStr = "������ҡ" Then
    Chk_Sort = "ASC"
 Else
    Chk_Sort = "DESC"
  End If
End Function

'������Ѻ���ҵ��˹觢ͧ Record ����ͧ������͹������ Index �ͧ Combobox ����� Style DropDown List ����������ö��˹��������ѹ��
Public Function FindListInCtrl(dbfName As String, FldName As String, fldValue As String, SelectCondition As String) As Integer
If fldValue = "" Then
   FindListInCtrl = -1
   Exit Function
End If
If SelectCondition = "" Then SelectCondition = "1=1"
  Set rsFunction = New Adodb.Recordset
  With rsFunction
  strCmdSQL = "select distinct " & FldName & " from " & dbfName & " where " & SelectCondition & " order by  " & FldName
  .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
FindCondition = FldName & " like '" & fldValue & "%'"
 If Not rsFunction.EOF Then
     .Find FindCondition
     If Not rsFunction.EOF Then
        FindListInCtrl = rsFunction.AbsolutePosition - 1 '���ͧ�ҡ Index ������ҡ 0
     Else
       FindListInCtrl = -1
     End If
  Else
     FindListInCtrl = -1
  End If
  End With
  Set rsFunction = Nothing
End Function

'������Ѻ���ҵ��˹觢ͧ Record ����ͧ������͹������ Index �ͧ
'Combobox ����� Style DropDown List ����������ö��˹��������ѹ��
'��䢡óշ�� Field ���2 ����դ�ҵ͹ AddToCtrl ����� Index �͹��Ҥ׹ ���١��ͧ
Public Function FindListInCtrl1(dbfName As String, fldName1 As String, fldName2 As String, fldValue As String, SelectCondition As String) As Integer
If fldValue = "" Then
   FindListInCtrl1 = -1
   Exit Function
End If
If SelectCondition = "" Then SelectCondition = "1=1"
  Set rsFunction = New Adodb.Recordset
  With rsFunction
        If Trim$(fldName2) = "" Then
            strCmdSQL = "select distinct " & fldName1 & " from " & dbfName & " where " & SelectCondition & " order by  " & fldName1
        Else
           strCmdSQL = "select distinct " & fldName1 & "," & fldName2 & " from " & dbfName & " where " & SelectCondition & " order by  " & fldName1
        End If
  
  .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
FindCondition = fldName1 & " like '" & fldValue & "%'"
 If Not rsFunction.EOF Then
     .Find FindCondition
     If Not rsFunction.EOF Then
        FindListInCtrl1 = rsFunction.AbsolutePosition - 1 '���ͧ�ҡ Index ������ҡ 0
     Else
       FindListInCtrl1 = -1
     End If
  Else
     FindListInCtrl1 = -1
  End If
  End With
  Set rsFunction = Nothing
End Function


'������Ѻ��Ǩ�ͺ�ѹ�������� Null �������������ŧ Table � Form CROTODR
Public Function Assign_DateToFld(fldDate As Variant) As Variant
    Assign_DateToFld = IIf(IsDate(fldDate), Format(fldDate, Fdate), Null)
End Function


Public Sub AddListToCtrl(TBname As String, fldName1 As String, fldName2 As String, CtrlList As Control, Condition As String)
        '������Ѻ�����¡�����Ѻ Control �� Combobox
        If Condition = "" Then Condition = "1=1"
        Set rsFunction = New Adodb.Recordset
        
        If Trim$(fldName2) = "" Then
                strCmdSQL = "SELECT distinct " & fldName1 & " FROM " & TBname
        Else
                strCmdSQL = "SELECT  distinct " & fldName1 & "," & fldName2 & " FROM " & TBname
        End If
        
        'Check TB month
        
        If TBname = "Month" Then
                strCmdSQL = strCmdSQL & " WHERE " & Condition
        Else
                strCmdSQL = strCmdSQL & " WHERE " & Condition & " ORDER BY  " & fldName1
        End If
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        CtrlList.Clear
        
        With rsFunction
                If Not .BOF Then
                        .MoveFirst
                        Do Until .EOF
                                If Trim$(fldName2) = "" Then
                                        CtrlList.AddItem .fields(fldName1)
                                        
                                Else
                                       '���ͷ�����õѴ��ͧ��ҧ�١��ͧ��������� field ��� 2 �Ҵ���
                                       '��Ң�Ҵ�ͧ Field-������Ǣͧ������ ���������ҡѺ 0
                                       
                                       If (.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1))) <> 0 Then
                                          CtrlList.AddItem Trim(.fields(fldName1)) & Space((.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1)))) & " - " & .fields(fldName2)
                                       Else
                                          CtrlList.AddItem Trim(.fields(fldName1)) & " - " & .fields(fldName2)
                                       End If
                                End If
                                .MoveNext
                        Loop
                End If
                .Close
        End With
                Set rsFunction = Nothing
End Sub

Public Sub AddDataToList(TBname As String, _
                                                  fldName1 As String, _
                                                  fldName2 As String, _
                                                  Condition As String, _
                                                  fldOrder As String, _
                                                  SortType As String, _
                                                  CtrlList As Control, _
                                                  Distinct As Boolean)
        '������Ѻ�����¡�����Ѻ Control �� Combobox
        If Condition = "" Then Condition = "1=1"
        Set rsFunction = New Adodb.Recordset
        
        strCmdSQL = "SELECT "
        If Distinct = True Then strCmdSQL = strCmdSQL & " DISTINCT "
        
        If Trim(fldName2) = "" Then
                strCmdSQL = strCmdSQL & fldName1 & " FROM " & TBname
        Else
                strCmdSQL = strCmdSQL & fldName1 & "," & fldName2 & " FROM " & TBname
        End If
        
        If Trim(Condition) <> "" Then
                strCmdSQL = strCmdSQL & " WHERE " & Condition
        End If
        
        If Trim(fldOrder) <> "" Then
                strCmdSQL = strCmdSQL & " ORDER BY " & fldOrder & " " & SortType
        End If
        
        'Check TB month
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        CtrlList.Clear
        
        With rsFunction
                
                If Not .BOF Then
                        .MoveFirst
                        Do Until .EOF
                                
                                If Trim$(fldName2) = "" Then
                                        If .fields(fldName1).Type = adDBTimeStamp Then
                                                  CtrlList.AddItem Format(.fields(fldName1), "MMM DD, YYYY")
                                        Else
                                                  CtrlList.AddItem .fields(fldName1)
                                        End If
                                Else
                                       '���ͷ�����õѴ��ͧ��ҧ�١��ͧ��������� field ��� 2 �Ҵ���
                                       '��Ң�Ҵ�ͧ Field-������Ǣͧ������ ���������ҡѺ 0
                                       If (.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1))) <> 0 Then
                                          CtrlList.AddItem Trim(.fields(fldName1)) & Space((.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1)))) & " - " & .fields(fldName2)
                                       Else
                                          CtrlList.AddItem Trim(.fields(fldName1)) & " - " & .fields(fldName2)
                                       End If
                                End If
                                .MoveNext
                        Loop
                End If
                .Close
        End With
        Set rsFunction = Nothing
End Sub

'Update Procedure
Public Sub AddItemToCtrl(TBname As String, _
                                                  fldName1 As String, _
                                                  fldName2 As String, _
                                                  Condition As String, _
                                                  fldOrder As String, _
                                                  SortType As String, _
                                                  CtrlList As Control, _
                                                  Distinct As Boolean, _
                                                  Clear As Boolean)
        '������Ѻ�����¡�����Ѻ Control �� Combobox
        If Condition = "" Then Condition = "1=1"
        Set rsFunction = New Adodb.Recordset
        
        strCmdSQL = "SELECT "
        If Distinct = True Then strCmdSQL = strCmdSQL & " DISTINCT "
        
        If Trim(fldName2) = "" Then
                strCmdSQL = strCmdSQL & fldName1 & " FROM " & TBname
        Else
                strCmdSQL = strCmdSQL & fldName1 & "," & fldName2 & " FROM " & TBname
        End If
        
        If Trim(Condition) <> "" Then
                strCmdSQL = strCmdSQL & " WHERE " & Condition
        End If
        
        If Trim(fldOrder) <> "" Then
                strCmdSQL = strCmdSQL & " ORDER BY " & fldOrder & " " & SortType
        End If
        
        'Check TB month
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        If Clear = True Then
                CtrlList.Clear
        End If
        
        With rsFunction
                
                Dim b As Byte
                If Not .BOF Then
                        .MoveFirst
                        Do Until .EOF
                                '�Ѵ . �͡�ҡ Column Name
                                b = InStr(1, fldName1, ".")
                                If b > 0 Then
                                        fldName1 = Right(fldName1, Len(fldName1) - b)
                                End If
                                b = InStr(1, fldName2, ".")
                                If b > 0 Then
                                        fldName2 = Right(fldName2, Len(fldName2) - b)
                                End If
                                
                                If Trim$(fldName2) = "" Then
                                        If .fields(fldName1).Type = adDBTimeStamp Then
                                                  CtrlList.AddItem Format(.fields(fldName1), "MMM DD, YYYY")
                                        Else
                                                  CtrlList.AddItem .fields(fldName1)
                                        End If
                                Else
                                       '���ͷ�����õѴ��ͧ��ҧ�١��ͧ��������� field ��� 2 �Ҵ���
                                       '��Ң�Ҵ�ͧ Field-������Ǣͧ������ ���������ҡѺ 0
                                       If (.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1))) <> 0 Then
                                          CtrlList.AddItem Trim(.fields(fldName1)) & Space((.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1)))) & " - " & .fields(fldName2)
                                       Else
                                          CtrlList.AddItem Trim(.fields(fldName1)) & " - " & .fields(fldName2)
                                       End If
                                End If
                                .MoveNext
                        Loop
                End If
                .Close
        End With
        Set rsFunction = Nothing
End Sub

Public Function Check_Approve() As Boolean

    '��Ǩ�ͺ������Է�� Approve ��� ��Ἱ� ��� level �������ö����� Approve ���Ǩ�ͺ�ҡ MENUTV
    Set rsFunction = New Adodb.Recordset
     strCmdSQL = "select  * From menutv  where mtvkey='" & strCurrentNodeClick & "'"
     rsFunction.Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
     Check_Approve = InStr(1, rsFunction!MtvApproveLevel, UsrLevel) <> 0 And InStr(1, rsFunction!MtvApproveDept, UsrDept) <> 0
     rsFunction.Close
     Set rsFunction = Nothing
    
End Function


'=====================================================
'[�Ѻ�ӹǹ Record � Table "TbName" ������͹� strCondition]
Public Function Count_Record(TBname As String, strCondition As String)
        Set rsFunction = New Adodb.Recordset
        If strCondition = "" Then
                strCmdSQL = "SELECT COUNT(*) AS Total_Record FROM " & TBname
        Else
                strCmdSQL = "SELECT COUNT(*) AS Total_Record FROM " & TBname & " WHERE " & strCondition
        End If
        With rsFunction
                .Open strCmdSQL, dbActive, adOpenDynamic, adLockReadOnly, adCmdText
                Count_Record = .fields("Total_Record").Value
                .Close
        End With
        Set rsFunction = Nothing
End Function

Public Function DCount_Record(TBname As String, strCondition As String, Optional sColumn As String, Optional bDistinct As Boolean)
        
        Dim sSQL As String
        
        If Trim(sColumn) = "" Then sColumn = "*"
        If bDistinct = True Then sColumn = "Distinct " & sColumn
        sSQL = "SELECT COUNT(" & sColumn & ") AS Total_Record FROM " & TBname
        If strCondition <> "" Then sSQL = sSQL & " WHERE " & strCondition
        
        Set rsFunction = New Adodb.Recordset
        With rsFunction
                .Open sSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                DCount_Record = .fields("Total_Record").Value
                .Close
        End With
        Set rsFunction = Nothing
        
End Function

Public Function Have_Record(TBname As String, strCondition As String) As Boolean
          Set rsFunction = New Adodb.Recordset
          If strCondition = "" Then
                    strCmdSQL = "SELECT * FROM " & TBname
          Else
                    strCmdSQL = "SELECT * FROM " & TBname & " WHERE " & strCondition
          End If
          
          With rsFunction
                    .Open strCmdSQL, dbActive, adOpenDynamic, adLockReadOnly, adCmdText
                    If Not .EOF Then
                              Have_Record = True
                    Else
                              Have_Record = False
                    End If
                    .Close
          End With
          Set rsFunction = Nothing
End Function

Public Sub tbrSave_Switch()
        With frmActive!tbrSave.Buttons
                .Item(1).Visible = Not .Item(1).Visible
                .Item(2).Visible = Not .Item(2).Visible
                .Item(3).Visible = Not .Item(3).Visible
                If .Item(3).Visible = True Then
                        frmActive!cbrCenter.Bands(2).MinWidth = frmActive!tbrSave.ButtonWidth
                Else
                        frmActive!cbrCenter.Bands(2).MinWidth = frmActive!tbrSave.ButtonWidth * 2
                End If
        End With
End Sub

Public Sub Frm_UnProtect() '¡��ԡ��û�ͧ�ѹ��������� frmOPO

        For Each Ctrl In frmActive
                If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Then
                        Ctrl.Enabled = True
                ElseIf (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Then
                        Ctrl.Enabled = True
                End If
        Next

End Sub

Public Sub Define_Field_Tag(frmSend As Form, rsSend As Adodb.Recordset)
         For Each Ctrl In frmSend.Controls
                  If Trim$(Ctrl.Tag) <> "" Then
                           Set Ctrl.DataSource = rsSend
                           Ctrl.DataField = Trim(Ctrl.Tag)
                           If (rsSend.fields(Ctrl.Tag).Type = adChar Or rsSend.fields(Ctrl.Tag).Type = adVarChar) _
                           And (TypeOf Ctrl Is TextBox) Then
                                    Ctrl.MaxLength = rsSend.fields(Ctrl.Tag).DefinedSize
                                    Ctrl.Text = Trim$(Ctrl.Text) '���ͷӡ�õѴ��ͧ��ҧ�͡
                           End If
                  End If
         Next
End Sub

Public Sub Frm_Protect() '��ͧ�ѹ��������� frmOPO
        For Each Ctrl In frmActive
                If ((TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is UpDown)) And Trim(Ctrl.Tag) <> "" Then
                        Ctrl.Enabled = False
                ElseIf (TypeOf Ctrl Is DTPicker) Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is MaskEdBox) Then
                        If TypeOf Ctrl Is CommandButton Then Ctrl.Enabled = False
                End If
        Next
End Sub


Public Sub FindList_2Condition(TBname As String, fldWhere As String, ctrlWhere As Control, fldOrder As String, fldFind As String, ctrlFind As Control, CtrlList As Control)
'����Ҥ�� Listindex �ͧ ComboBox
'�� Procedure ����� Search ��� ListIndex, �¤��Ҩҡ fldWhere ����դ�� = ctrlText ����� Properties "Text"
    Dim strCriteria As String
    If Len(ctrlWhere.Text) <> "" Then
        Set rsFunction = New Adodb.Recordset
        strCmdSQL = "SELECT * FROM " & TBname & " WHERE " & fldWhere & " = '" & Trim$(ctrlWhere.Text) & "' ORDER BY " & fldOrder
        rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
        With rsFunction
            strCriteria = fldFind & " = '" & Trim$(ctrlFind.Text) & "'"
            .Find strCriteria, 0, adSearchForward, adBookmarkFirst
            If .BOF Or .EOF Then
                Exit Sub
            End If
            CtrlList.ListIndex = .AbsolutePosition - 1
        End With
    ElseIf Len(ctrlFind.Text) = 0 Then
        CtrlList.ListIndex = -1
    End If
End Sub

Public Sub FindList(TBname As String, FindField As String, OrderField As String, SelectCondition As String, ctrlFindList As Control, refCtrl As String)
        '����Ҥ�� Listindex �ͧ ComboBox
        '�� Procedure ����� Search ��� ListIndex, �¤��Ҩҡ FindField ����դ�� = RefCtrl ����� Properties "Text"
        Dim strCriteria As String
        If Len(refCtrl) > 0 Then
                Set rsFunction = New Adodb.Recordset
                If SelectCondition <> "" Then
                        strCmdSQL = "SELECT * FROM " & TBname & " WHERE " & SelectCondition & " ORDER BY " & OrderField
                Else
                        strCmdSQL = "SELECT * FROM " & TBname & " WHERE " & FindField & " <> ''" & " ORDER BY " & OrderField
                End If
                rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                With rsFunction
                        strCriteria = FindField & " = '" & CStr(Trim(refCtrl)) & "'"
                        .Find strCriteria, 0, adSearchForward, adBookmarkFirst
                        If .BOF Or .EOF Then
                                ctrlFindList.ListIndex = -1
                                Exit Sub
                        End If
                        ctrlFindList.ListIndex = .AbsolutePosition - 1
                End With
        End If
End Sub


Public Sub Assign_HeaderGrid(rs As Adodb.Recordset, dg As DataGrid)
Dim FldName As String
Dim i As Integer
    i = 0
    For Each fields In rs.fields
                '��˹� Header �� ������
                FldName = Find_FldDesc(rs.fields(i).Name)
                dg.Columns(i).Caption = FldName
                '��� Field Type �� ����Ţ���Դ���
                If FieldTypeNumeric(rs, i) Then dg.Columns(i).Alignment = dbgRight
                    
                '�������ҧ Colum  �����ҡѺ 0 ����ͧ������
                If Find_FldColWidth(rs.fields(i).Name) <> 0 Then dg.Columns(i).Width = Find_FldColWidth(rs.fields(i).Name)
                i = i + 1
    Next
End Sub

Public Sub Assign_HeaderGrid_cmb(rs As Adodb.Recordset, dg As DataGrid, cmb As ComboBox)
Dim FldName As String
Dim i As Integer
    i = 0
    For Each fields In rs.fields
                '��˹� Header �� ������
                FldName = Find_FldDesc(rs.fields(i).Name)
                dg.Columns(i).Caption = FldName
                cmb.AddItem FldName
                '�������ҧ Colum  �����ҡѺ 0 ����ͧ������
                If Find_FldColWidth(rs.fields(i).Name) <> 0 Then dg.Columns(i).Width = Find_FldColWidth(rs.fields(i).Name)
                i = i + 1
    Next
End Sub

Public Sub SP_DELETE_RECORD(TableName As String, Condition As String)

        Set cmdSp = New Adodb.Command
        With cmdSp
                .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "DELETE_RECORD"

                .Parameters.Append cmdSp.CreateParameter("Table", adVarChar, adParamInput, 50)
                .Parameters("Table").Value = TableName

                .Parameters.Append cmdSp.CreateParameter("Condition", adVarChar, adParamInput, 1000)
                .Parameters("Condition").Value = Condition

                .Execute
        End With
        Set cmdSp = Nothing
        
End Sub


Public Function ConvertDate(dtp As DTPicker)
        ConvertDate = DateSerial(Year(dtp), Month(dtp), Day(dtp))
End Function

Public Function LookupData(TBLookup, strLookup, Condition, Optional FldReturn As String) As String
        If Condition = "" Then LookupCondition = "1=1"
        LookupCondition = Condition
        LookupColumn = FldReturn
        ActiveLookup = TBLookup
        LookupRetVal = Trim(strLookup)
        frmLookup.Show vbModal '�����辺�������¡ Lookup �����觤�� LookupRetVal ��Ѻ�����
        LookupData = LookupRetVal
End Function



Public Sub DROP_TABLE(Table As String)
        strCmdSQL = "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE ID = OBJECT_ID(N'" & Table & "') " & _
                                "AND OBJECTPROPERTY(id, N'ISUSERTABLE') = 1)" & vbCrLf & _
                                "DROP TABLE " & Table
        dbActive.Execute strCmdSQL
End Sub

Public Sub DROP_VIEW(View As String)
        strCmdSQL = "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE ID = OBJECT_ID(N'" & View & "') " & _
                                "AND OBJECTPROPERTY(id, N'IsView') = 1)" & vbCrLf & _
                                "DROP VIEW " & View
        dbActive.Execute strCmdSQL
End Sub

Public Function ChkFileExist(FileSpec) As Boolean '��Ǩ�ͺ����� File ����ͧ��������������
  Dim fso, msg
  Set fso = CreateObject("Scripting.FileSystemObject")
  ChkFileExist = (fso.FileExists(FileSpec))
End Function



Public Function Hide_btnSave(frmHide_btnSave As Form)
                '���������ö Close �����ҧ������ѧ�ҡ�Դ Error ����
          With frmHide_btnSave
                .tbrSave.Buttons("save").Enabled = False
                .tbrSave.Buttons("save").Image = 0
                .tbrSave.Buttons("cancel").Image = 3
            End With
End Function

Public Function Show_btnSave(frmHide_btnSave As Form)
                '���������ö Close �����ҧ������ѧ�ҡ�Դ Error ����
          With frmHide_btnSave
                .tbrSave.Buttons("save").Enabled = True
                .tbrSave.Buttons("save").Image = 1
                .tbrSave.Buttons("cancel").Image = 2
            End With
End Function

Public Sub ConnectionDB(Database As String, frmActivate As Form)

        Set dbTemp = New Adodb.Connection
        With dbTemp
                .CursorLocation = adUseClient
                .Open "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=alluser;password=alluser;Initial Catalog=" & Database & ";Data Source=" & CurrentServer
        End With
        Set dbActive = dbTemp
        frmActivate.lblDatabase.Caption = Database
        
        OldDB = CurrentDB
        CurrentDB = Database
        blnDataTest = UCase(CurrentDB) = "DATATEST"
        
End Sub

Public Sub RefreshDB()
        Dim sCond As String
        
        On Error Resume Next
        
        With dbTemp
                .Close
        End With
        Set dbTemp = Nothing
        
        'Set dbActive ��Ѻ��� Connect ��� PneumaxDB
        Set dbActive = dbSQL
        
        sCond = "CVcode = 'REPORT'"
        RptPath = DGetDataLetCtxt("ConstVar", "CVValue", sCond)
        CurrentDB = OldDB
        OldDB = ""
        
End Sub

'Oct 24, 2002
'*** Change date format on program loading
Public Sub CH_DateFormat()

        Dim locale As Long
        Dim result As Long
        Dim strLocaleInfo As String * 255

        locale = GetUserDefaultLCID()
        result = GetLocaleInfo(locale, LOCALE_SSHORTDATE, strLocaleInfo, 255)
        If Left(strLocaleInfo, 10) <> "dd/MM/yyyy" Then
                result = SetLocaleInfo(locale, LOCALE_SSHORTDATE, "dd/MM/yyyy")
        End If

End Sub


Public Function Find_TmpTime(Optional Table As String) As String
        If Trim(Table) = "" Then
                Find_TmpTime = Trim(UsrSTFCode) & Format(Now, "HHMMSS")
        Else
                Find_TmpTime = Trim(MaxValue(Table, "tmptime", ""))
        End If
End Function

Public Function Find_TmpTime_Old(Optional Table As String) As String
        If Trim(Table) = "" Then
                Find_TmpTime_Old = Trim(UsrDept) & Format(Now, "HHMMSS")
        Else
                Find_TmpTime_Old = Trim(MaxValue(Table, "tmptime", ""))
        End If
End Function


'FEB 5, 2003
'*** Search current quater from date.
Public Function Current_Quater(Optional sDate As String) As Integer

        Dim iMonth As Integer
        
        If Trim(sDate) = "" Then
                sDate = CurrentDate
        End If
                
        iMonth = Month(sDate)
        Select Case iMonth
        Case 1 To 3
                Current_Quater = 1
        Case 4 To 6
                Current_Quater = 2
        Case 7 To 9
                Current_Quater = 3
        Case 10 To 12
                Current_Quater = 4
        End Select
        
End Function

' *** Feb 19, 2003 ***
' ���§�ӴѺ Item � Detail �ͧ TBname ����
Public Sub ReOrder_ItemNo(TBname As String, TBkey As String)
        
        Set cmdSp = New Command
        With cmdSp
        .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "Reorder_Itemno"
                .Parameters.Append cmdSp.CreateParameter("doctype", adChar, adParamInput, 3)
                .Parameters("doctype").Value = TBname
                
                .Parameters.Append cmdSp.CreateParameter("docno", adVarChar, adParamInput, 15)
                .Parameters("docno").Value = TBkey
                
                .Execute
        End With
        Set cmdSp = Nothing

End Sub

' *** 19/10/2011 ***
' ���§�ӴѺ Item � Detail �ͧ TBname ���� �ó� Table �� 3 key
Public Sub ReOrder_SubItemNo(DocType As String, DocKey1 As String, DocKey2 As String, DocItem As Integer)
        
        Set cmdSp = New Command
        With cmdSp
        .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "Reorder_SubItemno"
                .Parameters.Append cmdSp.CreateParameter("doctype", adChar, adParamInput, 3)
                .Parameters("doctype").Value = Trim(DocType)
                
                .Parameters.Append cmdSp.CreateParameter("DocKey1", adVarChar, adParamInput, 16)
                .Parameters("DocKey1").Value = Trim(DocKey1)
                
                .Parameters.Append cmdSp.CreateParameter("DocKey2", adVarChar, adParamInput, 40)
                .Parameters("DocKey2").Value = Trim(DocKey2)
                
                .Parameters.Append cmdSp.CreateParameter("Docitem", adSmallInt, adParamInput)
                .Parameters("Docitem").Value = DocItem
                .Execute
        End With
        Set cmdSp = Nothing

End Sub


Public Sub Swap_Itemno(sDocType As String, sDocNo As String, iFromItem As Integer, iToItem As Integer)
       Set cmdSp = New Command
        With cmdSp
                .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "Swap_Itemno" '�ӡ�� ��Ѻ itemno
                
                .Parameters.Append cmdSp.CreateParameter("doctype", adChar, adParamInput, 3)
                .Parameters("doctype").Value = sDocType
                
                .Parameters.Append cmdSp.CreateParameter("docno", adChar, adParamInput, 10)
                .Parameters("docno").Value = Trim(sDocNo)
                
                .Parameters.Append cmdSp.CreateParameter("from_itemno", adInteger, adParamInput)
                .Parameters("from_itemno").Value = iFromItem
                
                .Parameters.Append cmdSp.CreateParameter("to_itemno", adInteger, adParamInput)
                .Parameters("to_itemno").Value = iToItem
                
                .Execute
        End With
        Set cmdSp = Nothing
End Sub

'============================================================================
'��á�˹� Function ��ҧ� ���Ѻ Report Menu
'
Public Sub RptMnu_Default(CtrlRcv As Object, pMnuLevel As Integer)
On Error Resume Next
Dim i As Integer
        For i = CtrlRcv.LBound To CtrlRcv.UBound
               With CtrlRcv(i)
                        .MouseIcon = frmLogin.imgHandPoint.Picture '�������ӧҹ�����Ǣ��
                        .MousePointer = 99
                        .FontUnderline = False
                        Select Case pMnuLevel
                                    Case ERptMnuLevel.Menu
                                                .ForeColor = EMnuColor.Default
                                    Case ERptMnuLevel.SubMenu
                                                .ForeColor = ESubMnuColor.Default
                                    Case ERptMnuLevel.Opt
                                                .ForeColor = EOptColor.Default
                        End Select
               End With
        Next
End Sub

Public Sub RptMnu_Frame_MouseMove(CtrlRcv As Object, pMnuLevel As Integer)
        On Error Resume Next
        Dim i As Integer
        For i = CtrlRcv.LBound To CtrlRcv.UBound
                With CtrlRcv(i)
                        .FontUnderline = False
                        .FontSize = 8
                        .FontBold = False
                        Select Case pMnuLevel
                                        Case ERptMnuLevel.Menu
                                                    If .ForeColor <> EMnuColor.Choice Then .ForeColor = EMnuColor.Default
                                        Case ERptMnuLevel.SubMenu
                                                    If .ForeColor <> EMnuColor.Choice Then .ForeColor = ESubMnuColor.Default
                                        Case ERptMnuLevel.Opt
                                                    If .ForeColor <> EOptColor.Choice Then .ForeColor = EOptColor.Default
                        End Select
                End With
                DoEvents
        Next i
End Sub

Public Sub RptMnu_Label_MouseMove(CtrlRcv As Object, pMnuLevel As Integer, Index As Integer)
        With CtrlRcv(Index)
                    Select Case pMnuLevel
                            Case ERptMnuLevel.Menu
                                            If .ForeColor = EMnuColor.Choice Then Exit Sub
                                            .FontUnderline = True
                                            .ForeColor = EMnuColor.Active
                            Case ERptMnuLevel.SubMenu
                                            If .ForeColor = ESubMnuColor.Choice Then Exit Sub
                                            .FontUnderline = True
                                            .ForeColor = ESubMnuColor.Active
                            Case ERptMnuLevel.Opt
                                            If .ForeColor = EOptColor.Choice Then Exit Sub
                                            .FontUnderline = True
                                            .ForeColor = EOptColor.Active
                    End Select
                    '.FontSize = 10
                    '.FontBold = True
        End With
End Sub


Public Sub RptMnu_Label_Click(CtrlRcv As Object, pMnuLevel As Integer, Index As Integer)
On Error Resume Next
Dim i  As Integer
    '�������
    For i = CtrlRcv.LBound To CtrlRcv.UBound
                With CtrlRcv(i)
                        .FontUnderline = False
                        Select Case pMnuLevel
                                 Case ERptMnuLevel.Menu
                                             .ForeColor = EMnuColor.Default
                                 Case ERptMnuLevel.SubMenu
                                             .ForeColor = ESubMnuColor.Default
                                 Case ERptMnuLevel.Opt
                                             .ForeColor = EOptColor.Default
                         End Select
                End With
    Next i
    '��Ƿ�����͡
    With CtrlRcv(Index)
                .FontUnderline = False
               Select Case pMnuLevel
                        Case ERptMnuLevel.Menu
                                    .ForeColor = EMnuColor.Choice
                        Case ERptMnuLevel.SubMenu
                                    .ForeColor = ESubMnuColor.Choice
                        Case ERptMnuLevel.Opt
                                    .ForeColor = EOptColor.Choice
                End Select
    End With
End Sub
'������Ѻ Control �����  Array
Public Sub Assign_CurrentDate_Todtp(dtpDate As Object)
Dim i  As Integer
        For i = dtpDate.LBound To dtpDate.UBound
                dtpDate(i).Value = DateAdd("d", -Day(CurrentDate), CurrentDate)
        Next
End Sub

Public Sub Assign_Dept_ToCmb(pCmbDept As Object, pCondition As String)
        On Error Resume Next
        Dim i  As Integer
                
        For i = pCmbDept.LBound To pCmbDept.UBound
            Call AddListToCtrl("Department", "DPcode", "DPname", pCmbDept(i), pCondition)
        Next
        
End Sub

Public Sub Assign_ChkDept_Click(pChkDept As Object, pCmbDept As Object, pLblDept As Object, _
                                                                      pChkSales As Object, pCmdSales As Object, ptxtSales As Object, Index As Integer)
On Error Resume Next
        Select Case pChkDept(Index).Value
        Case vbChecked
                pCmbDept(Index).Enabled = False
                pCmbDept(Index).ListIndex = -1
                pLblDept(Index).Caption = ""
                
                pChkSales(Index).Enabled = True
                pCmdSales(Index).Enabled = True
                ptxtSales(Index).Enabled = True
        
        Case vbUnchecked
                With pChkSales(Index)
                        .Enabled = False
                        .Value = vbUnchecked
                End With
                With ptxtSales(Index)
                        .Enabled = False
                        .Text = ""
                End With
                With pCmdSales(Index)
                        .Enabled = False
                End With
                pCmbDept(Index).Enabled = True
        End Select
End Sub


Public Sub Assign_ChkSales_Click(pLblDept As Object, pChkSales As Object, pCmdSales As Object, _
                                                                        ptxtSales As Object, Optional pdtpFromDate As Object, Optional pdtpToDate As Object, Optional Index As Integer)
On Error Resume Next
        Select Case pChkSales(Index).Value
        Case vbChecked
                pdtpFromDate(Index).Enabled = True
                pdtpToDate(Index).Enabled = True

                pCmdSales(Index).Enabled = False
                With ptxtSales(Index)
                        .Enabled = False
                        .Text = ""
                End With
                If Trim(pLblDept(Index)) = "" Then
                        strCondition = ""
                Else
                        strCondition = "DPcode = '" & Trim(pLblDept(Index)) & "'"
                End If
        Case vbUnchecked
                'pdtpFromDate(Index).Enabled = False
                'pdtpToDate(Index).Enabled = False

                pCmdSales(Index).Enabled = True
                With ptxtSales(Index)
                        .Enabled = True
                End With
        End Select
End Sub



Public Sub Assign_CtrlRpt_FormLoad(pChkDept As Object, pCmbDept As Object, _
                                                                      pChkSales As Object, pCmdSales As Object, ptxtSales As Object)

Dim i As Integer
                For i = pChkDept.LBound To pChkDept.UBound
                        pChkDept(i).Enabled = False
                Next
                
                For i = pCmbDept.LBound To pCmbDept.UBound
                        With pCmbDept(i)
                                .Enabled = False
                                .ListIndex = 0
                        End With
                Next
                
                If UsrLevel = "1" Then
                        For i = pChkSales.LBound To pChkSales.UBound
                                    pChkSales(i).Enabled = False
                                    pCmdSales(i).Enabled = False
                                    ptxtSales(i).Enabled = False
                                    ptxtSales(i).Text = UsrSCode
                        Next
                 End If
End Sub

Public Sub Define_Frame_Caption(pLblSubMnu As Object, pFraParam As Object)
Dim i As Integer
        '��˹� Caption ���Ѻ Frame Parameter
        For i = pLblSubMnu.LBound To pLblSubMnu.UBound
                pFraParam(i).Caption = pLblSubMnu(i).Caption
                pFraParam(i).ForeColor = FrameRptForeColor
        Next
End Sub


'==========================================================================
Public Sub AddStrMonthToList(CtrlList As Control)
  CtrlList.Clear
  CtrlList.AddItem "JAN"
  CtrlList.AddItem "FEB"
  CtrlList.AddItem "MAR"
  CtrlList.AddItem "APR"
  CtrlList.AddItem "MAY"
  CtrlList.AddItem "JUN"
  CtrlList.AddItem "JUL"
  CtrlList.AddItem "AUG"
  CtrlList.AddItem "SEP"
  CtrlList.AddItem "OCT"
  CtrlList.AddItem "NOV"
  CtrlList.AddItem "DEC"
End Sub


Public Sub ADefine_MouseIcon_HandPoint(CtrlRcv As Object)
Dim i As Integer
        For i = CtrlRcv.LBound To CtrlRcv.UBound
                With CtrlRcv(i)
                            .MouseIcon = frmLogin.imgHandPoint.Picture
                           .MousePointer = 99
                End With
    Next
End Sub


Public Sub CloseForm(frm As Form)
       Unload frm
End Sub

Public Sub DGetDataLetClist(CtrlList As Control, _
                                                TBname As String, _
                                                fldName1 As String, _
                                                Optional fldName2 As String, _
                                                Optional Condition As String, _
                                                Optional OrderField As String, _
                                                Optional OrderType As String, _
                                                Optional Distinct As Boolean, _
                                                Optional Clear As Boolean)

        Dim strSQL As String
        
        '������Ѻ�����¡�����Ѻ Control �� Combobox
        If Condition = "" Then Condition = "1=1"
        Set rsFunction = New Adodb.Recordset
        
        strSQL = "SELECT "
        If Distinct = True Then strSQL = strSQL & " DISTINCT "
        If Trim(fldName2) = "" Then
                strSQL = strSQL & fldName1 & " FROM " & TBname
        Else
                strSQL = strSQL & fldName1 & "," & fldName2 & " FROM " & TBname
        End If
        If Trim(Condition) <> "" Then strSQL = strSQL & " WHERE " & Condition
        If Trim(OrderField) <> "" Then strSQL = strSQL & " ORDER BY " & OrderField & " " & OrderType
        If Clear = True Then CtrlList.Clear
        
        'Check TB month
        rsFunction.Open strSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        With rsFunction
                
                Dim b As Byte
                If Not .BOF Then
                        .MoveFirst
                        Do Until .EOF
                                '�Ѵ . �͡�ҡ Column Name
                                b = InStr(1, fldName1, ".")
                                If b > 0 Then
                                        fldName1 = Right(fldName1, Len(fldName1) - b)
                                End If
                                b = InStr(1, fldName2, ".")
                                If b > 0 Then
                                        fldName2 = Right(fldName2, Len(fldName2) - b)
                                End If
                                
                                If Trim$(fldName2) = "" Then
                                        If .fields(fldName1).Type = adDBTimeStamp Then
                                                  CtrlList.AddItem Format(.fields(fldName1), "MMM DD, YYYY")
                                        Else
                                                  CtrlList.AddItem .fields(fldName1)
                                        End If
                                Else
                                       '���ͷ�����õѴ��ͧ��ҧ�١��ͧ��������� field ��� 2 �Ҵ���
                                       '��Ң�Ҵ�ͧ Field-������Ǣͧ������ ���������ҡѺ 0
                                       If (.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1))) <> 0 Then
                                          CtrlList.AddItem Trim(.fields(fldName1)) & Space((.fields(fldName1).DefinedSize) - Len(Trim(.fields(fldName1)))) & " - " & .fields(fldName2)
                                       Else
                                          CtrlList.AddItem Trim(.fields(fldName1)) & " - " & .fields(fldName2)
                                       End If
                                End If
                                .MoveNext
                        Loop
                End If
                .Close
        End With
        Set rsFunction = Nothing
End Sub

Public Function DGetDataLetCtxt(TBname As String, ColName As String, Condition As String, Optional ValFormat As String)

        Dim strSQL As String
        Dim i As Integer
        
        Set rsFunction = New Adodb.Recordset
        strSQL = "select  " & ColName & " from " & TBname & " where " & Condition
        rsFunction.Open strSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
        i = InStr(1, ColName, ".")
        If i > 0 Then
                ColName = Right(ColName, Len(ColName) - i)
        End If
        
        If Not rsFunction.EOF Then
            If Not IsNull(rsFunction.fields(ColName).Value) Then
                        DGetDataLetCtxt = Trim(rsFunction.fields(ColName))
                        If ValFormat <> "" Then DGetDataLetCtxt = Format(DGetDataLetCtxt, ValFormat)
                Else
                        DGetDataLetCtxt = ""
             End If
        Else
                  Select Case rsFunction.fields(ColName).Type
                  Case adBigInt, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adCurrency, 17
                           DGetDataLetCtxt = 0
                  Case Else
                           DGetDataLetCtxt = ""
                  End Select
        End If
        Set rsFunction = Nothing

End Function




'========================    SERVICE SYSTEM ===============================

Public Sub DGetDataTolvwMain(lvwMain As ListView, rs As Adodb.Recordset, blnSetColWidth As Boolean)
Dim j As Integer
Dim Col_Width As Double
Dim Col_Alignment As Byte
            With lvwMain
                    .View = lvwReport
                    .ColumnHeaders.Clear
                    .ListItems.Clear
                    .Enabled = True
                    .FullRowSelect = True
            End With
            If rs.RecordCount <> 0 Then
                          i = 1
                         j = 0
                         For Each fields In rs.fields
                                    Select Case fields.Type
                                    Case adChar, adVarChar, adLongVarChar
                                            Col_Alignment = lvwColumnLeft
                                    Case adDBTimeStamp, adDBDate
                                            Col_Alignment = lvwColumnLeft
                                    Case adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adCurrency, adBigInt
                                            Col_Alignment = lvwColumnRight
                                    Case adCurrency, adNumeric, adUnsignedTinyInt
                                            Col_Alignment = lvwColumnRight
                                    Case Else
                                            Col_Alignment = lvwColumnLeft
                                    End Select
                                    Col_Width = Find_FldColWidth(rs.fields(j).Name)
                                    '�����ҡѺ 0 �����ҡѺ������Ǣͧ����
                                     If Col_Width = 0 Then Col_Width = Len(Trim(rs.fields(j).Name)) * 100
                                     If blnSetColWidth Then
                                        lvwMain.ColumnHeaders.Add i, , Find_FldDesc(rs.fields(j).Name), Col_Width, Col_Alignment
                                    Else
                                        lvwMain.ColumnHeaders.Add i, , Find_FldDesc(rs.fields(j).Name), , Col_Alignment
                                    End If
                                     
                                     j = j + 1 '���ӴѺ Field
                                      i = i + 1 '�纨ӹǹ Colum ��������������
                         Next
                         
                         i = 1 'Row ��� i
                         With rs
                             .MoveFirst
                             Do While Not .EOF
                                    For j = 1 To .fields.Count
                                              If j = 1 Then
                                                    lvwMain.ListItems.Add i, , rs.fields(j - 1).Value
                                              Else
                                                    If IsNull(rs.fields(j - 1).Value) Then
                                                       lvwMain.ListItems(i).ListSubItems.Add j - 1, , ""
                                                    Else
                                                       lvwMain.ListItems(i).ListSubItems.Add j - 1, , rs.fields(j - 1).Value
                                                    End If
                                            End If
                                    Next 'Loop For
                                    i = i + 1
                                    .MoveNext
                                Loop
                    End With
            Else
                        lvwMain.ColumnHeaders.Add 1, , "Not Found Record."
            End If
End Sub


'#################  All Sub And Function For Use In Service System #######################

Public Function MaxValue(dbfName As String, FldName As String, Condition) As String
        Set rsFunction = New Adodb.Recordset
        With rsFunction
                Select Case Trim(Condition)
                Case ""
                        strCmdSQL = "SELECT MAX(" & FldName & ") AS Maximum FROM " & dbfName
                Case Else
                        strCmdSQL = "SELECT MAX(" & FldName & ") AS Maximum FROM " & dbfName & " WHERE " & Condition
                End Select
                .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                If Not .EOF And Not IsNull(.fields("Maximum").Value) Then
                        MaxValue = .fields("Maximum").Value
                Else
                        MaxValue = 0
                End If
        End With
        Set rsFunction = Nothing
End Function

Public Function MinValue(dbfName As String, FldName As String, Condition) As String
        Set rsFunction = New Adodb.Recordset
        With rsFunction
                Select Case Trim(Condition)
                Case ""
                        strCmdSQL = "SELECT MIN(" & FldName & ") AS Minimum FROM " & dbfName
                Case Else
                        strCmdSQL = "SELECT MIN(" & FldName & ") AS Minimum FROM " & dbfName & " WHERE " & Condition
                End Select
                .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                If Not .EOF And Not IsNull(.fields("Minimum").Value) Then
                        MinValue = .fields("Minimum").Value
                Else
                        MinValue = 0
                End If
        End With
        Set rsFunction = Nothing
End Function



'[�Ѻ�ӹǹ Record � Table "TbName" ������͹� strCondition]
Public Function DGetRecordset(ByRef rs As Adodb.Recordset, TBname As String, Optional Condition As String, Optional OrderClmn As String, Optional ClmnName As String) As Boolean
        Dim strSQL As String
        Set rs = New Adodb.Recordset
        strSQL = "SELECT * FROM " & TBname
        If Trim(ClmnName) = "" Then ClmnName = " * "
        If Trim(Condition) <> "" Then strSQL = "SELECT " & ClmnName & " FROM " & TBname & " WHERE " & Condition
        If Trim(OrderClmn) <> "" Then strSQL = strSQL & " ORDER BY " & OrderClmn
        With rs
                .Open strSQL, dbActive, adOpenDynamic, adLockReadOnly, adCmdText
                DGetRecordset = rs.RecordCount
        End With
End Function

Public Function DGetKeyOfTable(rsSend As Recordset) As String
                DGetKeyOfTable = ""
                Select Case UCase(tbActive)
                            Case "INVOICE"
                                            DGetKeyOfTable = "INVNO='" & Trim(rsSend!invno) & "'"
                            Case "INVOICEDETAIL"
                                            DGetKeyOfTable = "INVNO='" & Trim(rsSend!invno) & "' and Invdt_item=" & Trim(rsSend!invdt_item)
                            Case "CUSTOMER"
                                            DGetKeyOfTable = "CSCODE='" & Trim(rsSend!CsCode) & "'"
                            Case "EMSRATE", "REGSERVICE"
                                            DGetKeyOfTable = "minweight=" & Trim(rsSend!minweight)
                            Case "USERS"
                                            DGetKeyOfTable = "userid='" & Trim(rsSend!userid) & "'"
                End Select
'        Dim sSQL As String
'        Dim i As Integer
'                DGetKeyOfTable = ""
'                Set rsFunction = New Adodb.Recordset
'                With rsFunction
'                        sSQL = "SP_PKeys " & TBname
'                        .Open sSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
'                        i = .RecordCount
'                        If i > 0 Then
'                                .MoveFirst
'                                Do Until i = 0
'                                        FldRefresh = .fields("Column_Name").Value
'                                        If Trim(DGetKeyOfTable) = "" Then
'                                                DGetKeyOfTable = .fields("Column_Name").Value
'                                        Else
'                                                DGetKeyOfTable = DGetKeyOfTable & ", " & .fields("Column_Name").Value
'                                        End If
'                                        i = i - 1
'                                        .MoveNext
'                                        DoEvents
'                                Loop
'                        End If
'                End With
'                Set rsFunction = Nothing
End Function

Public Function DGetKeyOfTableToSort(TBname As String, SortType As String) As String
        Dim sSQL As String
        Dim i As Integer
                DGetKeyOfTableToSort = ""
                Set rsFunction = New Adodb.Recordset
                With rsFunction
                        sSQL = "SP_PKeys " & TBname
                        .Open sSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
                        i = .RecordCount
                        If i > 0 Then
                                .MoveFirst
                                Do Until i = 0
                                        FldRefresh = .fields("Column_Name").Value
                                        If Trim(DGetKeyOfTableToSort) = "" Then
                                                DGetKeyOfTableToSort = .fields("Column_Name").Value & " " & SortType
                                        Else
                                                DGetKeyOfTableToSort = DGetKeyOfTableToSort & ", " & .fields("Column_Name").Value & " " & SortType
                                        End If
                                        i = i - 1
                                        .MoveNext
                                        DoEvents
                                Loop
                        End If
                End With
                Set rsFunction = Nothing
End Function

Public Function DGetKeyToSelect(TBname As String, rsSend As Adodb.Recordset) As String

On Error Resume Next
                DGetKeyToSelect = ""
                Select Case UCase(tbActive)
                            Case "INVOICE"
                                            DGetKeyToSelect = "INVNO='" & Trim(rsSend!invno) & "'"
                            Case "INVOICEDETAIL"
                                            DGetKeyToSelect = "INVNO='" & Trim(rsSend!invno) & "' and Invdt_item=" & Trim(rsSend!invdt_item)
                            Case "CUSTOMER"
                                            DGetKeyToSelect = "CSCODE='" & Trim(rsSend!CsCode) & "'"
                            Case "EMSRATE", "REGRATE"
                                            DGetKeyToSelect = "minweight=" & Trim(rsSend!minweight)
                            Case "USERS"
                                            DGetKeyToSelect = "userid='" & Trim(rsSend!userid) & "'"
                End Select


'        Dim sSQL As String
'        Dim i As Integer
'        Dim j As Integer
'        Dim a As String
'        Dim blnNumeric As Boolean
'        Dim sKeyValue As String
'
'
'
'
'        DGetKeyToSelect = ""
'        a = ""
'        j = 1
'        Set rsFunction = New Adodb.Recordset
'        With rsFunction
'                sSQL = "SP_PKeys " & TBname
'                .Open sSQL, dbActive, adOpenForwardOnly, adLockReadOnly, adCmdText
'                i = .RecordCount
'                If i > 0 Then
'                        .MoveFirst
'                        Do Until i = 0
'                                a = .fields("Column_Name").Value
'                                KeyValue(j) = rs.fields(a).Value
'                                blnNumeric = FieldTypeNumeric(rs, a)
'                                If blnNumeric = True Then
'                                             If rs.fields(a).Value = Empty Then
'                                                      sKeyValue = 0
'                                             Else
'                                                      sKeyValue = rs.fields(a).Value
'                                             End If
'                                Else
'                                        sKeyValue = "'" & rs.fields(a).Value & "'"
'                                End If
'                                If Trim(DGetKeyToSelect) = "" Then
'                                        DGetKeyToSelect = a & " = " & sKeyValue
'                                Else
'                                        DGetKeyToSelect = DGetKeyToSelect & " AND " & a & " = " & sKeyValue
'                                End If
'
'                                i = i - 1
'                                j = j + 1
'                                .MoveNext
'                                DoEvents
'                        Loop
'                End If
'        End With
'        Set rsFunction = Nothing
End Function


Public Sub MoveForm_MouseDown(X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Public Sub MoveForm_MouseMove(sForm As Form, sButton As Integer, X As Single, Y As Single)
    DoEvents
    If sButton = vbLeftButton Then
        With sForm
            .Left = .Left - (MouseX - X)
            .Top = .Top - (MouseY - Y)
        End With
    End If

End Sub

'******************************* ListView ****************************

Public Function DGetDataLetCgrid(ByRef rs As Adodb.Recordset, lvwCtrl As Control) As Long

        Dim strCond As String
        Dim rsTmp As Adodb.Recordset
        Dim lvwText As String
        Dim i As Integer
        Dim strKey As String
        Dim itmH As ColumnHeader
        Dim itmX As ListItem

        Dim Col_Name As Adodb.Field
        Dim Col_Width As Double
        Dim Col_Alignment As Byte

        Screen.MousePointer = vbHourglass

        With lvwCtrl
                .View = lvwReport
                .ColumnHeaders.Clear
                .ListItems.Clear
                '.GridLines = True
                .Enabled = True
                .FullRowSelect = True
        End With

        DGetDataLetCgrid = rs.RecordCount

      Set rsTmp = Nothing
        Set rsTmp = New Adodb.Recordset
        Set rsTmp = rs

        ' *** Add header in listview ***
        With rsTmp

                ' *** �óըӹǹ Recordset = 0
                If .RecordCount = 0 Then
                        Set itmH = lvwCtrl.ColumnHeaders.Add(, , "Not Found", 10000, lvwColumnLeft)
                        lvwCtrl.GridLines = False
                        lvwCtrl.HideSelection = True
                        lvwCtrl.FullRowSelect = False
                        Screen.MousePointer = vbDefault
                        Exit Function
                End If

                i = 0
                For Each Col_Name In .fields
                        strKey = "K" & i
                        Select Case Col_Name.Type
                        Case adChar, adVarChar, adLongVarChar
                                Col_Alignment = lvwColumnLeft
                                Col_Width = Col_Name.ActualSize * 120
                                If Col_Name.ActualSize < 5 Then Col_Width = Col_Width * 2.5

                        Case adDBTimeStamp
                                Col_Alignment = lvwColumnLeft
                                Col_Width = 1100
                        Case adTinyInt, adSmallInt, adInteger, adSingle, adDouble, adCurrency, adBigInt
                                Col_Alignment = lvwColumnRight
                                Col_Width = 1000
                        Case adCurrency, adNumeric, adUnsignedTinyInt
                                Col_Alignment = lvwColumnRight
                                Col_Width = 1000
                        Case Else
                                Col_Alignment = lvwColumnLeft
                                Col_Width = 1500
                        End Select

                        Set itmH = lvwCtrl.ColumnHeaders.Add(, strKey, Col_Name.Name, Col_Width, Col_Alignment)

                        i = i + 1

                Next

                ' *** Add item in listview ***
                i = 0
                Do Until .EOF
                        For i = 0 To .fields.Count - 1
                                If IsNull(.fields(i).Value) Then
                                        lvwText = ""
                                Else
                                        Select Case .fields(i).Type
                                        Case adChar, adVarChar, adLongVarChar
                                                lvwText = Trim(.fields(i).Value)
                                        Case adDBTimeStamp
                                                lvwText = Format(Trim(.fields(i).Value), "dd/mm/yyyy")
                                        Case adTinyInt, adInteger, adSingle, adDouble, adCurrency, adBigInt
                                                lvwText = Format(Trim(.fields(i).Value), "#,##0.00")
                                        Case adSmallInt, adCurrency, adNumeric, adUnsignedTinyInt
                                                lvwText = Format(Trim(.fields(i).Value), "#,##0")
                                        Case Else
                                                lvwText = "Unknow Type"
                                        End Select
                                End If

                                If i = 0 Then
                                        Set itmX = lvwCtrl.ListItems.Add(, , lvwText)

                                Else
                                        itmX.SubItems(i) = lvwText

                                End If

                        Next
                        .MoveNext
                Loop

        End With

        With lvwCtrl
                .HideSelection = False
                .Refresh
        End With

        Screen.MousePointer = vbDefault

End Function


Public Function Convert_QtrToINT(iYear As Integer, iQuarter As Integer)
        Dim iMth As Integer
        Dim i As Integer
        
        Select Case iQuarter
        Case 1
                iMth = 3
        Case 2
                iMth = 6
        Case 3
                iMth = 9
        Case 4
                iMth = 12
        End Select
        
        i = Int(((iYear * 12) + (iMth - 1)) / 3)
        Convert_QtrToINT = i
        
End Function


Public Function GetStringToExec(a As String)
         GetStringToExec = Replace(a, "'", "''", 1)
End Function

Public Function AddZeroToStr(sText As String, iDigit As Integer, Optional sAddText As String)
        Dim i As Integer
        Dim j As Integer
        
        If Trim(sAddText) = "" Then sAddText = "0"
        i = Len(sText)
        
        For j = i To iDigit - 1
                If i >= iDigit Then GoTo End_Function
                sText = sAddText & sText
        Next j

End_Function:
        AddZeroToStr = sText
        
End Function

Public Sub Multi_Find(ByRef rs As Adodb.Recordset, sCriteria As String)

         Dim rsClone As Adodb.Recordset
         Set rsClone = rs.Clone
        '�繵�ǡ�˹� Bookmark
         rsClone.Filter = sCriteria

         If rsClone.EOF Or rsClone.BOF Then
                  sCriteria = ""
                  Exit Sub
                  'rs.MoveLast
                  'rs.MoveNext
         Else
                  rs.Bookmark = rsClone.Bookmark
         End If

         rsClone.Close
         Set rsClone = Nothing

End Sub

Public Function CStrToDate(sDate As String) As String
Dim sDay As String * 2
Dim sMonth As String * 3
Dim sYear As String * 4

If Not IsDate(sDate) Then
    CStrToDate = "Error."
    Exit Function
End If
sDay = CStr(Day(Format(CDate((sDate)), Fdate)))
sYear = CStr(Year(Format(CDate((sDate)), Fdate)))


Select Case Month(Format(CDate((sDate)), Fdate))
             Case 1
                         sMonth = "JAN"
             Case 2
                         sMonth = "FEB"
             Case 3
                         sMonth = "MAR"
             Case 4
                         sMonth = "APR"
             Case 5
                         sMonth = "MAY"
             Case 6
                         sMonth = "JUN"
             Case 7
                         sMonth = "JUL"
             Case 8
                         sMonth = "AUG"
             Case 9
                         sMonth = "SEP"
             Case 10
                         sMonth = "OCT"
             Case 11
                         sMonth = "NOV"
             Case 12
                         sMonth = "DEC"

End Select
CStrToDate = "" & sMonth & " " & sDay & "," & sYear & ""
End Function


Public Sub Insert_ItemNo(sDocType As String, sDocNo As String, iItem As Integer)
        Set cmdSp = New Command
        With cmdSp
        .ActiveConnection = dbActive
        .CommandType = adCmdStoredProc
        .CommandText = "Insert_Itemno"
        
        .Parameters.Append cmdSp.CreateParameter("doctype", adChar, adParamInput, 3)
        .Parameters("doctype").Value = Trim(sDocType)
        
        .Parameters.Append cmdSp.CreateParameter("docno", adChar, adParamInput, 15)
        .Parameters("docno").Value = Trim(sDocNo)
        
        .Parameters.Append cmdSp.CreateParameter("itemno", adInteger, adParamInput)
        .Parameters("itemno").Value = iItem
        
        .Execute
        End With
        Set cmdSp = Nothing
End Sub

Public Function blnAllowCancelDocNo(sDocType As String, sDocNo As String, sDocDate As String) As Boolean
'�ӡ�õ�Ǩ�ͺ����͡�������ö¡��ԡ���������
Select Case sDocType
              Case "ISS"
                          blnAllowCancelDocNo = Count_Record("ISSUEDT", " IssNo='" & sDocNo & "' And  StorePrepare='Y' And PartNid In (Select PartNid From Part where PartCntdate >='" & sDocDate & "')") = 0
End Select
End Function

Public Sub Define_MouseIcon_HandPoint(CtrlRcv As Object)
    With CtrlRcv
                .MouseIcon = frmLogin.imgHandPoint.Picture
               .MousePointer = 99
    End With
End Sub


'--================ Function ����������ѹ� Service System =====================

Public Sub InitializeLoadForm()
On Error Resume Next
        With frmActive
                '��˹� Mouse Pointer
                Call Define_MouseIcon_HandPoint(.lblFirst)
                Call Define_MouseIcon_HandPoint(.lblPrevious)
                Call Define_MouseIcon_HandPoint(.lblNext)
                Call Define_MouseIcon_HandPoint(.lblLast)
                
                
                'Call Define_MouseIcon_HandPoint(.tbrCommand)
                Call Define_MouseIcon_HandPoint(.tbrSave)
                Call Define_MouseIcon_HandPoint(.tbrExit)
                
                .Width = Screen.Width '��˹������ҡѺ�������ҧ�ͧ��
                .Height = Screen.Height  '��˹������ҡѺ�����٧�ͧ��
                
                 .Move 0, 300
                 .fraBrowse.Left = 0
                 .fraBrowse.Width = .Width
                 
                 .dgMain.Left = 0
                 .dgMain.Width = .fraBrowse.Width - 120
                 .lblTitle.Left = 0
                 .lblTitle.Width = .dgMain.Width
                 
                 .fraMain.Left = 0
                 .fraMain.Width = .Width
                 
                 .tbrCommand.Left = 0
                 .tbrCommand.Width = .Width
                 
                 .tbrExit.Left = .dgMain.Width - 800
                 
'                 .txtTbActive.Left = .tbrExit.Left - 3000
'                 .lblTbActive.Left = .txtTbActive.Left + 1500
                 
                  .txtTbActive.Left = Screen.Width / 2
                 .lblTbActive.Left = .txtTbActive.Left + 1500
                 .lblTbActive_Des.Left = .lblTbActive.Left
                 
                 .lblCntRecord.Left = .lblLast.Left + .lblLast.Width + 300
                 
                 
                 Err_Desc = "" '�ӡ�� Clear �������ҧ��͹

                
                Set rsBrowse = New Adodb.Recordset
                 strCmdSQL = Find_strCmdSQLForBrowse(tbActive)
                 rsBrowse.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                
               .lblCntRecord.Caption = "Record No :" & rsBrowse.AbsolutePosition & "/" & rsBrowse.RecordCount
                  
                  
                 '��˹� DataSource  ��� Tag
                 Set .dgMain.DataSource = rsBrowse
                 Call Define_Field_Tag(frmActive, rsBrowse)
                  '��˹�  Colum Header  ���Ѻ Data Grid ����� cmborder
                Call Assign_HeaderGrid(rsBrowse, .dgMain)
                Call Enable_tbrCommand
                Call Set_Button(rsBrowse)
                
        End With
End Sub

Public Sub Enable_tbrCommand()
        With frmActive
                    With .tbrCommand
                         .Buttons("ADD").Enabled = Check_Add(rsBrowse)
                         .Buttons("EDIT").Enabled = (rsBrowse.RecordCount <> 0) And Check_Edit(rsBrowse)
                         .Buttons("DELETE").Enabled = (rsBrowse.RecordCount <> 0) And Check_Delete(rsBrowse)
                         .Buttons("REFRESH").Enabled = (rsBrowse.RecordCount <> 0)
                    End With
                    .tbrSave.Visible = False
                    .tbrExit.Visible = Not .tbrSave.Visible
                    Call Frm_Protect
                    .fraBrowse.Enabled = True
        End With
End Sub

Public Sub Disable_tbrCommand()
        With frmActive
                    With .tbrCommand
                             .Buttons("ADD").Enabled = False
                             .Buttons("EDIT").Enabled = False
                             .Buttons("DELETE").Enabled = False
                             .Buttons("REFRESH").Enabled = False
                    End With
                    .fraBrowse.Enabled = False
                    .tbrSave.Visible = True
                    .tbrExit.Visible = Not .tbrSave.Visible
                    Call Frm_UnProtect
        End With
End Sub

Public Sub Set_Button(ByRef rsSend As Adodb.Recordset)
On Error Resume Next
        With frmActive
                With .tbrCommand
                         .Buttons("ADD").Enabled = Check_Add(rsSend)
                         .Buttons("EDIT").Enabled = (rsSend.RecordCount <> 0) And Check_Edit(rsSend)
                         .Buttons("DELETE").Enabled = (rsSend.RecordCount <> 0) And Check_Delete(rsSend)
                         .Buttons("REFRESH").Enabled = (rsSend.RecordCount <> 0)
                End With
                .fraBrowse.Enabled = True
        End With
End Sub
'
'Public Sub ActivateCommand(strMode As String, ByRef rsSend As Adodb.Recordset)
'        If (UCase(strMode) = "EDIT" Or UCase(strMode) = "DELETE") And (rsSend.RecordCount <> 0) Then
'                  With rsSend
'                           If .BOF Or .EOF Then
'                                    MsgBox "Please select record for edit !", vbInformation, "Error"
'                                    Exit Sub
'                           End If
'                  End With
'
'                '�ӡ���� Recode ����ͧ�ӡ�� Edit ���� Delete
'                sKey = DGetKeyOfTable(tbActive)
'                If sKey <> "" Then
'                        SelectCondition = DGetKeyToSelect(tbActive, rsSend)
'                End If
'                Set rsActive = New Adodb.Recordset
'                With rsActive
'                         strCmdSQL = "select  *  From " & tbActive & "  where " & SelectCondition
'                        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'                End With
'        End If
'
'        Select Case UCase(strMode)
'        Case "ADD"
'                        Set rsActive = New Adodb.Recordset
'                        With rsActive '��� ���͡������ҧ����㹡ó� ADD
'                                strCmdSQL = "SELECT  *  FROM " & tbActive & "  WHERE 1<>1 "
'                                .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'                                .AddNew
'                        End With
'                        Call Define_Field_Tag(frmActive, rsActive)
'                        Call Disable_tbrCommand
'                        bAddRecord = False '��� Add ����稨��դ���� True
'        Case "EDIT"
'                         Call Define_Field_Tag(frmActive, rsActive)
'                          If rsSend.RecordCount <> 0 Then Call Disable_tbrCommand
'
'        Case "DELETE"
'                Dim rsClone As Adodb.Recordset
'                Set rsClone = rsSend.Clone '���� Clone �����������Դ RowColChange
'                If MsgBox("Are you sure delete  this record and all detail.  ?", vbCritical + vbOKCancel, "Delete Record") = vbOK Then
'                        With rsSend
'                                If .RecordCount <> 0 Then
'                                        '�������˹����ǡѺ rsSend
'                                        rsClone.Bookmark = .Bookmark
'                                        rsClone.MoveNext '����� Record ����
'                                        If Not rsClone.EOF Then
'                                           SelectCondition = DGetKeyToSelect(tbActive, rsClone) '������͹�㹡�á�Ѻ�ҷ���ͧ���
'                                        Else
'                                           rsClone.Bookmark = .Bookmark '��������Ѻ�ҷ�����
'                                           rsClone.MovePrevious '����� Record ��͹˹��
'                                           If Not rsClone.BOF Then
'                                               SelectCondition = DGetKeyToSelect(tbActive, rsClone) '������͹�㹡�á�Ѻ�ҷ���ͧ���
'                                           End If
'                                        End If
'                                        rsClone.Close
'                                        Set rsClone = Nothing
'                                        With rsActive
'                                             .Delete  'ź�����Ţͧ rsActive ��� Main
'                                        End With 'End with rsactive
'                                        .Requery
'                                        sCriteria = SelectCondition
'                                        Call Multi_Find(rsSend, sCriteria)
'                                        'Call Refresh_Record(rsSend)
'                                End If
'                        End With
'                End If
'        Case "REFRESH"
'                 If rsSend.RecordCount <> 0 Then
'                        '�纤������͹��� Requery
'                        SelectCondition = DGetKeyToSelect(tbActive, rsSend)
'                        rsSend.Requery
'                        sCriteria = SelectCondition
'                        Call Multi_Find(rsSend, sCriteria)
'                        'Call Refresh_Record(rsSend)
'                  End If
'        End Select
'End Sub







Public Sub ActivateCommand_OldForm(strMode As String, ByRef rsSend As Adodb.Recordset, frmSend As Form)
        Dim sCond As String
        Dim dblSumAmt1 As Double
        Dim dblSumAmt2 As Double
        Dim savDocNo As String
        Dim savDocNo2  As String
        Dim savItem As Integer
        
        If (UCase(strMode) = "EDIT" Or UCase(strMode) = "DELETE") And (rsSend.RecordCount <> 0) Then
                  With rsSend
                           If .BOF Or .EOF Then
                                    MsgBox "Please select record for edit !", vbInformation, "Error"
                                    Exit Sub
                           End If
                  End With

'                Select Case UCase(tbActive)
'                            Case "INVOICE"
'                                            SelectCondition = "INVNO='" & Trim(rsSend!invno) & "'"
'                            Case "INVOICEDETAIL"
'                                            SelectCondition = "INVNO='" & Trim(rsSend!invno) & "' and Invdt_item=" & Trim(rsSend!invdt_item)
'                            Case "CUSTOMER"
'                                            SelectCondition = "CSCODE='" & Trim(rsSend!CsCode) & "'"
'                            Case "EMSRATE", "REGSERVICE"
'                                            SelectCondition = "minweight=" & Trim(rsSend!minweight)
'                            Case "USERS"
'                                            SelectCondition = "userid='" & Trim(rsSend!userid) & "'"
'
'                End Select
                
                '�ӡ���� Recode ����ͧ�ӡ�� Edit ���� Delete
'                sKey = DGetKeyOfTable(tbActive)
'                If sKey <> "" Then
                        SelectCondition = DGetKeyToSelect(tbActive, rsSend)
'                End If
                

                Set rsActive = New Adodb.Recordset
                With rsActive
                         strCmdSQL = "select  *  From " & tbActive & "  where " & SelectCondition
                        .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                End With
        End If
        
        Select Case UCase(strMode)
        Case "ADD"
                        Set rsActive = New Adodb.Recordset
                        With rsActive '��� ���͡������ҧ����㹡ó� ADD
                                strCmdSQL = "SELECT  *  FROM " & tbActive & "  WHERE 1<>1 "
                                .Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
                                Call Add_Rec
                        End With
                        Set frmActive = frmSend
                        If bAddRecord = True Then '����� Add ����� �����Ҩҡ
                            SelectCondition = DGetKeyToSelect(tbActive, rsActive)
                        Else
                            If rsSend.RecordCount <> 0 Then SelectCondition = DGetKeyToSelect(tbActive, rsSend)
                        End If
                        Set frmActive = frmSend
                        On Error Resume Next
                        rsSend.Requery
                        sCriteria = SelectCondition
                        Call Multi_Find(rsSend, sCriteria)
                        bAddRecord = False '��� Add ����稨��դ���� True
        Case "EDIT"
                          If rsSend.RecordCount = 0 Then
                                    MsgBox "Not Record For Edit", vbCritical, "Empty."
                                    Exit Sub
                          Else
                                     Call Edit_Rec(rsSend)
                                    Set frmActive = frmSend
                                    '�纤������͹��� Requery
'                                    SelectCondition = DGetKeyToSelect(tbActive, rsSend)
                                    On Error Resume Next
                                    rsSend.Requery
                                    sCriteria = SelectCondition
                                    Call Multi_Find(rsSend, sCriteria)
                         End If
        Case "DELETE"
                Dim rsClone As Adodb.Recordset
                Set rsClone = rsSend.Clone '���� Clone �����������Դ RowColChange
                If MsgBox("Are you sure delete  this record and all detail.  ?", vbCritical + vbOKCancel, "Delete Record") = vbOK Then
                        With rsSend
                                If .RecordCount <> 0 Then
                                        '�������˹����ǡѺ rsSend
                                        rsClone.Bookmark = .Bookmark
                                        rsClone.MoveNext '����� Record ����
                                        If Not rsClone.EOF Then
                                           SelectCondition = DGetKeyToSelect(tbActive, rsClone) '������͹�㹡�á�Ѻ�ҷ���ͧ���
                                        Else
                                           rsClone.Bookmark = .Bookmark '��������Ѻ�ҷ�����
                                           rsClone.MovePrevious '����� Record ��͹˹��
                                           If Not rsClone.BOF Then
                                               SelectCondition = DGetKeyToSelect(tbActive, rsClone) '������͹�㹡�á�Ѻ�ҷ���ͧ���
                                           End If
                                        End If
                                        rsClone.Close
                                        Set rsClone = Nothing
                                        '�ӡ�������͹���� Update ����ѧ
                                        Select Case tbActive
                                                    Case "WKISSUEOIL"
                                                                     dbSQL.Execute ("Update WkOrder set wiono='' where wiono='" & rsActive!wiono & "'")
                                                    Case "WKISSUEOILDT"
                                                                     dbSQL.Execute ("Update WkOrder set wiono='' where wkno='" & rsActive!WKno & "'")
                                                                    savDocNo = rsActive!wiono
                                                    Case "SCHEDULETIME"
                                                                    savDocNo = rsActive!CTNO
                                                    Case "CONTRACTDT"
                                                                    savDocNo = rsActive!CTNO
                                                    Case "CONTRACTSUBDETAIL"
                                                                    savDocNo = rsActive!CTNO
                                                                    savDocNo2 = ""
                                                                    savItem = rsActive!CTDT_Item
                                                    Case "MODELDT"
                                                                    savDocNo = rsActive!MadeCode
                                                                    savDocNo2 = rsActive!ModCode
                                                                    savItem = 0
                                        
                                        End Select
                                        With rsActive
                                             .Delete  'ź�����Ţͧ rsActive ��� Main
                                        End With 'End with rsactive
                                        Select Case tbActive
                                                    Case "WKISSUEOILDT" '�ӡ�� Update �ʹ�Թ� Header
                                                                '�ӡ�äӹǹ�ӹǹ�Թ���� Detail ��� Header
                                                                Set cmdSp = New Command
                                                                With cmdSp
                                                                         .ActiveConnection = dbActive
                                                                         .CommandType = adCmdStoredProc
                                                                         .CommandText = "Cal_WKIssueOilDt" '�ӹǹ�ӹǹ�Թ�����ҧ� Detail ��� header
                                                                         
                                                                         .Parameters.Append cmdSp.CreateParameter("docno", adChar, adParamInput, 10)
                                                                         .Parameters("docno").Value = Trim(savDocNo)
                                                                         
                                                                         .Parameters.Append cmdSp.CreateParameter("Err_Desc", adVarChar, adParamOutput, 200)
                                                                         .Execute
                                                                         If IsNull(cmdSp.Parameters("Err_Desc").Value) Then
                                                                            Err_Desc = "Syntax Error   ( " & .CommandText & ")"
                                                                         Else
                                                                            Err_Desc = cmdSp.Parameters("Err_Desc").Value
                                                                        End If
                                                                 End With
                                                                Set cmdSp = Nothing
                                                    Case "SCHEDULETIME"
                                                                Call ReOrder_ItemNo("SCH", savDocNo)
                                                    Case "CONTRACTDT"
                                                                Call ReOrder_ItemNo("CNT", savDocNo)
                                                    
                                                    '19/10/2011 �ӡ�� Reorder �ó����ҡ���� 1 ����
                                                    Case "CONTRACTSUBDETAIL"
                                                                   Call ReOrder_SubItemNo("CNS", savDocNo, savDocNo2, savItem)
                                                    Case "MODELDT"
                                                                '19/10/2011 �ӡ�� Reorder �ó����ҡ���� 1 ����
                                                                Call ReOrder_SubItemNo("MDT", savDocNo, savDocNo2, savItem)
                                       
                                        End Select
                                        
                                        Set frmActive = frmSend
                                        .Requery
                                        sCriteria = SelectCondition
                                        Call Multi_Find(rsSend, sCriteria)
                                End If
                        End With
                        Exit Sub
                End If
        Case "REFRESH"
                On Error Resume Next
                Set frmActive = frmSend
                 If rsSend.RecordCount <> 0 Then
                        '�纤������͹��� Requery
                        SelectCondition = DGetKeyToSelect(tbActive, rsSend)
                        rsSend.Requery
                        sCriteria = SelectCondition
                        Call Multi_Find(rsSend, sCriteria)
                        'Call Refresh_Record(rsSend)
                  End If
        Case "CRITERIA"
                   frmCriteria.Show vbModal
         End Select
End Sub
Public Sub First_Rec(ByRef rsSend As Adodb.Recordset)
  On Error GoTo GoFirstError
  With rsSend
  If Not (.BOF And .EOF) And .RecordCount > 0 Then
     .MoveFirst
    ' frmmain.stbfrmmain.Panels("record").Text = "Record No :" & .AbsolutePosition & "/" & .RecordCount
  End If
  End With
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Public Sub Last_Rec(ByRef rsSend As Adodb.Recordset)
  On Error GoTo GoLastError
  With rsSend
  If Not (.BOF And .EOF) And .RecordCount > 0 Then
     .MoveLast
    ' frmmain.stbfrmmain.Panels("record").Text = "Record No :" & .AbsolutePosition & "/" & .RecordCount
  End If
  End With
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Public Sub Next_Rec(ByRef rsSend As Adodb.Recordset)
  On Error GoTo GoNextError
  
  With rsSend
   If .BOF Then Exit Sub
  
  If Not .EOF Then
     .MoveNext
  End If
  
  If .EOF Then
    Beep
    .MoveLast
  End If
  'frmmain.stbfrmmain.Panels("record").Text = "Record No :" & .AbsolutePosition & "/" & .RecordCount
  End With
  Exit Sub

GoNextError:
  MsgBox Err.Description
End Sub

Public Sub Previous_Rec(ByRef rsSend As Adodb.Recordset)
  On Error GoTo GoPrevError
  With rsSend
  If .BOF Then Exit Sub
  If Not .BOF Then
       .MovePrevious
  End If
  
  If .BOF Then
    Beep
    .MoveFirst
  End If
  ' frmmain.stbfrmmain.Panels("record").Text = "Record No :" & .AbsolutePosition & "/" & .RecordCount
  End With
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Public Function Find_strCmdSQLForBrowse(tbActive As String, Optional strSpecialCond As String = "") As String

strCmdSQL = "select  *  from " & tbActive

        Select Case UCase(tbActive)
                      Case "INVOICEDETAIL"
                        Find_strCmdSQLForBrowse = strCmdSQL & " where Invno='" & Trim(rsBrowse!invno) & "'"
         Case Else
                        Find_strCmdSQLForBrowse = strCmdSQL
         End Select


''��������´��á�ͧ
''A - Filter by Head Menu (��ͧ��� Head Menu)
''B - Filter by ShareDept (��ͧ���������� ShareDept)
''C - Filter by DPcode (��ͧἹ� - �ءἹ�)
''D - Filter by Sales Dept (��ͧἹ� - ੾�н��¢����ҹ��)
''E - Filter by Sales Code (��ͧἹ���� Sales Code)
''F - Filter by Staff Code (��ͧἹ���� Staff Code)
''G - Filter by Root Menu (��ͧ��� Key �ͧ (Root Menu)
'
'
'Dim strTBBrowse As String
'Dim strOrderClmn As String
'Dim strOrderType As String
'Dim strTBName As String
'
'Dim sCond As String
'Dim blnShareDept As Boolean 'Share Department
'Dim blnHaveDPCode As Boolean '��ͧ Ἱ�
'Dim blnRoot As Boolean '�� Detail �����
'                 strMtvKey = tbActive
'                sCond = "MTVkey = '" & strMtvKey & "'"
'                DGetRecordset rsTemp, "MenuTV", sCond
'                With rsTemp
'                    strTBName = Trim(.fields("TBName"))
'                    strTBBrowse = Trim(.fields("TBName_Browse"))
'                    strOrderClmn = Trim(.fields("MTVOrderClmn"))
'                    strOrderType = Trim(.fields("MTVOrderType"))
'                   blnRoot = Trim(.fields("MTVRoot")) = "1"
'                   '����� Root ���������Ҩҡ����� �� ���͡�Ҩҡ Contract ��� Active
'                   '16-05-2019 ��������͡��੾�Чҹ�͡ ���� �ҹ� ��
'                   If blnRoot And Trim(HeaderCondition) = "" Then
'                      strCondition = Trim(.fields("MTVCondition"))
'                      If strSpecialCond <> "" Then
'                             '28-02-2020  ����� Table  Contract �����ҵ�����͹䢷�����͡
'                             If UCase(tbActive) = "CONTRACT" Then
'                                strCondition = strSpecialCond
'                            Else
'                                strCondition = strCondition & " and " & strSpecialCond
'                            End If
'                     End If
'                   End If
'                End With
'
'                sCond = "TBName = '" & Trim(UCase(strTBName)) & "'"
'                DGetRecordset rsTemp, "AllTable", sCond
'                With rsTemp
'                    blnShareDept = (Trim(.fields("TBHaveShareDept")) = "Y")
'                    blnHaveDPCode = (Trim(.fields("TBHaveDPCode")) = "Y")
'                    If Trim(strTBBrowse) = "" Then strTBBrowse = tbActive
'                    StrBrowseField = "*"
'                    '��� TBActive �����ҡѺ TBBrowse �ʴ������ View ����ʴ� Field ������
'                    If UCase(tbActive) = UCase(strTBBrowse) Then
'                        StrBrowseField = Trim(.fields("TBBrowseField")) 'Find_Ret_Val("AllTable", "TBBrowseField", " TBName = '" & Trim(UCase(tbActive)) & "'")
'                        If Trim$(StrBrowseField) = "" Then StrBrowseField = "*" '㹡óշ��������˹� Field ����ͧ�������ʴ�
'                    End If
'                End With
'
'                 If Trim(strCondition) = "" Then strCondition = "1=1"
'                 '�������� Root ����ͧʹ� Header Condition
'                 If blnRoot And HeaderCondition <> "" Then strCondition = strCondition & " and " & HeaderCondition
'
'                 '��Ǩ�ͺ Table ����� Field Name ���� Sharedept ����������͹�
'                 If blnShareDept And bSalesDept Then strCondition = strCondition & " and ShareDept Like '%" & UsrDept & "%'"
'                 '��Ǩ�ͺ Table ����� Field Name ���� DPCode ����������͹�
'                '㹡óշ���� Sales ��� Level ���¡���������ҡѺ C ���ӡ�á�ͧἹ�
'                 If blnHaveDPCode And bSalesDept And UsrLevel < "A" Then strCondition = strCondition & " and  DPCODE ='" & UsrDept & "'"
'
'                If Trim(strOrderClmn) <> "" Then '����� Order ��ͧ����� Order ���Ẻ�˹
'                    If strOrderType = "A" Then
'                        strCmdSQL = "select  " & StrBrowseField & " from " & strTBBrowse & " where " & strCondition & "  Order  By " & strOrderClmn & "  ASC"
'                    Else
'                        strCmdSQL = "select  " & StrBrowseField & " from " & strTBBrowse & " where " & strCondition & " Order  By " & strOrderClmn & "  DESC"
'                    End If
'                Else
'                    strCmdSQL = "select  " & StrBrowseField & " from " & strTBBrowse & " where " & strCondition
'                End If
'                Find_strCmdSQLForBrowse = strCmdSQL
End Function

Public Sub DataGrid_HeaderClick(ColIndex As Integer, ByRef rsSend As Adodb.Recordset)
        frmActive.txtFindValue = ""
        ColName = rsSend.fields(ColIndex).Name
'        If Trim(ColName) <> "" Then
'            If savColname = Trim(ColName) And Trim(savTBName) = Trim(tbActive) Then '�����ҡѹ�����Ѻ�ҡ�ҡ仹������͹�����ҡ
'                 If Trim(strSort) = "ASC" Then
'                     strSort = "DESC"
'                Else
'                     strSort = "ASC"
'                End If
'            Else
'                 If Find_Ret_Val("Menutv", "MTVOrderType", "TBName='" & Trim(tbActive) & "'") <> "D" Then
'                    strSort = "ASC"
'                 Else
'                    strSort = "DESC"
'                 End If
'            End If
'        End If
        rsSend.Sort = ColName & " " & strSort
        savTBName = Trim(tbActive)
        savColname = Trim(ColName)
End Sub



Public Function GetDefaultCondition() As String
Dim sCond  As String
            sCond = "MTVkey = '" & strMtvKey & "'"
            DGetRecordset rsTemp, "MenuTV", sCond
            GetDefaultCondition = Trim(rsTemp.fields("MTVCondition").Value)
End Function

Public Sub SearchDataInGrid(ByRef rsSend As Adodb.Recordset)
Dim sCriteria As String
Dim vBookMark
Dim strFind As String
        If Trim(ColName) = "" Then
           MsgBox "��س� Click ��� Colum Header  ��͹��ä��ҹФ�Ѻ.", vbCritical, "Click Colum Header."
           Exit Sub
        End If
        If rsSend.RecordCount = 0 Then Exit Sub
        strFind = frmActive.txtFindValue
        If FieldTypeNumeric(rsSend, ColName) Then
                sCriteria = ColName & " >= " & strFind
        Else
                sCriteria = ColName & " Like '" & strFind & "%'"
        End If
        With rsSend
                If Trim(strFind) = "" Then
                        .MoveFirst
                        Exit Sub
                End If

                vBookMark = .Bookmark
                .Find sCriteria
                If .BOF Or .EOF Then
                        .Bookmark = vBookMark
                End If
        End With
End Sub



'����ǡѺʶҹТͧ�١��� ��Һѭ�մ٨�����ա���� SACODE ��� DPCODE
Public Sub GetCust_Status(CsCode As String, SaCode As String, DPCode As String)
        If Trim(CsCode) <> "" Then
            Set rsFunction = New Adodb.Recordset
            strCmdSQL = "select  *  from Customer  where CSCode='" & Trim$(CsCode) & "'"
            rsFunction.Open strCmdSQL, dbActive, adOpenDynamic, adLockOptimistic, adCmdText
            If Not rsFunction.EOF Then
                    'txtCsThiname.Text = rsFunction!CSthiname
                    '�����˵��١���
                   If IsNull(rsFunction!CSRemark) Then
                       CSRemark = ""
                   Else
                      CSRemark = rsFunction!CSRemark
                   End If
                   '�١����ջѭ�Ҥ��   "Y"  Or  ("" Or "N" ����ջѭ��)
                   '    Y-BlackList ,G-WhiteList
                    If IsNull(rsFunction!CSblacklist) Or Trim(rsFunction!CSblacklist) = "" Then
                            BlackList = "N"
                    Else
                            BlackList = rsFunction!CSblacklist
                    End If
                    
                    If IsNull(rsFunction!CSblkLstrem) Then
                       CSblkLstrem = ""
                    Else
                       CSblkLstrem = rsFunction!CSblkLstrem
                    End If
                   
                   CSTerms = rsFunction!CSTerms
                   ARAMT = rsFunction!CS_ARAMT
            Else '�����辺
                    CSRemark = ""
                    BlackList = "N"
                    CSblkLstrem = ""
                    CSTerms = 0
                    ARAMT = 0
                     CSRemark = ""
            End If
            Set rsFunction = Nothing
            '��ǹŴ�ҵ�Ұҹ����Ἱ������ҡѹ
             If Trim(DPCode) <> "" Then
                StdAvgDisc = CDbl(Find_Ret_Num("CustDisc_Terms", "CsDiscRate", " Cscode ='" & Trim(CsCode) & "' And  DPCode='" & Trim(DPCode) & "'"))
             Else '�������ö�������ͧ�ҡ����Ἱ������ҡѹ
                StdAvgDisc = 0
             End If
             
      Else '��ҡѺ ""
            CSRemark = ""
            BlackList = "N"
            CSblkLstrem = ""
            CSTerms = 0
            ARAMT = 0
            CSRemark = ""
      End If
      If Trim(SaCode) <> "" Then
            '�ӹǹ�����ͧ�١��Ҥ�������ѧ��ҧ��������Ѻ�������Ф�
            CountBOR = Count_Record("vw_DOnoFinish", " Left(SAcode,4) ='" & Left(Trim(SaCode), 4) & "'")
     Else
             '�ӹǹ�����ͧ�١��Ҥ����
            CountBOR = Count_Record("vw_DOnoFinish", " CSCode ='" & Trim(CsCode) & "'")
     End If
End Sub


'����ǡѺ Credit Limit �ͧ�١���
Public Sub GetCust_CreditLimit(CsCode As String, DocNo As String)
            If Trim(CsCode) = "" Then Exit Sub
            
            Set cmdSp = New Command
            With cmdSp
                      .ActiveConnection = dbActive
                      .CommandType = adCmdStoredProc
                      .CommandText = "CHK_CREDITLIMIT" '��Ǩ�ͺ Credit Limit
                      .Parameters.Append cmdSp.CreateParameter("CSCODE", adChar, adParamInput, 10)
                      .Parameters("CSCODE").Value = Trim(CsCode)
                      
                      .Parameters.Append cmdSp.CreateParameter("DOCNO", adChar, adParamInput, 10)
                      .Parameters("DOCNO").Value = Trim(DocNo)
                      'Credit Limit ����˹�����١����������
                      .Parameters.Append cmdSp.CreateParameter("CRLMAMT", adCurrency, adParamOutput)
                      'Credit Limit ����ͧ�١��ҡ�������
                      .Parameters.Append cmdSp.CreateParameter("SUMCRLMAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ���������㹻Ѩ�غѹ�ͧ�١����������
                      .Parameters.Append cmdSp.CreateParameter("CURCRAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ���������㹻Ѩ�غѹ�ͧ�١��ҡ�������
                      .Parameters.Append cmdSp.CreateParameter("SUMCURCRAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ����ҧ���к� Dos �ͧ�١����������
                      .Parameters.Append cmdSp.CreateParameter("ARAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ����ҧ���к� Dos �ͧ�١��ҡ�������
                      .Parameters.Append cmdSp.CreateParameter("SUMARAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ��������ǧ˹�Ңͧ�١����������
                      .Parameters.Append cmdSp.CreateParameter("ADVCHQAMT", adCurrency, adParamOutput)
                      '�ʹ�Թ��������ǧ˹�Ңͧ�١��ҡ�������
                      .Parameters.Append cmdSp.CreateParameter("SUMADVCHQAMT", adCurrency, adParamOutput)
                      
                      .Execute
                     '����Ѻ�١����������
                      '�ӹǹ Credit ���������㹻Ѩ�غѹ�ͧ�١���������� ����� ��� QTT+CRO+INV+DO+�ʹ����ҧ�� DOS
                     CurCRAmt = cmdSp.Parameters("CURCRAMT").Value
                     CRLMAmt = cmdSp.Parameters("CRLMAMT").Value
                     ARAMT = cmdSp.Parameters("ARAMT").Value
                     AdvChqAmt = cmdSp.Parameters("ADVCHQAMT").Value
                     
                     '����Ѻ�����駡����
                      '�ӹǹ Credit ���������㹻Ѩ�غѹ�ͧ�١��� group ����� ��� QTT+CRO+INV+DO+�ʹ����ҧ�� DOS
                     SumCurCRAmt = cmdSp.Parameters("SUMCURCRAMT").Value
                     SUMCRLMAmt = cmdSp.Parameters("SUMCRLMAMT").Value
                     SUMARAmt = cmdSp.Parameters("SUMARAMT").Value
                     SumAdvChqAmt = cmdSp.Parameters("SUMADVCHQAMT").Value
               
               End With
               Set cmdSp = Nothing
End Sub


Public Sub Copy_Record(TBname As String)
        Dim strCol_Name As String
        Dim Value
        For Each Ctrl In frmActive.Controls
                If Trim$(Ctrl.Tag) <> "" Then
                        If Trim(strCol_Name) = "" Then
                                strCol_Name = Ctrl.Tag
                                If rsActive.fields(Ctrl.Tag).Type = adChar Or rsActive.fields(Ctrl.Tag).Type = adVarChar Or rsActive.fields(Ctrl.Tag).Type = adDate Then
                                        Value = "'" & Ctrl & "'"
                                Else
                                        Value = Ctrl
                                End If
                        Else
                                strCol_Name = strCol_Name & ", " & Ctrl.Tag
                                If rsActive.fields(Ctrl.Tag).Type = adChar Or rsActive.fields(Ctrl.Tag).Type = adVarChar Then
                                        Value = Value & ", " & "'" & Ctrl & "'"
                                ElseIf rsActive.fields(Ctrl.Tag).Type = adDBTimeStamp Then
                                        Value = Value & ", " & IIf(Trim(Ctrl) = "", "Null", "'" & Format(Ctrl, "MMM DD, YYYY") & "'")
                                Else
                                        Value = Value & ", " & CDbl(Ctrl)
                                End If
                        End If
                End If
        Next
        strCol_Name = "INSERT INTO " & TBname & " (" & strCol_Name & ") "
        Value = "VALUES (" & Value & ")"
        strCmdSQL = strCol_Name & Value
        
        dbActive.Execute strCmdSQL
        
End Sub




Public Sub Generate_ChkCreditLine(GR_Customer As String)
        'Input data to table strTempolay
        strTempolary = UsrDept & Format(Now, "HHMMSS")
        Set cmdSp = New Adodb.Command
        With cmdSp
                .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "CRT_Tmp_ChkCredit"

                .Parameters.Append cmdSp.CreateParameter("PrmCSgroup", adChar, adParamInput, 10)
                .Parameters("PrmCSgroup").Value = GR_Customer

                .Parameters.Append cmdSp.CreateParameter("TmpTime", adChar, adParamInput, 10)
                .Parameters("TmpTime").Value = strTempolary

                .Execute
        End With
        Set cmdSp = Nothing
        
        'Print report Customer
        Call Print_CS
        
        'Delete data from create on top line
        strTempolary = "TmpTime = '" & strTempolary & "'"
        Call Delete_Record("Tmp_ChkCredit", strTempolary)
        ''strTmp = "DELETE TMP_CHKCREDIT WHERE " & strTmp
        ''dbActive.Execute strTmp
End Sub

Public Sub Print_CS()
        Screen.MousePointer = vbHourglass
        With frmServiceSystem.CrptCS
                .Connect = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=alluser;PWD=alluser;Initial Catalog=" & CurrentDB & ";Data Source=" & CurrentServer
                .ReportFileName = RptPath & "rptChk_CreditLine_New.rpt"
                .ParameterFields(1) = "Time;" & strTempolary & ";True"
                .PrintReport
                Screen.MousePointer = vbDefault
        End With
End Sub

Public Function LookupAreaCode(Condition) As String

        If Condition = "" Then LookupCondition = "1=1"
        LookupCondition = Condition
        ActiveLookup = "AreaTB"
        LookupRetVal = ""
        frmLookup.Show vbModal
        LookupAreaCode = Trim(LookupRetVal)

End Function


Public Sub DisplayHelp(sTBname As String)
         If Count_Record("Alltable", "TBName='" & sTBname & "'") = 0 Then Exit Sub
          With frmRemark
                    .Caption = sTBname & "  Help..."
                    .lblTitle = "..." & sTBname & "  Help..."
                    .txtCause.Locked = UsrDept <> "MIS"
                    If UsrDept <> "MIS" Then
                        .tbrSave.Buttons("save").Enabled = False
                        .tbrSave.Buttons("save").Image = 0
                        .tbrSave.Buttons("cancel").Image = 3
                    End If
                    .txtCause.Text = Find_Ret_Val("ALLTABLE", "HELPDESC", "TBName='" & sTBname & "'")
                    .Show vbModal
                    If (Not .CancelEdit) Then
                        dbSQL.Execute "UPDATE ALLTABLE Set HELPDESC='" & Trim(RemarkValue) & "' ,LastUpdate=getdate(), LastUser='" & UsrSTFCode & "' Where TBName='" & sTBname & "'"
                    End If
          End With
End Sub



'========  ����������� ================

'Public Sub Delete_RecDt(dbfName As String, delCond As String)
'        Set rsFunction = New Adodb.Recordset
'            rsFunction.Open ("Delete from " & dbfName & "  where " & delCond), dbActive, adOpenDynamic, adLockOptimistic, adCmdText
'        Set rsFunction = Nothing
'End Sub


'Public Function GetFormActive(strFrmName As String) As Form
'        Select Case UCase(Trim(strFrmName))
'                        Case "SVTYPE"
'                                Set GetFormActive = frmAlltable
'                            Case Else
'                                Set GetFormActive = frmConstruction
'        End Select
'
'End Function



'
'Public Function GetSaleCondition(SaleLogin As String) As String
'    If UsrDept = "MIS" Then
'       GetSaleCondition = "1=1"
'       Exit Function
'    End If
'    Select Case strCurrentNodeClick
'                Case "SALESTB1"
'                            'Support ���� ���˹��Ἱ�����
'                            If (Find_Ret_Val("Position", "PSTSupport", "PSTCode='" & UsrPSTCode & "'") = "Y") Or UsrLevel >= "9" Then
'                                '���ͧ�ҡ SalesTb ��੾��� Sales ��ҹ�鹨֧��ͧ੾��Ἱ��������ͧ�ӡ�á�ͧ���˹�
'                                GetSaleCondition = "( SalesTb.DpCode='" & UsrDept & "')"
'                                Exit Function
'                             End If
'
'                '������͹��ҧ����������ͧ�ҡ�ա���к� ���� Table ����
'                Case "VWCUSTOMERINSALESAREA", "VW_LPO_CANCEL", "VW_LPO_RCVFULLNOAVL", "VW_LPO", "VW_CRO_PROBLEM1", _
'                             "VW_INVCUST_SALES", "VW_QTTCUST_SALES", "VW_CROCUST_SALES", "UPDTOP20CUSTOMER", "VW_CRO_APPRNOTSENDDR", "VW_CROTODR_ACCMATNOTCHECK"
'                            If UsrDept = "MAT" Or UsrLevel >= "A" Then
'                               GetSaleCondition = "1=1"
'                               Exit Function
'                            Else
'                                'Support ���� ���˹��Ἱ�����
'                                If (Find_Ret_Val("Position", "PSTSupport", "PSTCode='" & UsrPSTCode & "'") = "Y") Or UsrLevel >= "9" Then
'                                    '���ͧ�ҡ SalesTb ��੾��� Sales ��ҹ�鹨֧��ͧ੾��Ἱ��������ͧ�ӡ�á�ͧ���˹�
'                                    GetSaleCondition = "( DpCode='" & UsrDept & "')"
'                                    Exit Function
'                                 End If
'                             End If
'
'                Case "VW_QTT_AREAMGRAPPR", "VW_QTT_MGRAPPR", _
'                            "VW_CRO_AREAMGRAPPR", "VW_CRO_MGRAPPR", _
'                             "VW_LPO_MGRAPPR", "VW_CPA_MGRAPPR", "VW_CP_MGRAPPR", _
'                             "VW_CHG_PARTDES_MGRAPPR", "VW_FUNNEL1", _
'                             "VW_QTT_MGRAPPR_PO", "VW_QTT_DIRAPPR _PO"
'                            '���˹��Ἱ�����
'                             If UsrLevel >= "9" Then
'                                If Trim(DirGrpDept) <> "" Then
'                                     GetSaleCondition = "'" & DirGrpDept & " ' Like '%'+DPCODE+'%' "
'                                Else
'                                     GetSaleCondition = "DPCODE='" & UsrDept & "'"
'                                End If
'                                 Exit Function
'                             End If
'
'               Case "VW_QTT_DIRAPPR", "VW_CRO_DIRAPPR", _
'                            "VW_LPO_DIRAPPR", "VW_CHG_PARTDES_DIRAPPR"
'                            'Director ������繷ءἹ�
'                            If UsrLevel >= "A" Then
'                                 GetSaleCondition = "1=1"
'                                 Exit Function
'                             End If
'    End Select
'
'      top_stack = 0
'      SaleCondition = ""
'     Call SelectSale(SaleLogin)
'     Do While top_stack > 0
'        GetBoss = pop_stack()
'        Call SelectSale(GetBoss)
'     Loop
'     '���͹䢵�ҧ�ҡ������
'    If strCurrentNodeClick = "SALESTB1" Then
'             '�觤׹������͹䢷����
'             If SaleCondition <> "" Then
'                GetSaleCondition = SaleCondition & " or Left(sacode,4)='" & Left(SaleLogin, 4) & "'"
'            Else
'                GetSaleCondition = " Left(sacode,4)='" & Left(SaleLogin, 4) & "'"
'            End If
'    Else
'             '�觤׹������͹䢷����
'             If SaleCondition <> "" Then
'                GetSaleCondition = SaleCondition & " or Sacode ='" & SaleLogin & "'"
'            Else
'                GetSaleCondition = " Sacode ='" & SaleLogin & "'"
'            End If
'    End If
'
'    If GetSaleCondition <> "" Then GetSaleCondition = "(" + GetSaleCondition + ")"
'  End Function

'
'Public Sub SelectSale(GetBoss As String)
' Set rsFunction = New Adodb.Recordset
'  With rsFunction
'        strCmdSQL = "select  *  from salestb  where saboss like'%" & GetBoss & "%'"
'        .Open strCmdSQL, dbSQL, adOpenDynamic, adLockOptimistic, adCmdText
'        Do While Not .EOF
'                '���͹䢵�ҧ�ҡ������
'               If strCurrentNodeClick = "SALESTB1" Then
'                        If SaleCondition = "" Then
'                           SaleCondition = " Left(sacode,4) ='" & Left(rsFunction!SaCode, 4) & "'"
'                        Else
'                           SaleCondition = SaleCondition & " or Left(sacode,4) = '" & Left(rsFunction!SaCode, 4) & "'"
'                        End If
'                Else
'                        If SaleCondition = "" Then
'                           SaleCondition = "  sacode ='" & rsFunction!SaCode & "'"
'                        Else
'                           SaleCondition = SaleCondition & " or  sacode ='" & rsFunction!SaCode & "'"
'                        End If
'                End If
'                Call push_stack(rsFunction!SaCode, rsFunction!SABoss)
'                .MoveNext
'         Loop
'        .Close
'  End With
'  Set rsFunction = Nothing
'End Sub
'
'
'Public Function GetSaCodeCondition(SaleLogin As String) As String
'    top_stack = 0
'    SaleCondition = ""
'   Call SelectSale(SaleLogin)
'   Do While top_stack > 0
'      GetBoss = pop_stack()
'      Call SelectSale(GetBoss)
'   Loop
'   '�觤׹������͹䢷����
'   If SaleCondition <> "" Then
'      GetSaCodeCondition = SaleCondition & " or sacode='" & SaleLogin & "'"
'  Else
'      GetSaCodeCondition = " SaCode='" & SaleLogin & "'"
'  End If
'  End Function
'
'Sub push_stack(msale As String, mboss As String)
'top_stack = top_stack + 1
'Sale_Stack(top_stack, 1) = msale
'Sale_Stack(top_stack, 2) = mboss
'End Sub
'
'Function pop_stack() As String
'pop_stack = Sale_Stack(top_stack, 1)
'top_stack = top_stack - 1
'End Function


'Public Sub Create_Tmp_SalesTB(TempTime As String)
'
'        Screen.MousePointer = vbHourglass
'
'        Set cmdSp = New Adodb.Command
'        With cmdSp
'                .ActiveConnection = dbActive
'                .CommandType = adCmdStoredProc
'                .CommandText = "Crt_Tmp_SalesTB"
'
'                .Parameters.Append cmdSp.CreateParameter("TmpTime", adChar, adParamInput, 10)
'                .Parameters("TmpTime").Value = TempTime
'
'                .Parameters.Append cmdSp.CreateParameter("UserLevel", adChar, adParamInput, 1)
'                .Parameters("UserLevel").Value = UsrLevel
'
'                .Parameters.Append cmdSp.CreateParameter("SAcode", adChar, adParamInput, 10)
'                .Parameters("SAcode").Value = UsrSCode
'
'                .Parameters.Append cmdSp.CreateParameter("DPcode", adChar, adParamInput, 3)
'                .Parameters("DPcode").Value = UsrDept
'
'                .Execute
'
'        End With
'        Set cmdSp = Nothing
'
'        Screen.MousePointer = vbDefault
'
'End Sub


'
'Public Function CheckAreaMgr(Boss As String)
'
'        Dim StfCodeBoss  As String
'        StfCodeBoss = Left(Boss, InStr(Boss, "-") - 1)
'        CheckAreaMgr = (Find_Ret_Val("users", "User_Level", "STFCode='" & StfCodeBoss & "'") = "7")
'
'End Function


'Public Function LookupSales(strDPcode As String, Optional sCond As String) As String
'
'        Select Case UsrLevel
'        Case "A", "B", "C"
'                If Trim(strDPcode) = "MIS" Then
'                        LookupCondition = "1=1"
'                ElseIf Trim(strDPcode) <> "" Then
'                        LookupCondition = "DPcode = '" & strDPcode & "'"
'                Else
'                        LookupCondition = "1=1"
'                End If
'
'        Case "9", "5", "4", "1"
'                LookupCondition = "DPcode = '" & UsrDept & "'"
'                If (UsrDept = "MAT" Or UsrDept = "ACC") And Trim(strDPcode) = "" Then
'                        LookupCondition = ""
'                Else
'                        LookupCondition = "DPcode = '" & strDPcode & "'"
'                End If
'
'        Case "7"
'                LookupCondition = "SAboss = '" & UsrSCode & "' OR SAcode = '" & UsrSCode & "'"
'
'        Case Else
'                LookupCondition = "SAcode = '" & UsrSCode & "'"
'
'        End Select
'        If RTrim(sCond) <> "" Then
'                LookupCondition = "(" & LookupCondition & ") And " & sCond
'        End If
'        ActiveLookup = "VW_LOOKUP_SALESTB"
'        'LookupRetVal = ""
'        frmLookup.Show vbModal
'        LookupSales = Trim(LookupRetVal)
'
'End Function

Public Function Find_OilPrice(pYear As Integer, pMonth As Integer) As Double
        '17/09/2009 santi �ӡ�õ�Ǩ�ͺ����ա�û�͹��ҹ���ѹ���͹������������������ ����ѧ������͹��������͹��͹˹��
        If Count_Record("SalesMinimum_Factor", "SMYear=" & pYear & " and SMMonth=" & pMonth) <> 0 Then
                Find_OilPrice = Format(Find_Ret_Num("SalesMinimum_Factor", "SMOilPrice", "SMYear=" & pYear & " and SMMonth=" & pMonth), FReal)
        Else
Step_Black:
                If pMonth = 12 Then
                   pYear = pYear - 1
                   pMonth = 1
                Else
                   pMonth = pMonth - 1
                End If
                Find_OilPrice = Format(Find_Ret_Num("SalesMinimum_Factor", "SMOilPrice", "SMYear=" & pYear & " and SMMonth=" & pMonth), FReal)
                If Find_OilPrice = 0 Then GoTo Step_Black
        End If
End Function

Public Function Find_OilPriceFromWIO(ptxtDate As String) As Double
        '�Ҥ�ҹ���ѹ��ѹ����͹���觧ҹ ������ԡ��ҹ���ѹ ��Ҩҡ�Ҥҹ���ѹ�ҡ�ش���ԡ��ҹ���ѹ��ѹ��� ���ͧ�ҡ�ѧ������˹�ö������
        '�ִ�Ҥҹ���ѹ B5 �� ��ѡ ���ͧ�ҡ����觧ҹ������кط���¹ö ������������ö�������ҵ�ͧ�����ѹ������� �֧��ͧ���Ҥҹ���ѹ�ѹ����ش���ԡ��ҹ���ѹ
         Find_OilPriceFromWIO = MinValue("WKIssueOil", "OilPrice", "WIORealdate='" & CStrToDate(ptxtDate) & "' and oilprice<>0 and GTCode='01'")
        If Find_OilPriceFromWIO = 0 Then
            Find_OilPriceFromWIO = MinValue("WKIssueOil", "OilPrice", "WIORealDate=(select MAX(WIORealDate) from WKIssueOil where Oilprice<>0   and GTCode='01' )")
        End If
        Find_OilPriceFromWIO = Format(Find_OilPriceFromWIO, FReal)
        
End Function

Public Function Find_CurrentGaolinePrice(strGTCode As String, strDate As String) As Double
Select Case strGTCode
             Case "01"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT01_Price", "GPDate='" & strDate & "'")
             Case "02"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT02_Price", "GPDate='" & strDate & "'")
             Case "03"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT03_Price", "GPDate='" & strDate & "'")
             Case "04"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT04_Price", "GPDate='" & strDate & "'")
             Case "05"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT05_Price", "GPDate='" & strDate & "'")
             Case "06"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT06_Price", "GPDate='" & strDate & "'")
             Case "07"
                            Find_CurrentGaolinePrice = Find_Ret_Num("GasolinePrice", "GT07_Price", "GPDate='" & strDate & "'")

End Select
Find_CurrentGaolinePrice = Format(Find_CurrentGaolinePrice, FReal)
End Function


Public Sub Create_Tmp_SalesTB(TempTime As String)

        Screen.MousePointer = vbHourglass
                
        Set cmdSp = New Adodb.Command
        With cmdSp
                .ActiveConnection = dbActive
                .CommandType = adCmdStoredProc
                .CommandText = "Crt_Tmp_SalesTB"

                .Parameters.Append cmdSp.CreateParameter("TmpTime", adChar, adParamInput, 10)
                .Parameters("TmpTime").Value = TempTime

                .Parameters.Append cmdSp.CreateParameter("UserLevel", adChar, adParamInput, 1)
                .Parameters("UserLevel").Value = UsrLevel

                .Parameters.Append cmdSp.CreateParameter("SAcode", adChar, adParamInput, 10)
                .Parameters("SAcode").Value = UsrSCode

                .Parameters.Append cmdSp.CreateParameter("DPcode", adChar, adParamInput, 3)
                .Parameters("DPcode").Value = UsrDept

                .Execute
                
        End With
        Set cmdSp = Nothing
                
        Screen.MousePointer = vbDefault
        
End Sub

Public Function CDouble(vValue)
     If Trim(vValue) = "" Or Not IsNumeric(vValue) Then vValue = "0"
     CDouble = CDbl(vValue)
End Function

Public Function EndDayOfMonth(pDate As Date) As Integer
    Select Case Month(pDate)
                  Case 1, 3, 5, 7, 8, 10, 12
                        EndDayOfMonth = 31
                  Case 2
                        If Year(pDate) Mod 4 = 0 Then
                             EndDayOfMonth = 29
                        Else
                             EndDayOfMonth = 28
                        End If
                  Case Else
                        EndDayOfMonth = 30
    End Select
End Function



Public Sub Preview_Picture(strType As String, strFileName As String, strCsCode As String)
                With frmViewCustMap
                        Select Case UCase(strType)
                                      Case "MAP"
                                                .pic.Picture = LoadPicture(CustMapPath & Trim(strFileName))
                                                .Caption = "Ἱ���  " & Find_Ret_Val("Customer", "CSThiName", "CsCode='" & Trim(strCsCode) & "'") & "(" & strFileName & ")"
                                                .cmdPrint.Visible = (bGP_Mis Or bGP_Support)
                                      Case "PO"
                                                .pic.Picture = LoadPicture(CustPOPath & Trim(strFileName))
                                                .Stretch.Value = vbUnchecked
                                                .Caption = "PO  " & Find_Ret_Val("Customer", "CSThiName", "CsCode='" & Trim(strCsCode) & "'") & "(" & strFileName & ")"
                                                .cmdPrint.Visible = True
                        Case Else
                        End Select
                        .Show vbModal
                End With
End Sub
Public Sub PrintPictureToFitPage(pic As Picture)
    Dim PicRatio As Double
    Dim printerWidth As Double
    Dim printerHeight As Double
    Dim printerRatio As Double
    Dim printerPicWidth As Double
    Dim printerPicHeight As Double

    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation.
    If pic.Height >= pic.Width Then
        Printer.Orientation = vbPRORPortrait ' Taller than wide.
    Else
        Printer.Orientation = vbPRORLandscape ' Wider than tall.
    End If
    ' Calculate device independent Width-to-Height ratio for picture.
    PicRatio = pic.Width / pic.Height
    ' Calculate the dimentions of the printable area in HiMetric.
    printerWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHimetric)
    printerHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHimetric)
    ' Calculate device independent Width to Height ratio for printer.
    printerRatio = printerWidth / printerHeight
    ' Scale the output to the printable area.
    If PicRatio >= printerRatio Then
        ' Scale picture to fit full width of printable area.
        printerPicWidth = Printer.ScaleX(printerWidth, vbHimetric, Printer.ScaleMode)
        printerPicHeight = Printer.ScaleY(printerWidth / PicRatio, vbHimetric, Printer.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area.
        printerPicHeight = Printer.ScaleY(printerHeight, vbHimetric, Printer.ScaleMode)
        printerPicWidth = Printer.ScaleX(printerHeight * PicRatio, vbHimetric, Printer.ScaleMode)
    End If
    ' Print the picture using the PaintPicture method.
    Printer.PaintPicture pic, 0, 0, printerPicWidth, printerPicHeight
End Sub



'=============== Open PDF FIle =======================
Public Sub OpenPDF(strFileName As String, frmSend As Form)
    strFileName = Trim(strFileName)
    If ChkFileExist(strFileName) Then
        ShellExecute frmSend.hwnd, "Open", strFileName, vbNullString, "D:\", SW_SHOWNORMAL
    Else
       MsgBox "FileName Change Or Not Set Path.", vbCritical, "File Not Found."
    End If
End Sub
'=============== End Open Pdf FIle =======================
