VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MS SQL Server Script Generator"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnMail 
      Height          =   315
      Left            =   6780
      Picture         =   "main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Mail the script"
      Top             =   6220
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnSave 
      Height          =   315
      Left            =   6420
      Picture         =   "main.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save the script"
      Top             =   6220
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox txtScript 
      Height          =   1695
      Left            =   60
      TabIndex        =   13
      Top             =   4440
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"main.frx":0646
   End
   Begin VB.CommandButton cmdRelogin 
      Caption         =   "Re&login..."
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CheckBox chkScriptForWholeTable 
      Caption         =   "Script &whole table"
      Height          =   255
      Left            =   7020
      TabIndex        =   11
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Timer tmrNextGridFillTry 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7860
      Top             =   2940
   End
   Begin VB.CommandButton btnClearScript 
      Height          =   315
      Left            =   8220
      Picture         =   "main.frx":06C8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Clear the script"
      Top             =   6220
      Width           =   375
   End
   Begin VB.ListBox lstTables 
      Height          =   3180
      ItemData        =   "main.frx":07CA
      Left            =   120
      List            =   "main.frx":07CC
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtTableName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancleLoad 
      Cancel          =   -1  'True
      Caption         =   "S&top Load"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   4050
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1680
      TabIndex        =   5
      Top             =   2970
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton btnOptions 
      Height          =   315
      Left            =   7860
      Picture         =   "main.frx":07CE
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Options"
      Top             =   6220
      Width           =   375
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2580
      TabIndex        =   10
      Top             =   1020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblTableName 
      AutoSize        =   -1  'True
      Caption         =   "Table Data (Select the rows you want to script):"
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   3360
   End
   Begin VB.Label lblGenerateScript 
      AutoSize        =   -1  'True
      Caption         =   "Ú"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5400
      MouseIcon       =   "main.frx":08D0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Generate Script"
      Top             =   3900
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select &Table:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblGetTableData 
      AutoSize        =   -1  'True
      Caption         =   "Ø"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2580
      MouseIcon       =   "main.frx":0BDA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Get data from table"
      Top             =   1875
      Width           =   390
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuGenerateDeleteStatement 
         Caption         =   "Include &Delete Statement"
      End
      Begin VB.Menu mnuGeneratePrintStatement 
         Caption         =   "Inlcude &Print Statements"
      End
      Begin VB.Menu mnuOptionIncludeComments 
         Caption         =   "Include &Comments"
      End
      Begin VB.Menu mnuOptionTurnOffRowCount 
         Caption         =   "&Turn Off Row Count"
      End
      Begin VB.Menu mnuIdentityColumn 
         Caption         =   "Identity Column"
         Begin VB.Menu mnuOptionIdentity 
            Caption         =   "&Retain Identity column values in Script"
            Index           =   0
         End
         Begin VB.Menu mnuOptionIdentity 
            Caption         =   "&Allow new values for Identity columns"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowLoginDialogAtStartUp 
         Caption         =   "Show Login Dialog At Start Up"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_FINDSTRING = &H18F     'To search the item in list box
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194    'To set horizontal scroll bar in listbox
Private Const LB_SETCURSEL = &H186  'To select the item in list box without causing Click event

Private Const msREG_ROOT As String = "ScriptGen"

Private Const msREG_SECTION_SETTINGS As String = "Settings"
Private Const msREG_KEY_INCLUDE_DELETE As String = "IncludeDelete"
Private Const msREG_KEY_INCLUDE_PRINT As String = "IncludePrint"
Private Const msREG_KEY_LOGON_AT_STARTUP As String = "LogonAtStartUp"
Private Const msREG_KEY_INCLUDE_COMMENT As String = "IncludeComment"
Private Const msREG_KEY_TURN_OFF_ROW_COUNT As String = "TurnOffRowCount"
Private Const msREG_KEY_RETAIN_IDENTITY_VALUES As String = "RetainIdentityVals"

Private Const msREG_SECTION_PROG_INFO As String = "ProgInfo"
Private Const msREG_KEY_VERSION As String = "Version"
Private Const msREG_KEY_PATH As String = "Path"

Private Const msREG_SECTION_USAGE_STATS As String = "Usage"
Private Const msREG_KEY_USAGE_COUNT As String = "UsageCount"
Private Const msREG_KEY_LAST_USED As String = "LastUsed"
Private Const msREG_KEY_USAGE_INTERVAL_SUM As String = "UsageIntervalSum"

Const sDATE_FORMAT As String = "dd mmm yyyy hh:mm:ss AM/PM"

Dim oADCConn As ADODB.Connection
Dim rsTables As ADODB.Recordset     'Stores the list of available tables
Dim rsTableData As ADODB.Recordset  'Data in the table
Attribute rsTableData.VB_VarHelpID = -1
Dim mbLoadCanceled As Boolean
Dim mbScriptingCanceled As Boolean
Dim mbGridIsBeingFilled As Boolean
Dim mbIdentityColumnExist As Boolean

Dim mbOptionShowLogonDialog As Boolean
Dim mbOptionIncludeDeleteStatement As Boolean
Dim mbOptionIncludePrintStatement As Boolean
Dim mbOptionIncludeComment As Boolean
Dim mbOptionTurnOffRowCount As Boolean
Dim mbOptionRetainIdentityValues As Boolean

Dim mbIncludeHeaderComment As Boolean

Const msAPP_TITLE As String = "MS SQL Server Script Generator"
Const msSCRIPT_BOX_DEFAULT_TEXT As String = "Select the rows in the grid and press the down arrow button"

Dim cDatabaseServerName As String
Dim cDefaultDatabaseName As String
Dim cUserId As String
Dim cPassword As String
Dim cTableName As String
Dim mavPrimaryKeysForSelectedTable As Variant

Private Function GetFixedCaption() As String
    GetFixedCaption = msAPP_TITLE & "  " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Sub btnClearScript_Click()
    'Clear the script text box
    txtScript.Text = vbNullString
End Sub

Private Sub cmdCancleLoad_Click()
    'Set the flag to indicate that load is cancled
    mbLoadCanceled = True
    mbScriptingCanceled = True
End Sub


Private Sub btnOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu mnuOption, , btnOptions.Left + X, btnOptions.Top + Y
End Sub

Private Sub cmdRelogin_Click()
    On Error GoTo ERR_cmdRelogin_Click
    Call DoLogin
Exit Sub
ERR_cmdRelogin_Click:
    Call ShowError
End Sub

Private Sub cmdCopy_Click()
    Call Clipboard.SetText(txtScript, vbCFText)
End Sub

Private Sub Command1_Click()
    GetPrimaryKeyCols (lstTables.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancleLoad_Click
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    Call GetOptions
    Call UpdateProgInfo
    Call UpdateUsageStats
    
    'Initial settings from registry
    'Call GetConnectionString(cDatabaseServerName, cDefaultDatabaseName, cUserId, cPassword)
    
    'Do login
    Call DoLogin(mbOptionShowLogonDialog)
    
    txtScript.Text = msSCRIPT_BOX_DEFAULT_TEXT

Exit Sub
Err_Form_Load:
    ShowError
End Sub

Private Sub GetOptions()
    mbOptionIncludeDeleteStatement = GetAppOption(msREG_KEY_INCLUDE_DELETE, True)
    mbOptionIncludePrintStatement = GetAppOption(msREG_KEY_INCLUDE_PRINT, True)
    mbOptionShowLogonDialog = GetAppOption(msREG_KEY_LOGON_AT_STARTUP, True)
    mbOptionIncludeComment = GetAppOption(msREG_KEY_INCLUDE_COMMENT, True)
    mbOptionTurnOffRowCount = GetAppOption(msREG_KEY_TURN_OFF_ROW_COUNT, True)
    mbOptionRetainIdentityValues = GetAppOption(msREG_KEY_RETAIN_IDENTITY_VALUES, True)
    mbIncludeHeaderComment = True
End Sub

Private Sub UpdateUsageStats()

    Dim lUsageCount As Long
    Dim lUsageIntervalSum As Long
    Dim dtLastUsed As Date

    lUsageCount = GetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_USAGE_COUNT, 0)
    lUsageCount = lUsageCount + 1
    Call SetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_USAGE_COUNT, lUsageCount)
    
    dtLastUsed = GetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_LAST_USED, Format(0, sDATE_FORMAT))
    If Format(dtLastUsed, sDATE_FORMAT) = Format(0, sDATE_FORMAT) Then
        Call SetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_USAGE_INTERVAL_SUM, 0)
    Else
        lUsageIntervalSum = GetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_USAGE_INTERVAL_SUM, 0)
        lUsageIntervalSum = lUsageIntervalSum + DateDiff("n", dtLastUsed, Now)
        Call SetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_USAGE_INTERVAL_SUM, lUsageIntervalSum)
    End If
    Call SetAppKey(msREG_SECTION_USAGE_STATS, msREG_KEY_LAST_USED, Format(Now, sDATE_FORMAT))

End Sub

Private Sub UpdateProgInfo()
    Dim dblVer As Double
    Dim sAppPath As String
    
    dblVer = GetAppKey(msREG_SECTION_PROG_INFO, msREG_KEY_VERSION, 0)
    If dblVer <> CDbl(App.Major & "." & App.Minor) + (App.Revision / 10000) Then
        Call SetAppKey(msREG_SECTION_PROG_INFO, msREG_KEY_VERSION, CDbl(App.Major & "." & App.Minor) + (App.Revision / 10000))
    End If
    
    sAppPath = GetAppKey(msREG_SECTION_PROG_INFO, msREG_KEY_PATH, "")
    If sAppPath <> App.Path Then
        Call SetAppKey(msREG_SECTION_PROG_INFO, msREG_KEY_PATH, App.Path)
    End If
End Sub

Private Sub DoLogin(Optional ByVal vboolShowDialog As Boolean = True)
    
    On Error GoTo ERR_DoLogin
    
    Dim ofrmLogin As frmLogin
    Dim bLoginDialogReturn As Boolean
    
    'Set the caption
    If oADCConn Is Nothing Then
        Me.Caption = GetFixedCaption & " - Not Connected"
    End If
    
    If vboolShowDialog Then
        Set ofrmLogin = New frmLogin
        bLoginDialogReturn = ofrmLogin.DisplayForm(cDatabaseServerName, cDefaultDatabaseName, cUserId, cPassword)
        Set ofrmLogin = Nothing
    Else
        bLoginDialogReturn = True
    End If

    If bLoginDialogReturn Then
        Call InitTableList
        'Set the caption
        Me.Caption = GetFixedCaption & " - " & cDatabaseServerName & "/" & cDefaultDatabaseName & " as " & cUserId
    Else
        'If not connected the tell user that he is not logged in
        If oADCConn Is Nothing Then
            Err.Raise 1000, , "User cancelled login. Can not connect to database."
        End If
    End If
   
Exit Sub
ERR_DoLogin:
    Call SaveErrorObj
    Call CleanUpOnDisconnect
    Call ReRaiseError(True)
End Sub

Private Sub CleanUpOnDisconnect()
    Call cmdCancleLoad_Click
    Set oADCConn = Nothing
    Me.Caption = GetFixedCaption & " - Not Connected"
    lstTables.Clear
    MSFlexGrid1.Rows = 0
    MSFlexGrid1.Cols = 0
    txtTableName.Text = vbNullString
End Sub

Private Sub InitTableList()
    'Get the list of tables
    
    Me.MousePointer = vbHourglass
    
    Dim cConnection As String
    
    cConnection = MakeConnectString(cDatabaseServerName, cDefaultDatabaseName, cUserId, cPassword)
    
    Set oADCConn = Nothing
    
    Set oADCConn = New ADODB.Connection
    With oADCConn
        .Open cConnection
        Set rsTables = .OpenSchema(adSchemaTables)
    End With
    
    'Fill the listbox with the names of tables
    Call FillComboOrList(rsTables, lstTables, "TABLE_NAME")
    
    MSFlexGrid1.Rows = 0
    MSFlexGrid1.Cols = 0

    txtTableName.Text = "Enter table name"
    txtTableName.SelStart = 0
    txtTableName.SelLength = Len(txtTableName.Text)
    
    'Show the horizontal scroll bar in list box
    Call SendMessage(lstTables.hwnd, LB_SETHORIZONTALEXTENT, 1000, 0)
    
    Me.MousePointer = vbDefault
    
End Sub

'Fill the grid with content of the recordset
Private Sub FillGrid(ByVal vrsRecordSet As ADODB.Recordset, ByVal vflxGrid As MSFlexGrid)
    Dim lRows As Long
    Dim lCols As Long
    Dim lRowIndex As Long
    Dim lColIndex As Long
    Dim sColString As String
    
    mbGridIsBeingFilled = True
    
    lblProgress.Visible = True
    
    With vflxGrid
    
        'If this flag is set to True from outside, For loops are terminated
        mbLoadCanceled = False
        
        'Disable the redraw while setting up the columns
        .Redraw = False
        
        'Set the numbers of columns
        lCols = vrsRecordSet.Fields.Count
        .Cols = lCols + .FixedCols
        
        'Set numbers of rows to atleast FixedRows + 1
        .Rows = 2
        .FixedRows = 1
        'After setting FixedRows, numbers of row can be reduced to minimum
        
        'Set numbers of columns
        .FixedCols = 0
        
        .Row = 0
        
        mbIdentityColumnExist = False
        
        'Set up the column headers
            For lColIndex = 0 To lCols - 1
                .Col = lColIndex + .FixedCols
                With vrsRecordSet.Fields.Item(lColIndex)
                    vflxGrid.Text = .Name
                    vflxGrid.CellFontBold = IsPrimaryKey(.Name)
                    If .Properties("ISAUTOINCREMENT").Value = True Then
                        vflxGrid.CellFontItalic = True
                        mbIdentityColumnExist = True
                    Else
                        vflxGrid.CellFontItalic = False
                    End If
                End With
            Next lColIndex
                
        'Get the numbers of rows in recordset
        If Not (vrsRecordSet.BOF And vrsRecordSet.EOF) Then
            vrsRecordSet.MoveLast
            vrsRecordSet.MoveFirst
            lRows = vrsRecordSet.RecordCount
        Else
            lRows = 0
        End If
        
        'Show the grid now
        .Redraw = True
        
        'Set the total numbers of rows = 1 which is fixed row
        .Rows = 1
        
        'If numbers of rows in recordset is non zero
        If lRows <> 0 Then
        
            vrsRecordSet.MoveFirst
            
            'Scan all the rows
            For lRowIndex = 0 To lRows - 1
            
                'This string holds the data for whole row
                sColString = vbNullString
                
                'Get the string that contains data for row in recordset
                For lColIndex = 0 To lCols - 1
                    sColString = sColString & IIf(IsNull(vrsRecordSet.Fields(lColIndex)), "(null)", vrsRecordSet.Fields(lColIndex)) & vbTab
                Next lColIndex
                                
                'Set the data in the grid
                Call vflxGrid.AddItem(sColString)
                
                Call ShowProgress(lRowIndex, lRows, "Records Loaded")
                
                'Give chance to user to cancel the loading
                DoEvents
                
                'If load was cancled, exit the loop
                If mbLoadCanceled Then Exit For
                
                'Go to next record
                vrsRecordSet.MoveNext
            
            Next lRowIndex
            
        End If
        
    End With

    mbGridIsBeingFilled = False
    lblProgress.Visible = False
End Sub

Private Sub FillComboOrList(ByVal vrsRecordSet As ADODB.Recordset, ByVal vcmbCombo As Object, ByVal vvColumnNameOrIndexForItem As Variant, Optional ByVal vvColumnNameOrIndexForItemData As Variant = Empty)
    
    Dim lRows As Long
    Dim lRowIndex As Long
    
    'Get the numbers of rows in recordset
    If Not (vrsRecordSet.BOF And vrsRecordSet.EOF) Then
        vrsRecordSet.MoveLast
        vrsRecordSet.MoveFirst
        lRows = vrsRecordSet.RecordCount
    Else
        lRows = 0
    End If
    
    vcmbCombo.Clear
    
    'Fill the content in the combo
    For lRowIndex = 0 To lRows - 1
        Call vcmbCombo.AddItem(vrsRecordSet.Fields(vvColumnNameOrIndexForItem) & "")
        If Not IsEmpty(vvColumnNameOrIndexForItemData) Then
            vcmbCombo.ItemData(vcmbCombo.NewIndex) = vrsRecordSet.Fields(vvColumnNameOrIndexForItemData) & ""
        End If
        vrsRecordSet.MoveNext
    Next lRowIndex
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Stop loading the grid if it's being loaded
    mbLoadCanceled = True
    mbScriptingCanceled = True
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oADCConn = Nothing
    Set rsTables = Nothing
    
    'Stop loading the grid if it's being loaded
    DoEvents
    
    'Whatever might be, terminate the program!
    End
End Sub

Private Sub lblGetTableData_Click()

    On Error GoTo ERR_lblGetTableData_Click
    
    If mbGridIsBeingFilled Then
        mbLoadCanceled = True
        tmrNextGridFillTry.Enabled = True
        Exit Sub
    End If
    
    tmrNextGridFillTry.Enabled = False
    
    Set rsTableData = Nothing
    
    If lstTables.Text <> vbNullString Then
    
        Set rsTableData = New ADODB.Recordset
        Me.MousePointer = vbHourglass
        cTableName = lstTables.Text
        
        Dim sSQL As String
        Dim sOrderByClause As String
        
        sOrderByClause = vbNullString
        sSQL = "Select * from " & lstTables.Text
        
        mavPrimaryKeysForSelectedTable = GetPrimaryKeyCols(lstTables.Text)
        If Not IsEmpty(mavPrimaryKeysForSelectedTable) Then
            'Form comma seperated list for ORDER BY clause
            Dim lPrimaryKeyIndex As Long
            Dim lPrimaryKeyCount As Long
            lPrimaryKeyCount = UBound(mavPrimaryKeysForSelectedTable)
            For lPrimaryKeyIndex = LBound(mavPrimaryKeysForSelectedTable) To lPrimaryKeyCount
                sOrderByClause = sOrderByClause & mavPrimaryKeysForSelectedTable(lPrimaryKeyIndex)
                If lPrimaryKeyIndex <> lPrimaryKeyCount Then
                    sOrderByClause = sOrderByClause & ", "
                End If
            Next lPrimaryKeyIndex
        End If
        
        If sOrderByClause <> vbNullString Then
            sOrderByClause = " ORDER BY " & sOrderByClause
        End If
        sSQL = sSQL & sOrderByClause
        
        'Do not use adOpenStatic because client side curor will fali to open if total row size exceeds some limit
        Screen.MousePointer = vbHourglass
        Call rsTableData.Open(sSQL, oADCConn, adOpenKeyset, adLockReadOnly, adCmdText)
        
        Screen.MousePointer = vbDefault
        
        lblTableName.Caption = "Table Data For: " & lstTables.Text
        
        Call FillGrid(rsTableData, MSFlexGrid1)
       
    Else
    
        Err.Raise 1000, , "No table selected or no database connection."
        
    End If
    
Exit Sub
ERR_lblGetTableData_Click:
    Call ShowError
End Sub

Private Sub lblGenerateScript_Click()

    On Error GoTo ERR_lblGenerateScript_Click

    Dim bContinueScripting As Boolean
    
    bContinueScripting = True
    
    Randomize
    If mbGridIsBeingFilled Then
        Dim mrMessageResponse As VbMsgBoxResult
        mrMessageResponse = MsgBox("Stop the data loading and continue with scripting?", vbYesNo, "ScriptGen")
        If mrMessageResponse = vbYes Then
            mbLoadCanceled = True
            DoEvents
            DoEvents
        Else
            bContinueScripting = False
        End If
    End If
    
    If bContinueScripting Then

        Dim lRowIndex As Long
        Dim lColIndex As Long
        Dim lRowStart As Long
        Dim lColStart As Long
        Dim lRowEnd As Long
        Dim lColEnd As Long
        Dim sInsertStatement As String
        Dim sColNames As String
        Dim sFieldValue As String
        Dim vFieldValue As Variant
        Dim lStartingTextLength As Long
        
        Dim bRowCountTurnedOff As Boolean
        bRowCountTurnedOff = False
        
        If txtScript.Text = msSCRIPT_BOX_DEFAULT_TEXT Then
            txtScript.Text = vbNullString
        End If
        
        If Not (rsTableData Is Nothing) Then
        
            If Not ((rsTableData.EOF = True) And (rsTableData.BOF = True)) Then
        
                Me.MousePointer = vbHourglass
                
                With MSFlexGrid1
                
                    If mbOptionIncludePrintStatement Then
                        If mbIncludeHeaderComment Then
                            txtScript.Text = txtScript.Text & "/* Script generated by " & GetCurUserLoginName & " using ScriptGen Ver " & App.Major & "." & App.Minor & "." & App.Revision & " on " & Format(Now, sDATE_FORMAT) & " */" & vbCrLf & vbCrLf
                            mbIncludeHeaderComment = False
                        End If
                        txtScript.Text = txtScript.Text & "--------------------------------------------------------------------------------" & vbCrLf
                        txtScript.Text = txtScript.Text & "PRINT ""Inserting data in table '" & cTableName & "' ...""" & vbCrLf
                        txtScript.Text = txtScript.Text & "GO" & vbCrLf & vbCrLf
                    End If
                    If mbOptionTurnOffRowCount Then
                        txtScript.Text = txtScript.Text & "SET NOCOUNT ON" & vbCrLf & vbCrLf
                        bRowCountTurnedOff = True
                    End If
                    
                    If chkScriptForWholeTable.Value = vbUnchecked Then
                
                        If .Row > .RowSel Then
                            lRowStart = .RowSel
                            lRowEnd = .Row
                        Else
                            lRowStart = .Row
                            lRowEnd = .RowSel
                        End If
                        
                        If mbOptionIncludeDeleteStatement Then
                            'Generate the DELETE statement if primary keys are available
                            If Not IsEmpty(mavPrimaryKeysForSelectedTable) Then
                                txtScript.Text = txtScript.Text & vbCrLf
                                
                                If mbOptionIncludeComment Then
                                    txtScript.Text = txtScript.Text & "/* Delete the old rows */" & vbCrLf
                                End If
                                
                                txtScript.Text = txtScript.Text & "DELETE " & cTableName & vbCrLf & "WHERE "
                                
                                Dim lPrimaryKeyIndex As Long
                                Dim sPrimaryKeyFieldName As String
                                Dim lPrimaryKeyCount As Long
                                Dim sWhereClausePart1 As String
                                Dim sWhereClausePart2 As String
                                Dim sWhereClausePart3 As String
                                Dim sComparisonOperator As String
                                
                                lPrimaryKeyCount = UBound(mavPrimaryKeysForSelectedTable)
                                sWhereClausePart3 = "("
                                
                                'Make the part 1 clause
                                Call rsTableData.Move(lRowStart - 1, adBookmarkFirst)
                                sWhereClausePart1 = "("
                                For lPrimaryKeyIndex = 1 To lPrimaryKeyCount
                                    sPrimaryKeyFieldName = mavPrimaryKeysForSelectedTable(lPrimaryKeyIndex)
                                    If lPrimaryKeyIndex = 1 Then
                                        sComparisonOperator = " = "
                                        sWhereClausePart3 = sWhereClausePart3 & "(" & sPrimaryKeyFieldName & " > " & SQLFieldValueAsString(rsTableData(sPrimaryKeyFieldName), rsTableData(sPrimaryKeyFieldName).Type) & ")"
                                    Else
                                        sComparisonOperator = " >= "
                                    End If
                                    sWhereClausePart1 = sWhereClausePart1 & "(" & sPrimaryKeyFieldName & sComparisonOperator & SQLFieldValueAsString(rsTableData(sPrimaryKeyFieldName), rsTableData(sPrimaryKeyFieldName).Type) & ")"
                                    If lPrimaryKeyIndex < lPrimaryKeyCount Then
                                        sWhereClausePart1 = sWhereClausePart1 & " AND "
                                    End If
                                Next lPrimaryKeyIndex
                                sWhereClausePart1 = sWhereClausePart1 & ")"
                                
                                'Make the part 2 clause
                                Call rsTableData.Move(lRowEnd - 1, adBookmarkFirst)
                                sWhereClausePart2 = "("
                                For lPrimaryKeyIndex = 1 To lPrimaryKeyCount
                                    sPrimaryKeyFieldName = mavPrimaryKeysForSelectedTable(lPrimaryKeyIndex)
                                    If lPrimaryKeyIndex = 1 Then
                                        sComparisonOperator = " = "
                                        sWhereClausePart3 = sWhereClausePart3 & " AND " & "(" & sPrimaryKeyFieldName & " < " & SQLFieldValueAsString(rsTableData(sPrimaryKeyFieldName), rsTableData(sPrimaryKeyFieldName).Type) & ")"
                                    Else
                                        sComparisonOperator = " <= "
                                    End If
                                    sWhereClausePart2 = sWhereClausePart2 & "(" & sPrimaryKeyFieldName & sComparisonOperator & SQLFieldValueAsString(rsTableData(sPrimaryKeyFieldName), rsTableData(sPrimaryKeyFieldName).Type) & ")"
                                    If lPrimaryKeyIndex <> lPrimaryKeyCount Then
                                        sWhereClausePart2 = sWhereClausePart2 & " AND "
                                    End If
                                Next lPrimaryKeyIndex
                                sWhereClausePart2 = sWhereClausePart2 & ")"
                                
                                sWhereClausePart3 = sWhereClausePart3 & ")"
                                
                                txtScript.Text = txtScript.Text & sWhereClausePart1 & vbCrLf & "OR" & sWhereClausePart2 & vbCrLf & "OR " & sWhereClausePart3 & vbCrLf
                                
                            Else
                                txtScript.Text = txtScript.Text & vbCrLf
                                txtScript.Text = txtScript.Text & "/* Can't generate DELETE statement because no primary keys in table found */" & vbCrLf
                            End If
                            txtScript.Text = txtScript.Text & vbCrLf & vbCrLf
                        End If
                    Else
                        lRowStart = 1
                        lRowEnd = .Rows - 1
                        
                        txtScript.Text = txtScript.Text & vbCrLf
                        If mbOptionIncludeDeleteStatement Then
                            txtScript.Text = txtScript.Text & "DELETE " & cTableName & vbCrLf
                        End If
                        txtScript.Text = txtScript.Text & vbCrLf & vbCrLf
                    End If
                    
                    lColStart = 0
                    lColEnd = rsTableData.Fields.Count - 1
                    
                    sColNames = vbNullString
                    Dim bSkipColumn As Boolean
                    
                    For lColIndex = lColStart To lColEnd
                        bSkipColumn = False
                        If (Not mbOptionRetainIdentityValues) And mbIdentityColumnExist Then
                            If rsTableData.Fields(lColIndex).Properties("ISAUTOINCREMENT").Value = True Then
                                bSkipColumn = True
                            End If
                        End If
                        If Not bSkipColumn Then
                            If sColNames <> vbNullString Then
                                sColNames = sColNames & ", "
                            End If
                            sColNames = sColNames & rsTableData.Fields(lColIndex).Name
                        End If
                    Next lColIndex
                    
                    If mbOptionIncludeComment Then
                        txtScript.Text = txtScript.Text & "/* Insert the data */" & vbCrLf
                    End If
                    
                    If mbOptionRetainIdentityValues And mbIdentityColumnExist Then
                        txtScript.Text = txtScript.Text & vbCrLf & "SET IDENTITY_INSERT " & cTableName & " ON" & vbCrLf & vbCrLf
                    End If
                    
                    '****Main loop to make INSERT statements
                    mbScriptingCanceled = False
                    Call DisableControlsWhileProcessing(True)
                    lblProgress.Visible = True
                    DoEvents
                    
                    Dim sAllInsertStatements As String
                    sAllInsertStatements = vbNullString
                    lStartingTextLength = 0
                    Call rsTableData.Move(lRowStart - 1, adBookmarkFirst)
                    For lRowIndex = lRowStart To lRowEnd
                        sInsertStatement = "INSERT " & cTableName & " (" & sColNames & ") VALUES ("
                        
                        Dim bFieldAdded As Boolean
                        bFieldAdded = False
                        For lColIndex = lColStart To lColEnd
                            bSkipColumn = False
                            If (Not mbOptionRetainIdentityValues) And mbIdentityColumnExist Then
                                If rsTableData.Fields(lColIndex).Properties("ISAUTOINCREMENT").Value = True Then
                                    bSkipColumn = True
                                End If
                            End If
                            If Not bSkipColumn Then
                                If bFieldAdded Then
                                    sInsertStatement = sInsertStatement & ", "
                                End If
                                bFieldAdded = True
                                vFieldValue = rsTableData.Fields(lColIndex).Value
                                sFieldValue = SQLFieldValueAsString(vFieldValue, rsTableData.Fields(lColIndex).Type)
                                sInsertStatement = sInsertStatement & sFieldValue
                            End If
                        Next lColIndex
                        sInsertStatement = sInsertStatement & ")"
                        sAllInsertStatements = sAllInsertStatements & sInsertStatement & vbCrLf
                        If (Len(sAllInsertStatements) - lStartingTextLength) > (2 ^ 15) Then
                            sAllInsertStatements = sAllInsertStatements & "GO" & vbCrLf
                            lStartingTextLength = Len(sAllInsertStatements)
                        End If
                        If mbScriptingCanceled Then Exit For
                        rsTableData.MoveNext
                        Call ShowProgress(lRowIndex - lRowStart + 1, lRowEnd - lRowStart + 1, "Records Scripted")
                        DoEvents
                    Next lRowIndex
                    lblProgress.Visible = False
                    Call DisableControlsWhileProcessing(False)
                    '****Main loop ends
                    
                    txtScript.Text = txtScript.Text & sAllInsertStatements
                    
                    If mbOptionRetainIdentityValues And mbIdentityColumnExist Then
                        txtScript.Text = txtScript.Text & vbCrLf & "SET IDENTITY_INSERT " & cTableName & " OFF" & vbCrLf
                    End If
                    
                    Me.MousePointer = vbDefault

                    
                End With
                
                If mbOptionIncludePrintStatement Or bRowCountTurnedOff Then
                    txtScript.Text = txtScript.Text & vbCrLf & "GO" & vbCrLf & vbCrLf
                End If
                If bRowCountTurnedOff Then
                    txtScript.Text = txtScript.Text & "SET NOCOUNT OFF" & vbCrLf
                End If
                If mbOptionIncludePrintStatement Then
                    txtScript.Text = txtScript.Text & "Print ""Data insetion in table '" & cTableName & "' completed.""" & vbCrLf & "Print ''" & vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf
                End If
                Me.MousePointer = vbDefault
            Else
                Err.Raise 1000, , "There is no data in table to script!"
            End If
        Else
            Err.Raise 1000, , "No connection to database. Please do login first."
        End If
    End If
    
Exit Sub
ERR_lblGenerateScript_Click:
    Call ShowError
End Sub

Private Sub lstTables_Click()
    txtTableName.Text = lstTables.Text
    Call lblGetTableData_Click
End Sub

Private Sub mnuGenerateDeleteStatement_Click()
    mnuGenerateDeleteStatement.Checked = Not mnuGenerateDeleteStatement.Checked
    mbOptionIncludeDeleteStatement = mnuGenerateDeleteStatement.Checked
    Call SetAppOption(msREG_KEY_INCLUDE_DELETE, mbOptionIncludeDeleteStatement)
End Sub

Private Sub mnuGeneratePrintStatement_Click()
    mnuGeneratePrintStatement.Checked = Not mnuGeneratePrintStatement.Checked
    mbOptionIncludePrintStatement = mnuGeneratePrintStatement.Checked
    Call SetAppOption(msREG_KEY_INCLUDE_PRINT, mbOptionIncludePrintStatement)
End Sub

Private Sub mnuOption_Click()
    mnuGenerateDeleteStatement.Checked = mbOptionIncludeDeleteStatement
    mnuGeneratePrintStatement.Checked = mbOptionIncludePrintStatement
    mnuShowLoginDialogAtStartUp.Checked = mbOptionShowLogonDialog
    mnuOptionIncludeComments.Checked = mbOptionIncludeComment
    mnuOptionTurnOffRowCount.Checked = mbOptionTurnOffRowCount
End Sub

Private Sub mnuOptionIdentity_Click(Index As Integer)
    Const mlRETAIN_IDENTITY As Long = 0
    
    mnuOptionIdentity(0).Checked = Not mnuOptionIdentity(0).Checked
    mnuOptionIdentity(1).Checked = Not mnuOptionIdentity(1).Checked
    
    mbOptionRetainIdentityValues = mnuOptionIdentity(0).Checked
    
    Call SetAppOption(msREG_KEY_RETAIN_IDENTITY_VALUES, mbOptionRetainIdentityValues)
    
End Sub

Private Sub mnuOptionIncludeComments_Click()
    mnuOptionIncludeComments.Checked = Not mnuOptionIncludeComments.Checked
    mbOptionIncludeComment = mnuOptionIncludeComments.Checked
    Call SetAppOption(msREG_KEY_INCLUDE_COMMENT, mbOptionIncludeComment)
End Sub

Private Sub mnuOptionTurnOffRowCount_Click()
    mnuOptionTurnOffRowCount.Checked = Not mnuOptionTurnOffRowCount.Checked
    mbOptionTurnOffRowCount = mnuOptionTurnOffRowCount.Checked
    Call SetAppOption(msREG_KEY_TURN_OFF_ROW_COUNT, mbOptionTurnOffRowCount)
End Sub

Private Sub mnuShowLoginDialogAtStartUp_Click()
    mnuShowLoginDialogAtStartUp.Checked = Not mnuShowLoginDialogAtStartUp.Checked
    mbOptionShowLogonDialog = mnuShowLoginDialogAtStartUp.Checked
    Call SetAppOption(msREG_KEY_LOGON_AT_STARTUP, mbOptionShowLogonDialog)
End Sub

Private Sub tmrNextGridFillTry_Timer()
    Call lblGetTableData_Click
End Sub

Private Sub txtScript_Change()
    If txtScript.Text = vbNullString Then
        mbIncludeHeaderComment = True
    End If
End Sub

Private Sub txtScript_GotFocus()
    txtScript.SelStart = 0
    txtScript.SelLength = Len(txtScript.Text)
End Sub

Private Sub txtTableName_Change()
    
    On Error GoTo ERR_txtTableName_Change

    Dim lListIndex As Long
    lListIndex = SendMessage(lstTables.hwnd, LB_FINDSTRING, -1, ByVal CStr(txtTableName.Text))
    Call SendMessage(lstTables.hwnd, LB_SETCURSEL, lListIndex, 0)
    'lstTables.ListIndex = lListIndex
    
Exit Sub
ERR_txtTableName_Change:
    Call ShowError
End Sub

Private Sub ShowError()
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    lblProgress.Visible = False
    mbGridIsBeingFilled = False
    mbLoadCanceled = False
    mbScriptingCanceled = False
    Call DisableControlsWhileProcessing(False)
    MsgBox "Error " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Sub

Private Function IsGridBeingFilled() As Boolean
    
    'Set the flag to false
    mbGridIsBeingFilled = False
    
    'Let the grid routine to run
    DoEvents
    
    IsGridBeingFilled = mbGridIsBeingFilled
    
End Function

Private Sub txtTableName_GotFocus()
    txtTableName.SelStart = 0
    txtTableName.SelLength = Len(txtTableName.Text)
End Sub

Private Sub txtTableName_KeyPress(KeyAscii As Integer)
    'If Enter pressed
    If KeyAscii = 13 Then
        'Prevent beep
        KeyAscii = 0
        'Show table data
        lblGetTableData_Click
    End If
End Sub

Private Function GetPrimaryKeyCols(ByVal vsTableName As String) As Variant
        
    Dim rsPrimaryKeys As Recordset
    Me.MousePointer = vbHourglass
    Set rsPrimaryKeys = oADCConn.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, vsTableName))
    If Not (rsPrimaryKeys.EOF And rsPrimaryKeys.EOF) Then
        rsPrimaryKeys.MoveLast
        rsPrimaryKeys.MoveFirst
    End If
    Me.MousePointer = vbDefault
    
    Dim avPrimaryKeys As Variant
    avPrimaryKeys = Empty
    
    If rsPrimaryKeys.RecordCount <> 0 Then
        ReDim avPrimaryKeys(1 To rsPrimaryKeys.RecordCount)
    End If
    
    Dim lPrimaryKeyIndex As Long
    
    For lPrimaryKeyIndex = 1 To rsPrimaryKeys.RecordCount
        avPrimaryKeys(lPrimaryKeyIndex) = rsPrimaryKeys(3)
        rsPrimaryKeys.MoveNext
    Next lPrimaryKeyIndex
    
    GetPrimaryKeyCols = avPrimaryKeys
End Function

Private Function SQLFieldValueAsString(ByVal vvFieldValue As Variant, ByVal venmDataType As ADODB.DataTypeEnum) As String
    
    Dim sFieldValue As String
    
    If Not IsNull(vvFieldValue) Then
        Select Case venmDataType
            Case adChar, adBSTR, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
                sFieldValue = "'" & ReplaceString(vvFieldValue, "'", "''") & "'"
            Case adDate, adDBDate
                sFieldValue = "Convert(DateTime,'" & Format$(vvFieldValue, "dd mmm yyyy") & "')"
            Case adDBTime, adDBTimeStamp
                sFieldValue = "Convert(DateTime,'" & Format$(vvFieldValue, "dd mmm yyyy hh:nn:ss") & "')"
            Case Else
                sFieldValue = vvFieldValue
        End Select
    Else
        sFieldValue = "Null"
    End If
    
    SQLFieldValueAsString = sFieldValue
    
End Function

Private Function IsPrimaryKey(ByVal vsFieldName As String) As Boolean
    Dim lPrimaryKeyIndex As Long
    IsPrimaryKey = False
    If Not IsEmpty(mavPrimaryKeysForSelectedTable) Then
        For lPrimaryKeyIndex = LBound(mavPrimaryKeysForSelectedTable) To UBound(mavPrimaryKeysForSelectedTable)
            If LCase(vsFieldName) = LCase(mavPrimaryKeysForSelectedTable(lPrimaryKeyIndex)) Then
                IsPrimaryKey = True
                Exit For
            End If
        Next lPrimaryKeyIndex
    End If
End Function

Private Function GetAppOption(ByVal vsKeyId As String, Optional vvDefault As Variant = Empty) As Variant
    GetAppOption = GetSetting(msREG_ROOT, msREG_SECTION_SETTINGS, vsKeyId, vvDefault)
End Function

Private Sub SetAppOption(ByVal vsKeyId As String, ByVal vvValue As Variant)
    Call SaveSetting(msREG_ROOT, msREG_SECTION_SETTINGS, vsKeyId, vvValue)
End Sub

Private Function GetAppKey(ByVal vsSection As String, ByVal vsKeyId As String, Optional vvDefault As Variant = Empty) As Variant
    GetAppKey = GetSetting(msREG_ROOT, vsSection, vsKeyId, vvDefault)
End Function

Private Sub SetAppKey(ByVal vsSection As String, ByVal vsKeyId As String, ByVal vvValue As Variant)
    Call SaveSetting(msREG_ROOT, vsSection, vsKeyId, vvValue)
End Sub

Private Function CheckToBool(ByVal venmCheckBoxValue As CheckBoxConstants) As Boolean
    If venmCheckBoxValue = vbChecked Then
        CheckToBool = True
    Else
        CheckToBool = False
    End If
End Function

Private Function BoolToCheck(ByVal vboolValue As Boolean) As CheckBoxConstants
    If vboolValue Then
        BoolToCheck = vbChecked
    Else
        BoolToCheck = vbUnchecked
    End If
End Function

Private Sub ShowProgress(ByVal vdblCurrentValue As Double, ByVal vdblEndValue As Double, ByVal vsToolTipPrefix As String)
    lblProgress.Caption = Chr(183 + CLng(10 * (vdblCurrentValue / vdblEndValue)))
    lblProgress.ToolTipText = vsToolTipPrefix & " : " & CLng(vdblCurrentValue) & " of " & CLng(vdblEndValue)
End Sub

Private Sub DisableControlsWhileProcessing(ByVal vboolDisable As Boolean)
    txtScript.Enabled = Not vboolDisable
    txtTableName.Enabled = Not vboolDisable
    lstTables.Enabled = Not vboolDisable
    MSFlexGrid1.Enabled = Not vboolDisable
    btnClearScript.Enabled = Not vboolDisable
    btnMail.Enabled = Not vboolDisable
    btnSave.Enabled = Not vboolDisable
    cmdRelogin.Enabled = Not vboolDisable
    lblGenerateScript.Enabled = Not vboolDisable
    lblGetTableData.Enabled = Not vboolDisable
End Sub

