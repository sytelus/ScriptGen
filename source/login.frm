VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Server Login"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1260
      TabIndex        =   10
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1035
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1620
      Width           =   2235
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   2235
   End
   Begin VB.TextBox txtDBName 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2235
   End
   Begin VB.TextBox txtServerName 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Login:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Database:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Server Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbIsOkPressed As Boolean

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbIsOkPressed = True
    Me.Hide
End Sub

Public Function DisplayForm(ByRef rsServerName As String, ByRef rsDatabaseName As String, ByRef rsLogin As String, ByRef rsPassword As String) As Boolean
    
    Dim bFormValidated As Boolean
    Dim sMessage As String
    
    txtServerName.Text = rsServerName
    txtDBName.Text = rsDatabaseName
    txtLogin.Text = rsLogin
    txtPassword.Text = rsPassword
    
    bFormValidated = False
    
    Do While Not bFormValidated
    
        mbIsOkPressed = False
        
        Me.Show vbModal
        
        If mbIsOkPressed Then
            bFormValidated = ValidateForm(sMessage)
            If bFormValidated Then
                rsServerName = txtServerName.Text
                rsDatabaseName = txtDBName.Text
                rsLogin = txtLogin.Text
                rsPassword = txtPassword.Text
            Else
                MsgBox sMessage
            End If
        Else
            bFormValidated = True
        End If
    Loop
        
    DisplayForm = mbIsOkPressed
    
    Unload Me
    
End Function

Private Function ValidateForm(ByRef rsMessage As String) As Boolean
    Dim bReturn As Boolean
    rsMessage = vbNullString
    bReturn = True
    
    If IsTextBoxEmpty(txtServerName) Then
        bReturn = False
        rsMessage = rsMessage & "Server name can not be blank" & vbCrLf
    End If
    
    If IsTextBoxEmpty(txtDBName) Then
        bReturn = False
        rsMessage = rsMessage & "Database name can not be blank" & vbCrLf
    End If
    
    If IsTextBoxEmpty(txtLogin) Then
        bReturn = False
        rsMessage = rsMessage & "Login name can not be blank" & vbCrLf
    End If
    
    ValidateForm = bReturn
    
End Function

Private Function IsTextBoxEmpty(ByVal voTextBox As TextBox) As Boolean
    If Trim$(voTextBox.Text) = vbNullString Then
        IsTextBoxEmpty = True
    Else
        IsTextBoxEmpty = False
    End If
End Function

Private Sub cmdTest_Click()
    
    On Error Resume Next
    
    Dim oADCConn As ADODB.Connection
    Dim cConnection As String
    Dim cDatabaseServerName As String
    Dim cDefaultDatabaseName As String
    Dim cUserId As String
    Dim cPassword As String

    cDatabaseServerName = txtServerName.Text
    cDefaultDatabaseName = txtDBName.Text
    cUserId = txtLogin.Text
    cPassword = txtPassword.Text
    
    'Build the connection string
    cConnection = MakeConnectString(cDatabaseServerName, cDefaultDatabaseName, cUserId, cPassword)

    Set oADCConn = New ADODB.Connection
    
    Me.MousePointer = vbHourglass
    
    oADCConn.Open cConnection
    
    Me.MousePointer = vbDefault
    
    If Err.Number = 0 Then
        MsgBox "Connection successfull!", , "Login Test"
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, , "Login Test"
    End If

End Sub

Private Sub txtDBName_GotFocus()
    Call SelectAllInTextBox(txtDBName)
End Sub

Private Sub txtLogin_GotFocus()
    Call SelectAllInTextBox(txtLogin)
End Sub

Private Sub txtPassword_GotFocus()
    Call SelectAllInTextBox(txtPassword)
End Sub

Private Sub txtServerName_GotFocus()
    Call SelectAllInTextBox(txtServerName)
End Sub

