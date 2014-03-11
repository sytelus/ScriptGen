VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Progress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3450
      TabIndex        =   0
      Top             =   570
      Width           =   1155
   End
   Begin VB.Label lblProgresstext 
      AutoSize        =   -1  'True
      Caption         =   "Progress"
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moParentForm As Form

Public Sub DisplayProgress(ByVal vsTitle As String, ByVal vsMessage As String, ByVal voParentForm As Form)
    Me.Caption = vsTitle
    Me.lblProgresstext.Caption = vsMessage
    Set moParentForm = voParentForm
    Me.Show vbModeless, voParentForm
End Sub

Public Function HideProgress()
    Set moParentForm = Nothing
    Me.Hide
End Function

Private Sub cmdCancel_Click()
    On Error Resume Next
    Call moParentForm.OnProgressFormCanceled
    If Err.Number = 438 Then    'Object doesn't support this property or method
        'Ignore it
    Else
        ReRaiseError
    End If
    On Error GoTo Err_cmdCancel_Click
    Call HideProgress
Exit Sub
Err_cmdCancel_Click:
    ReRaiseError
End Sub
