VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MXLookup using StdIO by Amine Haddad"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstLog 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "Lookup"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   120
      MaxLength       =   256
      TabIndex        =   0
      Text            =   "hotmail.com"
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MxLookup As cMXLookupStdIO
Attribute MxLookup.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    If MxLookup.Ready = True Then
        Log "Not looking up anything at this time."
    Else
        MxLookup.Cancel
        Log "Canceling.."
    End If
End Sub

Private Sub cmdLookup_Click()
    If MxLookup.Ready = False Then
        Log "Please hold until it is done looking up current domain (" & MxLookup.Domain & ")"
    Else
        lstLog.Clear
        Log "Looking up MX Server's for: " & Trim(txtDomain.Text)
        Log "using " & MxLookup.Version
        If MxLookup.MxLookup(Trim(txtDomain.Text)) = False Then
            Log "MxLookup failed (or canceled)."
        End If
    End If
End Sub

Private Sub Form_Load()
    Set MxLookup = New cMXLookupStdIO
    Log "Successfully loaded " & MxLookup.Version
End Sub

Private Sub Log(ByVal strData As String)
    lstLog.AddItem strData
    lstLog.ListIndex = lstLog.NewIndex
    lstLog.ListIndex = -1
End Sub

Private Sub MxLookup_Complete(ByVal iFound As Integer)
    Log "Complete! Found " & iFound & " MX Servers for " & MxLookup.Domain
End Sub

Private Sub MxLookup_GotMX(ByVal Preference As Integer, ByVal Host As String)
    Log "Found MX Server: " & Host & " (" & Preference & ")"
End Sub
