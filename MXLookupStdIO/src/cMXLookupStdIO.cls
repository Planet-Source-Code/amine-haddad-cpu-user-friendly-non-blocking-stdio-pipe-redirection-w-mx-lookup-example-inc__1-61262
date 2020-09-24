VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMXLookupStdIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The following are just global constant for the class itself
Const c_Version = "MXLookup Class v0.1 BETA by Amine Haddad"

'The following are events to be launched throughout the class
Public Event GotMX(ByVal Preference As Integer, ByVal Host As String)
Public Event Complete(ByVal iFound As Integer)

'The following are variable settings
Dim mDomain As String
Dim mFound As Integer

'The following are the classes needed throughout the program
Dim WithEvents StdIO As cStdIO
Attribute StdIO.VB_VarHelpID = -1

'The following are the properties used in this class
Public Property Let Domain(ByVal sDomain As String)
    mDomain = sDomain
End Property

Public Property Get Domain() As String
    Domain = mDomain
End Property

Public Property Get Ready() As Boolean
    Ready = StdIO.Ready
End Property

Public Property Get Version() As String
    Version = c_Version
End Property

'The following is the class initialization sub
Private Sub Class_Initialize()
    Set StdIO = New cStdIO
End Sub

'The following are functions used in this class
Public Function MxLookup(Optional ByVal sDomain As String) As Boolean
    Dim sOutput As String
    If StdIO.Ready = False Then
        Exit Function
    Else
        If Len(sDomain) > 0 Then
            mDomain = sDomain
        End If
        
        If Len(mDomain) = 0 Then
            Exit Function
        End If
        
        mFound = 0
        sOutput = StdIO.ExecuteCommand(Environ("ComSpec") & " /c nslookup -timeout=2 -type=mx " & Trim(mDomain))
        'now we got all the output in sOutput.. let's get the mxServers
        If InStr(1, sOutput, "mail exchanger =", vbTextCompare) = 0 Then
            Exit Function
        Else
            Dim arrLines() As String
            Dim i As Integer
            arrLines() = Split(sOutput, vbCrLf)
            For i = 0 To UBound(arrLines)
                If InStr(1, arrLines(i), "mail exchanger =", vbTextCompare) <> 0 Then
                    mFound = mFound + 1
                    RaiseEvent GotMX(Split(Split(arrLines(i), "MX Preference = ", , vbTextCompare)(1), ",", , vbTextCompare)(0), Split(arrLines(i), "mail exchanger =", , vbTextCompare)(1))
                End If
            Next i
            RaiseEvent Complete(mFound)
            If mFound > 0 Then
                MxLookup = True
            End If
        End If
    End If
End Function

Public Sub Cancel()
    StdIO.Cancel
End Sub
