Attribute VB_Name = "ModException"
Option Explicit '2007_05_17 Zeilen: 96

Public Sub GlobalErrHandler(KlassName As String, ProcName As String, Optional Addinfo As String)
Dim mess As String
  mess = "In der Klasse: " & KlassName & _
    " ist in der Prozedur: " & ProcName & vbCrLf & _
    "ein Fehler mit der Nummer: " & CStr(Err.Number) & _
    "aufgetreten, mit der Meldung: " & vbCrLf & _
    Err.Description '& vbCrLf & '_
    If Len(Addinfo) Then _
      mess = mess & vbCrLf & "ZusatzInfo: " & vbCrLf & Addinfo
  MsgBox mess, vbCritical
  If Not IsInIDE Then
    'so jetzt noch an ne Log-Datei anh‰ngen
    Call AppendToLogFile(mess)
  End If
End Sub

Private Function IsInIDE() As Boolean
TryE: On Error GoTo CatchE
  Debug.Print 1 / 0
  Exit Function
CatchE:
  IsInIDE = True
End Function

Private Sub AppendToLogFile(errmess As String)
TryE: On Error GoTo FinallyE
  Dim FNm As String: FNm = App.Path & "\" & App.EXEName & "_ErrorLog.txt"
  Dim FNr As Integer: FNr = FreeFile
  Open FNm For Append As FNr
  Dim sd As String: sd = Format$(Now, "dd.mm.yyyy hh:mm:ss")
  Print #FNr, sd
  Print #FNr, errmess
  'und ne Meldung ausgeben, daﬂ ein logfile existiert
  MsgBox "watchout for logfile: " & vbCrLf & FNm
FinallyE:
  Close FNr
End Sub

