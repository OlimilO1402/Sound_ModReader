VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 140
Private mBytLen As Long 'die Länge des Buffers
Private mSndBuff() As Byte
Private mWave   As WaveSound

Private Sub Class_Initialize()
  ReDim mSndBuff(0)
  mSndBuff(0) = &H7F
End Sub

Public Sub Read(ByVal aFNr As Integer, Optional ByVal aFileOffset As Long = -1)
TryE: On Error GoTo CatchE
  If mBytLen > 0 Then
    If aFileOffset >= 0 Then
      Get aFNr, aFileOffset, mSndBuff
    Else
      Get aFNr, , mSndBuff
    End If
  End If
  Exit Sub
CatchE:
  Call ErrHandler("Read")
End Sub
Public Sub WWrite(ByVal aFNr As Integer, Optional ByVal aFileOff As Long = -1)
'ist FileOffset = 1 , dann die ganze Datei schreiben
'ohne FileOffset, oder fileOffset <> 1 nur den Datenbereich schreiben
TryE: On Error GoTo CatchE
  If mBytLen Then
    If aFileOff >= 0 Then
      Put aFNr, aFileOff, mSndBuff
    Else
      Put aFNr, , mSndBuff
    End If
  End If
  Exit Sub
CatchE:
  Call ErrHandler("Write")
End Sub
'Public Sub LoadFromFile(aFileName As String)
'  Dim FNr As Integer: FNr = FreeFile
'  Open aFileName For Binary Access Read As FNr
'  Call SetBufferLength(LOF(FNr) - HL)
'  Call Read(FNr, 1)
'  Close FNr
'End Sub
'Public Sub SaveToFile(aFileName As String)
'TryE: On Error GoTo CatchE
'  Dim FNr As Integer: FNr = FreeFile
'  Open aFileName For Binary Access Write As FNr
'  Call WWrite(FNr, 1)
'  Close FNr
'  Exit Sub
'CatchE:
'  MsgBox "Fehler beim Schreiben der Datei: " & aFileName
'End Sub
Public Sub SetBufferLength(ByVal BytLen As Long)
  mBytLen = BytLen
  If mBytLen > 0 Then ReDim Preserve mSndBuff(0 To BytLen - 1)
End Sub

Public Function ToString(Optional linelen As Long) As String
Dim i As Long, s As String
  For i = 0 To UBound(mSndBuff)
    s = s & PadRight(PadLeft(Hex$(mSndBuff(i)), 2, "0"), 3)
    If Len(s) >= linelen Then
      ToString = ToString & s & vbCrLf
      s = vbNullString
    End If
  Next
End Function

Public Function DrawToPictureBox(aPB As PictureBox, ByVal bolPoints As Boolean)
Dim x As Long, stepX As Double
Dim y As Long, stepY As Double
Dim i As Long, iy As Integer 'signed integer
Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelY
  x = 0
  y = aPB.ScaleHeight \ 2
  stepX = (aPB.ScaleWidth / (UBound(mSndBuff) + 1))
  stepY = (aPB.ScaleHeight - brdr) / 256
  aPB.CurrentX = x
  aPB.CurrentY = y
  For i = 0 To UBound(mSndBuff)
    iy = UByteToSInt16(mSndBuff(i))
    x = i * stepX
    y = stepY * iy + aPB.ScaleHeight \ 2
    'Linien oder Punkte zeichnen?
    If bolPoints Then
      aPB.PSet (x, y)
    Else
      aPB.Line -(x, y)
    End If
  Next
End Function

Public Function ConvertToWaveSound() As WaveSound
TryE: On Error GoTo FinallyE
  Dim ws As WaveSound: Set ws = New_WaveSound(8, 1, 14000)
  If mBytLen Then
    Dim i As Long, bl As Long: bl = UBound(mSndBuff) + 1
    Call ws.SetBufferLength(bl)
    Dim bb As TByteBuffer: Call InitByteBuffer(bb, ws.pSndData, bl)
    If UBound(bb.p) = UBound(mSndBuff) Then
      For i = 0 To UBound(mSndBuff)
        bb.p(i) = ByteBE2LE(mSndBuff(i))
      Next
    Else
      Call Err.Raise(9)
    End If
  End If
  Set ConvertToWaveSound = ws
FinallyE:
  Call DeleteByteBuffer(bb)
  If Err = 0 Then Exit Function
CatchE:
  Call ErrHandler("ConvertToWaveSound")
End Function

'##############################'   Public Properties   '##############################'
'Public Property Get pSndBuff() As Long
'  pSndBuff = VarPtr(mSndBuff(0))
'End Property

Public Property Get Wave() As WaveSound
  If mWave Is Nothing Then _
    Set mWave = ConvertToWaveSound
  Set Wave = mWave
End Property

'##############################'  Private Procedures  '##############################'
Private Function ByteBE2LE(aByte As Byte) As Byte
  If aByte <= &H7F Then ByteBE2LE = aByte + &H7F Else ByteBE2LE = aByte - &H7F
End Function

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("ModSound", aProcName, Addinfo)
End Sub

