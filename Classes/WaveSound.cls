VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 270
'Klasse stellt einen einzelnen Sound dar
Public Enum WaveFormatTag
  PCM = &H1&     'Microsoft Pulse Code Modulation (PCM)
  MULAW = &H101& 'IBM mu-law format
  ALAW = &H102&  'IBM a-law format
  ADPCM = &H103& 'IBM AVC Adaptive Differential PCM format
End Enum
'der TRIFFWAVEfmt paßt zwar nicht immer, aber meistens jedenfalls so 90% der Fälle
Private Type TRIFFWAVEfmt
  BytsRIFF    As Long '&H46464952 = 1179011410 == Das Wort "RIFF"
  BuffLen     As Long 'die Länge des gesamten ByteBuffers
  BytsWAVE    As Long '&H45564157 = 1163280727 == Das Wort "WAVE"
  Bytsfmt     As Long '&H20746D66 = 544501094 == Das Wort "fmt "
  WAVEfmtLen  As Long '&H10 =16 Länge des darauffolgenden WAVEfmt-Headers
  FormatTag         As Integer 'eigentlich WaveFormatTag PCM = 1 == immer 1
  Channels          As Integer '1 oder 2
  SamplesPerSecond  As Long    '= 11025, 22050, 44100 ...
  AvgBytesPerSecond As Long    '= SamplesPerSecond * BlockAlign
  BlockAlign        As Integer '= Channels * BitsPerSample / 8
  BitsPerSample     As Integer '8 oder 16
  Bytsdata    As Long '&H61746164 = 1635017060 == Das Wort "data"  'richtig
  DataLen     As Long '=BuffLen - 44 '44=&H2C
End Type
Private Type PRIFFWaveFmt
  pudt As TUDTPtr
  p() As TRIFFWAVEfmt
End Type
Private Type PBuffer
  pudt As TUDTPtr
  p() As Byte
End Type
Private pHead As PRIFFWaveFmt
Private pBuff As PBuffer
Private Const HL As Long = 44 'HeaderLength
Private mOrigSPS As Long
'Private Const CHalfTone As Double = 1.0594630943593
Private mSndBuff() As Byte

Private Sub Class_Initialize()
  ReDim mSndBuff(0 To HL - 1)
  Call InitPHead: Call InitRIFFWAVE
  pSndBuff = VarPtr(mSndBuff(0))
End Sub
Private Sub Class_Terminate()
  Call DeletePHead
  Call DeletePBuff
End Sub

Friend Sub New_(ByVal bps As Integer, ByVal nchannels As Integer, ByVal sps As Long)
  With pHead.p(0)
    .Channels = nchannels
    .SamplesPerSecond = sps
    .BitsPerSample = bps
    .BlockAlign = .Channels * .BitsPerSample / 8
    .AvgBytesPerSecond = .SamplesPerSecond * .BlockAlign
  End With
End Sub
Public Sub NewF(aFileName As String)
  Call LoadFromFile(aFileName)
'  Dim FNr As Integer: FNr = FreeFile
'  Open aFileName For Binary As FNr
'  Call Read(FNr)
'  Close FNr
End Sub
Public Function ToString(Optional ByVal linelen As Long = 40) As String
Dim i As Long, s As String
  For i = HL To UBound(mSndBuff)
    s = s & PadRight(PadLeft(Hex$(mSndBuff(i)), 2, "0"), 3)
    If Len(s) >= linelen Then
      ToString = ToString & s & vbCrLf
      s = vbNullString
    End If
  Next
End Function

Public Sub Read(ByVal aFNr As Integer, Optional ByVal aFileOff As Long = -1)
'Das wav-Format ist normalerweise nicht so fix wie hier, meistens genügt der hier
'zwar, aber es kann auch andere Möglichkeiten geben.
'Dann müßte man die Position des WaveFormats anhand "fmt ", und
'die Position der Daten anhand "data" suchen.
'
'nur überhaupt was lesen, wenn der Buffer ein Größe > HL hat
  Dim n As Long: n = UBound(mSndBuff) - HL + 1
  If n > 0 Then
    Call InitPBuff
    If aFileOff >= 0 Then
      Get aFNr, aFileOff, pBuff.p
    Else
      Get aFNr, , pBuff.p
    End If
    Call DeletePBuff
    pSndBuff = VarPtr(mSndBuff(0))
    Call InitLen
  End If
End Sub
Public Sub WWrite(ByVal aFNr As Integer, Optional ByVal aFileOff As Long = -1)
'ist FileOffset = 1 , dann die ganze Datei schreiben
'ohne FileOffset, oder fileOffset <> 1 nur den Datenbereich schreiben
  If aFileOff = 1 Then
    Put aFNr, , mSndBuff
  Else
    Call InitPBuff
    If aFileOff >= 0 Then
      Put aFNr, aFileOff, pBuff.p
    Else
      Put aFNr, , pBuff.p
    End If
    Call DeletePBuff
  End If
End Sub
Public Sub LoadFromFile(aFileName As String)
  Dim FNr As Integer: FNr = FreeFile
  Open aFileName For Binary Access Read As FNr
  Call SetBufferLength(LOF(FNr))
  Call Read(FNr, 1)
  Close FNr
  pSndBuff = VarPtr(mSndBuff(0))
End Sub
Public Sub SaveToFile(aFileName As String)
TryE: On Error GoTo CatchE
  Dim FNr As Integer: FNr = FreeFile
  Open aFileName For Binary Access Write As FNr
  Call WWrite(FNr, 1)
  Close FNr
  Exit Sub
CatchE:
  Call ErrHandler("SaveToFile", aFileName)
End Sub
Public Sub SetBufferLength(ByVal BytLen As Long)
  If BytLen > 0 Then
    ReDim Preserve mSndBuff(0 To HL + BytLen - 1)
    pSndBuff = VarPtr(mSndBuff(0))
    Call InitLen
  End If
End Sub
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
  For i = HL To UBound(mSndBuff)
    iy = mSndBuff(i) - &H7F
    x = (i - HL) * stepX
    y = stepY * iy + aPB.ScaleHeight \ 2
    'Linien oder Punkte zeichnen?
    If bolPoints Then
      aPB.PSet (x, y)
    Else
      aPB.Line -(x, y)
    End If
  Next
End Function

Private Property Let pSndBuff(ByVal LngVal As Long)
  pHead.pudt.pvData = LngVal
  mOrigSPS = pHead.p(0).SamplesPerSecond
End Property
'##############################'   Public Properties   '##############################'
Public Property Get pSndBuff() As Long
  pSndBuff = VarPtr(mSndBuff(0))
End Property
Public Property Get pSndData() As Long
  pSndData = VarPtr(mSndBuff(HL))
End Property

Public Property Get OrigSPS() As Long
  OrigSPS = mOrigSPS
End Property

'##############################'  v WaveFormat v  '##############################'
Public Property Get FormatTag() As WaveFormatTag
  FormatTag = pHead.p(0).FormatTag
End Property
Public Property Let FormatTag(wftgVal As WaveFormatTag)
  pHead.p(0).FormatTag = wftgVal
End Property

Public Property Get Channels() As Integer
  Channels = pHead.p(0).Channels
End Property
Public Property Let Channels(nchannels As Integer)
  With pHead.p(0)
    .Channels = nchannels
    .BlockAlign = nchannels * .BitsPerSample / 8
    .AvgBytesPerSecond = .SamplesPerSecond * .BlockAlign
  End With
End Property

Public Property Get SamplesPerSecond() As Long
  SamplesPerSecond = pHead.p(0).SamplesPerSecond
End Property
Public Property Let SamplesPerSecond(LngVal As Long)
  With pHead.p(0)
    .SamplesPerSecond = LngVal
    .AvgBytesPerSecond = LngVal * .BlockAlign
  End With
End Property

Public Property Get AvgBytesPerSecond() As Long
  AvgBytesPerSecond = pHead.p(0).AvgBytesPerSecond
End Property
Public Property Get BlockAlign() As Integer
  BlockAlign = pHead.p(0).BlockAlign
End Property
  
Public Property Get BitsPerSample() As Integer
  BitsPerSample = pHead.p(0).BitsPerSample
End Property
Public Property Let BitsPerSample(IntVal As Integer)
  With pHead.p(0)
    .BitsPerSample = IntVal
    .BlockAlign = .Channels * .BitsPerSample / 8
    .AvgBytesPerSecond = .SamplesPerSecond * .BlockAlign
  End With
End Property
'##############################'  ^ WaveFormat ^  '##############################'

Public Property Let KeyNote(aLngVal As Long)
  If aLngVal Then
    SamplesPerSecond = CLng(mOrigSPS * 2 ^ (aLngVal / 12))
  Else
    SamplesPerSecond = mOrigSPS
  End If
End Property

'##############################'  Private Procedures  '##############################'
Private Sub InitPHead()
  With pHead.pudt: .pSA = VarPtr(.cDims): .cDims = 1
    .fFeatures = 32: .cbElements = HL: .cElements = 1
    .pvData = VarPtr(mSndBuff(0))
  End With
  Call RtlMoveMemory(ByVal ArrPtr(pHead.p), ByVal VarPtr(pHead), 4)
End Sub
Private Sub InitRIFFWAVE()
  With pHead.p(0)
    .BytsRIFF = &H46464952
    .BytsWAVE = &H45564157
    .Bytsfmt = &H20746D66
    .WAVEfmtLen = &H10
    .FormatTag = WaveFormatTag.PCM
    .Bytsdata = &H61746164
  End With
End Sub
Private Sub InitLen()
  With pHead.p(0)
    .BuffLen = UBound(mSndBuff) + 1 - 8
    .DataLen = UBound(mSndBuff) - HL + 1
  End With
End Sub
Private Sub InitPBuff()
  With pBuff.pudt: .pSA = VarPtr(.cDims): .cDims = 1
    .cbElements = 1: .cElements = UBound(mSndBuff) - HL + 1
    .pvData = VarPtr(mSndBuff(HL))
  End With
  Call RtlMoveMemory(ByVal ArrPtr(pBuff.p), ByVal VarPtr(pBuff), 4)
End Sub
Private Sub DeletePHead()
  Call RtlZeroMemory(ByVal ArrPtr(pHead.p), 4)
End Sub
Private Sub DeletePBuff()
  Call RtlZeroMemory(ByVal ArrPtr(pBuff.p), 4)
End Sub

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("WaveSound", aProcName, Addinfo)
End Sub
