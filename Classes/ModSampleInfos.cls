VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModSampleInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 129
Private Type SampleInfo
  SampleName(0 To 21) As Byte
  'Amiga==HHLL (BigEndWord) <==>Intel==LLHH (LowEndWord)
  SampleLen(0 To 1)   As Byte
  Finetune            As Byte 'SNibble == signed nibble (first 4bit)
  LinPlayBVol         As Byte 'Byte 'PlayBack-Valume LINEAR
  RepeatStartOff(0 To 1) As Byte 'Big-End-Word
  RepeatLen(0 To 1)   As Byte 'Big-End-Word
End Type 'Sum: 30Bytes
Private mBank() As SampleInfo

Public Sub SetLength(ByVal nCount As Long)
  ReDim mBank(0 To nCount - 1)
End Sub

Public Sub Read(ByVal aFNr As Long, ByVal aFileOffset As Long)
TryE: On Error GoTo CatchE
  Get aFNr, aFileOffset, mBank
  Exit Sub
CatchE:
  Call ErrHandler("Read")
End Sub

Public Function CountUsedSamples() As Long
Dim i As Long, c As Long ': c = 1
  For i = 0 To UBound(mBank) '- 1
    If BE2LE(mBank(i).SampleLen) > 0 Then c = c + 1
  Next
  CountUsedSamples = c
End Function

Public Function ToString() As String
Dim i As Long, s As String
  ToString = "SmpleInfos:" & vbCrLf & _
  " Nr SampleName            Length FineTune LinPlayBVol RepeatStartOff RepeatLen" & vbCrLf
  For i = LBound(mBank) To UBound(mBank)
    With mBank(i)
      s = PadLeft(PadLeft(CStr(i + 1), 2), 3) & " " & _
          PadRight(AS2US(.SampleName), 20) & " " & _
          PadLeft(CStr(BE2LE(.SampleLen)), 7) & " " & _
          PadLeft(CStr(.Finetune), 8) & " " & _
          PadLeft(CStr(.LinPlayBVol), 11) & " " & _
          PadLeft(CStr(BE2LE(.RepeatStartOff)), 14) & " " & _
          PadLeft(CStr(BE2LE(.RepeatLen)), 9) & vbCrLf
      ToString = ToString & s
      s = vbNullString
    End With
  Next
End Function

'##############################'   Public Properties   '##############################'
Public Property Get MaxSamples() As Long
  MaxSamples = UBound(mBank) + 1
End Property

Public Property Get SampleName(Index As Long) As String
  SampleName = AS2US(mBank(Index).SampleName)
End Property
Public Property Let SampleName(Index As Long, StrVal As String)
  AS2US(mBank(Index).SampleName) = StrVal
End Property

Public Property Get SampleLen(Index As Long) As Long
  SampleLen = BE2LE(mBank(Index).SampleLen)
End Property
Public Property Let SampleLen(Index As Long, LngVal As Long)
  BE2LE(mBank(Index).SampleLen) = LngVal
End Property

Public Property Get AllSampleLen() As Long
  Dim i As Long: For i = 0 To MaxSamples - 1
    AllSampleLen = AllSampleLen + SampleLen(i) * 2
  Next
End Function

Public Property Get Finetune(Index As Long) As Byte
  Finetune = mBank(Index).Finetune
End Property 'SNibble == signed nibble (first 4bit)
Public Property Let Finetune(Index As Long, BytVal As Byte)
  mBank(Index).Finetune = BytVal
End Property

Public Property Get LinearPlayBackVolume(Index As Long) As Byte
  LinearPlayBackVolume = mBank(Index).LinPlayBVol
End Property 'Byte 'PlayBack-Volume LINEAR
Public Property Let LinearPlayBackVolume(Index As Long, BytVal As Byte)
  mBank(Index).LinPlayBVol = BytVal
End Property

Public Property Get RepeatStartOffset(Index As Long) As Long
  RepeatStartOffset = BE2LE(mBank(Index).RepeatStartOff)
End Property
Public Property Let RepeatStartOff(Index As Long, LngVal As Long)
  BE2LE(mBank(Index).RepeatStartOff) = LngVal
End Property

Public Property Get RepeatLength(Index As Long) As Long
  RepeatLength = BE2LE(mBank(Index).RepeatLen)
End Property
Public Property Let RepeatLength(Index As Long, LngVal As Long)
  BE2LE(mBank(Index).RepeatLen) = LngVal
End Property

Public Property Get ByteLength() As Long
Dim n As Long
  n = UBound(mBank) - LBound(mBank) + 1
  ByteLength = n * LenB(mBank(0))
End Property

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("ModSampleInfos", aProcName, Addinfo)
End Sub

