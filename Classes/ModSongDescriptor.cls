VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModSongDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 182; 05_21: 201
Private Type TModSongDescriptor
  NrOfPatternsInUse As Byte
  SongEndJmpPos     As Byte
  PatSeqTab(0 To 127)   As Byte 'PatternSequenceTable
  FileFormatTag(0 To 3) As Byte
End Type 'Sum: 134Bytes
Private mHasNoTag As Boolean '=True wenn kein Fileformattag
Private mSD As TModSongDescriptor
Private mTrackerProg As String 'which TrackProg is detected by the FileFormatTag

Public Sub Read(ByVal aFNr As Long, ByVal filepos As Long)
Dim oldFPos As Long
TryE: On Error GoTo CatchE
  If filepos Then
    oldFPos = filepos
  Else
    oldFPos = Seek(aFNr)
  End If
  Get aFNr, filepos, mSD
  'jetzt FileFormatTag überprüfen
  Dim sChan As String: sChan = GetNrOfChannelsToString
  If Len(sChan) = 0 Then
    'die Position um 4 byte zurücksetzen, jetzt sollten Mod-Dateien
    'auch ohne FileFormatTag richtig eingelesen werden
    'mit einer Reader oder FileStream Klasse hier einfach
    'das Property Position setzen.
    'oder von der alten Position weg, nochmal einlesen, diesmal
    'ohne FileFormatTag
    Get aFNr, oldFPos, mSD.NrOfPatternsInUse
    Get aFNr, , mSD.SongEndJmpPos
    Get aFNr, , mSD.PatSeqTab
    mHasNoTag = True
  End If
  Exit Sub
CatchE:
  Call ErrHandler("Read")
End Sub

Public Function ToString() As String
  With mSD
    ToString = "FileFormatTag     : " & FileFormatTag & vbCrLf & _
               "TrackerProg       : " & mTrackerProg & vbCrLf & _
               "NrOfPatternsInUse : " & CStr(.NrOfPatternsInUse) & vbCrLf & _
               "SongEndJumpPos    : " & CStr(.SongEndJmpPos) & vbCrLf & _
               "PatternSequence   : " & vbCrLf & PatternSeqToString(3) '& vbCrLf '& '_
  End With
End Function

'##############################'   Public Properties   '##############################'
Public Property Get TrackerProg() As String
  TrackerProg = mTrackerProg
End Property
Public Property Let TrackerProg(StrVal As String)
  mTrackerProg = StrVal
End Property

Public Property Get SongEndJumpPos() As Byte
  SongEndJumpPos = mSD.SongEndJmpPos
End Property
Public Property Let SongEndJumpPos(ByVal aPos As Byte)
  mSD.SongEndJmpPos = aPos
End Property

Public Property Get CountPatternsInUse() As Byte
  CountPatternsInUse = mSD.NrOfPatternsInUse
End Property
Public Property Let CountPatternsInUse(ByVal BytVal As Byte)
  mSD.NrOfPatternsInUse = BytVal
End Property

Public Property Get PatternSequence(ByVal Index As Byte) As Byte
  PatternSequence = mSD.PatSeqTab(Index)
End Property
Public Property Let PatternSequence(ByVal Index As Byte, ByVal PatternNr As Byte)
  mSD.PatSeqTab(Index) = PatternNr
End Property

Public Property Get FileFormatTag() As String
  FileFormatTag = AS2US(mSD.FileFormatTag)
'hmm hier könnte es evtl nützlich sein, nur die
'nonascii-zeichen durch space zu ersetzen
  If Len(FileFormatTag) = 0 Then
    Dim i As Long
    For i = 0 To 3
      If &H20 <= mSD.FileFormatTag(i) And mSD.FileFormatTag(i) <= 127 Then
        FileFormatTag = FileFormatTag & Chr$(mSD.FileFormatTag(i)) '& " "
      Else
        FileFormatTag = FileFormatTag & "&H" & PadLeft(Hex$(mSD.FileFormatTag(i)), 2, "0") & " "
      End If
    Next
    If mHasNoTag Then _
      FileFormatTag = FileFormatTag & " * No Tag * (-4) "
  End If
End Property
Public Property Let FileFormatTag(StrVal As String)
  AS2US(mSD.FileFormatTag) = StrVal
End Property

'##############################'   Friend Procedures  '##############################'
Friend Property Get ByteLength() As Long
  ByteLength = LenB(mSD)
End Property
Friend Function PatternSeqMaxNr() As Byte
  Dim i As Long: For i = 0 To 127
    PatternSeqMaxNr = MaxB(PatternSeqMaxNr, mSD.PatSeqTab(i))
  Next
End Function

Private Function GetNrOfChannelsToString() As String
'T'ja da muß man jetzt das FileFormatTag parsen
'M.K.', 'FLT4',
'M!K!', '4CHN' : 4 channels, 31 instruments
'6CHN'         : 6 channels, 31 instruments
'8CHN', 'OCTA' : 8 channels, 31 instruments
'FLT4', 'FLT8': Startrekker 4/8 channel file. ('FLT6' doesn't exists)
'CD81'        : Falcon 8 channel MODs
'2CHN'        : FastTracker 2 Channel MODs
'yyCH' where yy can be 10, 12, .. 30, 32: FastTracker yy Channel MODs
'yyCH' where yy can be 11, 13, 15: TakeTracker 11, 13, 15 channel MODs
'TDZx' where x can be 1, 2 or 3: TakeTracker 1, 2, 3 channel MODs
'xCHN' where x can be 5, 7 or 9: TakeTracker 5, 7, 9 channel MODs
'wir schauen nach dem ersten Byte, dann nach dem zweiten Byte...
Dim sChan As String
  With mSD
    Select Case .FileFormatTag(0)
    Case 77 '"M" ==> 'M.K.', 'M!K!'
      sChan = CStr(4)
      'unterscheidet bei ProTracker das 15 vom 31Samples-Format
      mTrackerProg = "AMIGA PROTracker (Mahoney & Kaktus)"
    Case 70 '"F" ==> 'FLT4', 'FLT8'
      If .FileFormatTag(1) = 76 And .FileFormatTag(2) = 84 Then
        sChan = Chr$(.FileFormatTag(3))
        mTrackerProg = "Startrekker"
      End If
    Case 67 '"C" ==> 'CD81'
      If .FileFormatTag(1) = 68 Then
        sChan = Chr$(.FileFormatTag(2))
        mTrackerProg = "Falcon"
      End If
    Case 84 '"T" ==> 'TDZx
      If .FileFormatTag(1) = 68 And .FileFormatTag(2) = 90 Then
        sChan = Chr$(.FileFormatTag(3))
        mTrackerProg = "TakeTracker"
      End If
    Case 79 '"O" ==> 'OCTA'
      If .FileFormatTag(1) = 67 And .FileFormatTag(2) = 84 And .FileFormatTag(3) = 65 Then
        sChan = CStr(8)
        mTrackerProg = "OCTA (unknown prog)"
      End If
    Case Else
      ' ==> 'xCHN'
      If .FileFormatTag(1) = 67 And .FileFormatTag(2) = 72 And .FileFormatTag(3) = 78 Then
        sChan = Chr$(.FileFormatTag(0))
        If CLng(sChan) = 2 Then
          mTrackerProg = "FastTracker"
        Else '5, 7 oder 9
          mTrackerProg = "TakeTracker"
        End If
      Else
        ' ==> 'xxCH'
        If .FileFormatTag(2) = 67 And .FileFormatTag(3) = 72 Then
          sChan = Chr$(.FileFormatTag(0)) & Chr$(.FileFormatTag(1))
          Select Case CLng(sChan)
          Case 11, 13, 15
            mTrackerProg = "TakeTracker"
          Case Else '2 oder 10, 12, ... , 30, 32
            mTrackerProg = "FastTracker"
          End Select
        End If
        If .FileFormatTag(1) = 120 Then
          sChan = CStr(4)
          mTrackerProg = "Protracker (early)"
        End If
      End If
    End Select
  End With
  GetNrOfChannelsToString = sChan
  'mTrackerProg = sFFTg & " " & mTrackerProg
End Function
Friend Function ComputeNrOfChannels() As Long
Dim sChan As String: sChan = GetNrOfChannelsToString
  If Len(sChan) Then ComputeNrOfChannels = CLng(sChan) Else ComputeNrOfChannels = 4
End Function

'##############################'  Private Procedures  '##############################'
Private Function PatternSeqToString(ind As Long) As String
Dim i As Long, s As String
Dim ppl As Long 'MaxPatternNrPerLine
  ppl = 16
  PatternSeqToString = String$(ind, " ")
  With mSD
    For i = LBound(.PatSeqTab) To UBound(.PatSeqTab)
      s = PadLeft(CStr(.PatSeqTab(i)), 4)
      If i < UBound(.PatSeqTab) Then s = s & " "
      If (i + 1) Mod ppl = 0 Then s = s & vbCrLf & String$(ind, " ")
      PatternSeqToString = PatternSeqToString & s
    Next
  End With
End Function

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("ModSongDescriptor", aProcName, Addinfo)
End Sub

