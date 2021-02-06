VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 238
'Many thanks to:
'  'kurt.kennett@gravis.com' or 'kurtt@sfu.ca'
'Modfil10.txt by THUNDER (kurtt@sfu.ca) for creating the file, and
'Modfil11.txt by Erland Van Olmen (erlandvo@hotmail.com) for editing the file
'and by www.wotsit.org, for providing the files
'
Private mFileName       As String
Private mFileLenBytes   As Long 'die Länge der Datei in Bytes
Private mSongTitle(0 To 19) As Byte 'String * 20
Private mSampleInfoBank As ModSampleInfos
Private mSongDesc       As ModSongDescriptor
Private mMaxSamples     As Long
Private mMaxChannels    As Long   'nicht stereo links-rechts sondern Anzahl der Spuren
Private mSizeOfOnePattern As Long
Private mMaxPatterns    As Long
Private mPatterns       As Collection 'speichert ModPattern-Objekte
Private mSounds         As Collection 'speichert ModSound-Objekte

Public Sub ReadFile(aFileName As String)
TryE: On Error GoTo FinallyE
  Dim i As Long, oldseeker As Long, FNr As Long
  mFileName = aFileName
  FNr = FreeFile
  Open mFileName For Binary As FNr 'Access Read
  mFileLenBytes = LOF(FNr)

  If IsCompressed(FNr) Then Exit Sub

  ' 1 ...  20 SongTitle
  Get FNr, 1, mSongTitle

  oldseeker = Seek(FNr) 'die aktuelle Dateiposition sichern
  mMaxSamples = ComputeNrOfSamples(FNr)
  '21 ... 950 SampleInfoBank
  Set mSampleInfoBank = New_ModSampleInfos(FNr, mMaxSamples, oldseeker)
  'ReDim mSampleInfoBank(0 To mNrOfSamples - 1)
  'Get mFNr, oldseeker, mSampleInfoBank 'die aktuelle Dateiposition wiederherstellen

  '951 - 1084 SongDescriptor
  Set mSongDesc = New_ModSongDescriptorF(FNr, Seek(FNr))

  '1085 ... 'Pattern and Sample Data
  mMaxChannels = mSongDesc.ComputeNrOfChannels
  mSizeOfOnePattern = GetSizeOfOnePattern(mMaxChannels)
  mMaxPatterns = ComputeNrOfPatterns(mFileLenBytes)
  Call ClearPatterns
  Dim p As ModPattern
  For i = 0 To mMaxPatterns '- 1 'nein nicht -1 !
    Set p = New_ModPattern(Me, i)
    Call p.Read(FNr)
    Call mPatterns.Add(p, CStr(i))
  Next
  'so jetzt noch die Samples lesen
  Call ClearSounds
  Dim ms As ModSound
  Dim sl As Long
  'Dim sn As String
  'nö es kann vorkommen, daß zwei Samples den gleichen Namen haben
  For i = 0 To mSampleInfoBank.MaxSamples - 1
    sl = mSampleInfoBank.SampleLen(i) * 2
    'If sl > 0 Then
      Set ms = New_ModSound(FNr, sl)
      'sn = mSampleInfoBank.SampleName(i)
      Call mSounds.Add(ms, CStr(i + 1)) ', sn)
    'End If
  Next
  Call CheckMessFileLenSeek(FNr)
FinallyE:
  Close FNr
  If Err = 0 Then Exit Sub
CatchE:
  Call ErrHandler("ReadFile", aFileName)
End Sub
Public Sub ClearPatterns()
  Set mPatterns = New Collection
End Sub
Public Property Get Patterns() As Collection
  Set Patterns = mPatterns
End Property
Public Property Get Pattern(ByVal aVarKey As Variant) As ModPattern
  Set Pattern = mPatterns.Item(aVarKey)
End Property

Public Property Get SampleInfoBank() As ModSampleInfos
  Set SampleInfoBank = mSampleInfoBank
End Property

Public Sub ClearSounds()
  Set mSounds = New Collection
End Sub
Public Property Get Sounds() As Collection
  Set Sounds = mSounds
End Property
Public Property Get Sound(ByVal aVarKey As Variant) As ModSound
  Set Sound = mSounds.Item(aVarKey)
End Property

Public Function ToString() As String
  ToString = _
    "SongTitle         : " & SongTitle & vbCrLf & _
    "MaxChannels       : " & CStr(mMaxChannels) & vbCrLf & _
    "Samples Used/Max  : " & CStr(mSampleInfoBank.CountUsedSamples) & "/" & CStr(mMaxSamples) & vbCrLf & _
    "PatternSize       : " & CStr(mSizeOfOnePattern) & vbCrLf & _
    mSongDesc.ToString & vbCrLf & _
    mSampleInfoBank.ToString & vbCrLf
End Function

'##############################'   Public Properties   '##############################'
Public Property Get SongTitle() As String
  SongTitle = AS2US(mSongTitle)
End Property
Public Property Let SongTitle(StrVal As String)
  'Str2BA(mSongTitle) = StrVal
  AS2US(mSongTitle) = StrVal
End Property

Public Property Get MaxChannels() As Long
  MaxChannels = mMaxChannels
End Property
Public Property Let MaxChannels(LngVal As Long)
  mMaxChannels = LngVal
End Property

Public Property Get MaxSamples() As Long
  MaxSamples = mMaxSamples
End Property
Public Property Let MaxSamples(LngVal As Long)
  mMaxSamples = LngVal
End Property

Public Property Get PatternSize() As Long
  'GetSizeOfOnePattern = nChannels * 4 * 64
  'PatternSize = nChannels * 4 * 64
  PatternSize = mSizeOfOnePattern
End Property
'Public Property Let PatternSize(LngVal As Long)
'  mPatternSize = LngVal
'End Property

'##############################'   Private Procedures   '##############################'
Private Function IsCompressed(ByVal aFNr As Long) As Boolean
  Dim B(0 To 4) As Byte
  Get aFNr, , B
  'Asc("P") = 80
  'PACK
  'oder
  'PP20
  Dim p As String: p = AS2US(B)
  If (B(0) = 80 And B(1) = 80 And B(2) = 50 And B(3) = 48) Or _
     (B(0) = 80 And B(1) = 65 And B(2) = 67 And B(3) = 75) Then
    Dim mess As String
    mess = "File is probably compressed using a packer like PowerPacker: " & p & vbCrLf & _
           "Use a decompressing utility first, like pp20unp.exe!" & vbCrLf & _
           "Maybe you can get it there:" & vbCrLf & _
           "www.Hornet.org/music/programs/players/tnt-mpXX.zip" & vbCrLf & _
           "or search for the fileformat on www.wotsit.org" & vbCrLf & _
           "Try to preceed anyway?"
    Dim mr As VbMsgBoxResult
    mr = MsgBox(mess, vbQuestion Or vbYesNo, "Compressed File Preceed Anyway?")
    If mr = vbNo Then
      IsCompressed = True
    End If
  End If
End Function
Private Function ComputeNrOfSamples(aFNr As Long) As Long
Dim B As Byte
'As for how many instruments there are, check the bytes at location 471 in
'the file. If there is text  there (ASCII $20-$7E (32-126)),  then you can
'probably assume it 's a 31-instrument file.  Otherwise, it's an older 15
'instrument file
'LAMESPEC_OM: no, either it must be the other way, or it's simply not true
  '&D472 == &H01D8
  Get aFNr, &H1D8, B 'in VB it must be 471 + 1 = 472
  'If (32 <= B) And (B <= 126) Then
  If B = &H78 Then
    'in 4-15 Format mods here often is "x" == &H78, &D120
    ComputeNrOfSamples = 15
    'ComputeNrOfSamples = 31
  Else
    'ComputeNrOfSamples = 15
    ComputeNrOfSamples = 31
  End If
End Function
Private Function GetSizeOfOnePattern(nchannels As Long) As Long
  'Nr_Of_Channels*4*64 '4= Länge in Byte einer Note, 64= Anzahl an Zeilen in einem Pattern
  GetSizeOfOnePattern = nchannels * 4 * 64
End Function
Private Function ComputeNrOfPatterns(FileLength As Long) As Long
  'this is again a rather unsafe-method
  'there are two totally different methods described in the paper
  'but imho both are very unsafe
  Dim n1 As Long, n2 As Long
  'n1 = GetNOPFileLen(FileLength)
  n2 = mSongDesc.PatternSeqMaxNr
  If n1 > 128 Then n1 = 0
  If n2 > 128 Then n2 = 0
  'If n1 = n2 Then
  '  'MsgBox "Everything is OK"
  'End If
  ComputeNrOfPatterns = MaxL(n1, n2)
End Function
Private Function GetNOPFileLen(FLen As Long) As Long
Dim HdLen As Long: HdLen = GetHeaderLen
Dim OSLen As Long: OSLen = mSampleInfoBank.AllSampleLen
  If mSizeOfOnePattern <> 0 Then _
    GetNOPFileLen = ((mFileLenBytes - (HdLen + OSLen)) / mSizeOfOnePattern)
End Function
Private Function GetHeaderLen() As Long
  GetHeaderLen = UBound(mSongTitle) + 1 + mSampleInfoBank.ByteLength + mSongDesc.ByteLength
End Function
Private Sub CheckMessFileLenSeek(ByVal aFNr As Long)
  'wenn sich ein Unterschied von 4 ergibt, dann müßte man das
  'Ganze nochmal einlesen.
  Dim sk As Long: sk = Seek(aFNr) - 1
  Dim df As Long: df = sk - mFileLenBytes
  If df Then
    Dim mess As String
    mess = "Dateilänge: " & CStr(mFileLenBytes) & "   " & "Seek: " & CStr(sk) & vbCrLf
    If df > 0 Then
      If df = 4 Then
        mess = mess & "Nochmal ohne FileFormatTag einlesen! (ModSongDescriptor.Read) -4"
      Else
        mess = mess & "zuviel eingelesen"
      End If
    Else
      mess = mess & "zuwenig eingelesen, evtl nochmal mit mehreren Tracks einlesen."
    End If
    MsgBox mess
  End If
End Sub

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("ModFile", aProcName, Addinfo)
End Sub
