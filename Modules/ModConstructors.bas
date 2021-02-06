Attribute VB_Name = "ModConstructors"
Option Explicit '2007_05_17 Zeilen: 96

Public Function New_ModFile(aFileName As String) As ModFile
  Set New_ModFile = New ModFile
  Call New_ModFile.ReadFile(aFileName)
End Function
Public Function New_ModSampleInfos(aFNr As Long, ByVal count As Long, ByVal aFileOffset As Long) As ModSampleInfos
  Set New_ModSampleInfos = New ModSampleInfos
  Call New_ModSampleInfos.SetLength(count)
  Call New_ModSampleInfos.Read(aFNr, aFileOffset)
End Function
Public Function New_ModSongDescriptorF(aFNr As Long, ByVal aFileOffset As Long) As ModSongDescriptor
  Set New_ModSongDescriptorF = New ModSongDescriptor
  Call New_ModSongDescriptorF.Read(aFNr, aFileOffset)
End Function
Public Function New_ModPattern(ByVal aModFile As ModFile, ByVal aIndex As Long) As ModPattern
  Set New_ModPattern = New ModPattern
  Call New_ModPattern.NewC(aModFile, aIndex)
End Function
Public Function New_ModSound(aFNr As Long, ByVal lngSampleLen As Long) As ModSound
  Set New_ModSound = New ModSound
  Call New_ModSound.SetBufferLength(lngSampleLen)
  Call New_ModSound.Read(aFNr)
End Function
Public Function CModSound(ms As ModSound) As ModSound
  Set CModSound = ms
End Function

Public Function New_WaveSound(ByVal bps As Integer, ByVal nchannels As Integer, ByVal sps As Long) As WaveSound
  Set New_WaveSound = New WaveSound
  Call New_WaveSound.NewC(bps, nchannels, sps)
End Function
Public Function New_WaveSoundF(aFileName As String) As WaveSound
  Set New_WaveSoundF = New WaveSound
  Call New_WaveSoundF.NewF(aFileName)
End Function

Public Function MaxB(ByVal Byt1 As Byte, ByVal Byt2 As Byte) As Byte
  If Byt1 > Byt2 Then MaxB = Byt1 Else MaxB = Byt2
End Function
Public Function MaxL(ByVal Lng1 As Long, ByVal Lng2 As Long) As Long
  If Lng1 > Lng2 Then MaxL = Lng1 Else MaxL = Lng2
End Function
Public Property Get BE2LE(ByRef Bytes() As Byte) As Long 'Integer
'konvertiert einen 2 byte BigEndian(Amiga Int16) nach 2Byte LittleEndian(Intel Int16)
'besser nach Long, unsigned
  BE2LE = Bytes(0) * &H100& + Bytes(1)
End Property
Public Property Let BE2LE(ByRef Bytes() As Byte, LngVal As Long) 'Integer
'konvertiert einen 2 byte BigEndian(Amiga Int16) nach 2Byte LittleEndian(Intel Int16)
'besser nach Long, unsigned
  Bytes(0) = (&HFF00& And LngVal) \ 256
  Bytes(1) = (&HFF& And LngVal)
End Property

Public Property Get AS2US(ByRef astr() As Byte) As String
'Konvertiert Ansi-String im ByteArray nach UniCodeString
  Dim s As String: s = StrConv(astr, vbUnicode)
  Dim p As Long: p = InStr(1, s, vbNullChar)
  If p > 0 Then AS2US = Left$(s, p - 1)
End Property
Public Property Let AS2US(ByRef astr() As Byte, aStrVal As String) 'Str2BA
  Dim i As Long: For i = 0 To Len(aStrVal) - 1
    If i <= UBound(astr) Then astr(i) = Asc(Mid$(aStrVal, i + 1, 1))
  Next
End Property

Public Function PadLeft(StrVal As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
'der String wird mit der angegebenen Länge zurückgegeben, der
'String wird nach rechts gerückt, und links mit PadChar aufgefüllt
'ist PadChar nicht angegeben, so wird mit RSet der String in
'Spaces eingefügt.
  If Len(paddingChar) Then
    If Len(StrVal) <= totalWidth Then _
      PadLeft = String$(totalWidth - Len(StrVal), paddingChar) & StrVal
  Else
    PadLeft = Space$(totalWidth)
    RSet PadLeft = StrVal
  End If
End Function
Public Function PadRight(StrVal As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
'der String wird mit der angegebenen Länge zurückgegeben, der
'String wird nach links gerückt, und rechts mit PadChar aufgefüllt
'ist PadChar nicht angegeben, so wird mit LSet der String in
'Spaces eingefügt.
  If Len(paddingChar) Then
    If Len(StrVal) <= totalWidth Then _
      PadRight = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
  Else
    PadRight = Space$(totalWidth)
    LSet PadRight = StrVal
  End If
End Function

Public Function UByteToSInt16(aByte As Byte) As Integer
'wird von ModSound benützt um zu zeichnen
  If aByte <= &H7F Then UByteToSInt16 = aByte Else UByteToSInt16 = aByte - &HFF
End Function

