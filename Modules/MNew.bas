Attribute VB_Name = "MNew"
Option Explicit '2007_05_17 Zeilen: 96

Public Function ModFile(aFileName As String) As ModFile
    Set ModFile = New ModFile: ModFile.ReadFile aFileName
End Function
Public Function ModSampleInfos(aFNr As Long, ByVal count As Long, ByVal aFileOffset As Long) As ModSampleInfos
    Set ModSampleInfos = New ModSampleInfos: ModSampleInfos.SetLength count: ModSampleInfos.Read aFNr, aFileOffset
End Function
Public Function ModSongDescriptorF(aFNr As Long, ByVal aFileOffset As Long) As ModSongDescriptor
    Set ModSongDescriptorF = New ModSongDescriptor: ModSongDescriptorF.Read aFNr, aFileOffset
End Function
Public Function ModPattern(ByVal aModFile As ModFile, ByVal aIndex As Long) As ModPattern
    Set ModPattern = New ModPattern: ModPattern.New_ aModFile, aIndex
End Function
Public Function ModSound(aFNr As Long, ByVal lngSampleLen As Long) As ModSound
    Set ModSound = New ModSound: ModSound.SetBufferLength lngSampleLen: ModSound.Read aFNr
End Function
Public Function CModSound(ms As ModSound) As ModSound
    Set CModSound = ms
End Function

Public Function WaveSound(ByVal bps As Integer, ByVal nchannels As Integer, ByVal sps As Long) As WaveSound
    Set WaveSound = New WaveSound: WaveSound.New_ bps, nchannels, sps
End Function
Public Function WaveSoundF(aFileName As String) As WaveSound
    Set WaveSoundF = New WaveSound: WaveSoundF.NewF aFileName
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

