VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 217
'       Byte  1   Byte  2   Byte  3   Byte 4
'      --------- --------- --------- ---------
'      7654-3210 7654-3210 7654-3210 7654-3210
'      wwww XXXX xxxxxxxxx yyyy ZZZZ zzzzzzzzz
'
'      wwwwyyyy ( 8 bits) : sample number
'      XXXXxxxxxxxx (12 bits) : sample 'period'
'      ZZZZzzzzzzzz (12 bits) : effect and argument

'Info for each note:
'
' _____byte 1______  byte2__  _____byte 3_______  byte4__
'/                 \/       \/                  \/       \
'0000          0000-00000000  0000          0000-00000000
'
'Upper 4 bits  12 bits for    Lower 4 bits  12 bits for
'of sample-    note period.   of sample-    effect command.
'number.                      number.
'
'(2^12 = 4096)
'Periodtable for Tuning 0, Normal
'  C-1 to B-1 : 856,808,762,720,678,640,604,570,538,508,480,453
'  C-2 to B-2 : 428,404,381,360,339,320,302,285,269,254,240,226
'  C-3 to B-3 : 214,202,190,180,170,160,151,143,135,127,120,113

Private Type ModNote
  B1 As Byte
  B2 As Byte
  B3 As Byte
  B4 As Byte
End Type
Public Enum ModEffects
  EfNone = &H0&           '- None/Arpeggio
  PortamentoUp = &H1&     '- Portamento Up
  PortamentoDown = &H2&   '- Portamento Down
  TonePortamento = &H3&   '- TonePortamento
  Vibrato = &H4&          '- Vibrato
  TonePVolSlide = &H5&    '- ToneP + VolSlide
  VibraVolSlide = &H6&    '- Vibra + VolSlide
  Tremolo = &H7&
  FinePanning = &H8 '8 - * NOT USED *
  SampleOffset = &H9&
  VolumeSlide = &HA&
  PositionJump = &HB&
  SetVolume = &HC&
  PatternBreak = &HD&
  MiscCmds = &HE&
  SetSpeed = &HF&
  'Extended E -commands
  'The E command has been altered to
  'contain more commands than one.
  FilterOnOff = &HE0&      '
  FineslideUp = &HE1&
  FineslideDown = &HE2&
  GlissandoControl = &HE3&
  VibratoControl = &HE4&
  SetFinetune = &HE5&
  Patternloop = &HE6&
  TremoloControl = &HE7&
  '* NOT USED * = &HE8&
  RetrigNote = &HE9&
  FineVolUp = &HEA&
  FineVolDown = &HEB&
  NoteCut = &HEC&
  NoteDelay = &HED&
  PatternDelay = &HEE&
  InvertLoop = &HEF&
End Enum
Private mModFile  As ModFile 'Zeiger auf das Document
Private mIndex    As Long    'der Index dieses Patterns '0-basiert
Private mChannels As Long    'die Anzahl an Channels (Tracks)
Private Type NoteLine
  l() As ModNote             'a note for each channel
End Type
'ein Pattern besteht immer aus einer festen Anzahl von 64 Zeilen
Private mPat(0 To 63) As NoteLine

Friend Sub New_(aModFile As ModFile, ByVal aIndex As Long)
  Set mModFile = aModFile
  mChannels = aModFile.MaxChannels
  mIndex = aIndex
End Sub
'Public Function ModPattern(aModFile As ModFile) As ModPattern
'  Set New_ModPattern = New ModPattern
'  Call New_ModPattern.NewC(aModFile)
'End Function

Public Sub Read(ByVal aFNr As Long)
Dim i As Long
  If mChannels > 0 Then
    For i = 0 To 63
      With mPat(i)
        ReDim .l(0 To mChannels - 1)
        Get aFNr, , .l
      End With
    Next
  End If
End Sub

Public Function ToString() As String
Dim i As Long, j As Long
  ToString = "Begin Pattern: " & CStr(mIndex) & vbCrLf
  For i = 0 To 63
  'bei 0 oder bei 1 beginnen? - bei 0
    ToString = ToString & PadLeft(CStr(i), 2) & " " & LineToString(i) & vbCrLf
  Next
  ToString = ToString & "End Pattern: " & CStr(mIndex) & vbCrLf
End Function

'##############################'   Public Properties   '##############################'
Public Property Get Index() As Long
  Index = mIndex
End Property

'               Byte  1   Byte  2   Byte  3   Byte 4
'              --------- --------- --------- ---------
'              7654-3210 7654-3210 7654-3210 7654-3210
'              wwww XXXX xxxxxxxxx yyyy ZZZZ zzzzzzzzz
'
'                  wwwwyyyy ( 8 bits) : sample number
'              XXXXxxxxxxxx (12 bits) : sample 'period'
'              ZZZZzzzzzzzz (12 bits) : effect and argument
'ln=LineNr, cn=ChannelNr
Public Property Get SampleNr(ln As Long, cn As Long) As Long
  With mPat(ln).l(cn): SampleNr = (&HF0 And .B1) + ((&HF0 And .B3) / &H10): End With
End Property
Public Property Let SampleNr(ln As Long, cn As Long, LngVal As Long)
  'n.y.i. not yet implemented
End Property

Public Property Get Note(ln As Long, cn As Long) As Long
  With mPat(ln).l(cn): Note = (&HF And .B1) * &HFF + .B2: End With
End Property
Public Property Let Note(ln As Long, cn As Long, LngVal As Long)
  'n.y.i. not yet implemented
End Property
Public Property Get SampleRate(ln As Long, cn As Long) As Double
  SampleRate = 7093789.2 / (2 * Note(ln, cn)) 'PAL
 'SampleRate = 7159090.5 / (2 * Note(ln, cn)) 'NTSC
End Property

'                C    C#   D    D#   E    F    F#   G    G#   A    A#   B
'     Octave 1: 856, 808, 762, 720, 678, 640, 604, 570, 538, 508, 480, 453
'     Octave 2: 428, 404, 381, 360, 339, 320, 302, 285, 269, 254, 240, 226
'     Octave 3: 214, 202, 190, 180, 170, 160, 151, 143, 135, 127, 120, 113
'
'     Octave 0:1712,1616,1525,1440,1357,1281,1209,1141,1077,1017, 961, 907
'     Octave 4: 107, 101,  95,  90,  85,  80,  76,  71,  67,  64,  60,  57

'(2^12 = 4096)
'Periodtable for Tuning 0, Normal
'  C-1 to B-1 : 856, 808, 762, 720, 678, 640, 604, 570, 538, 508, 480, 453
'           &H: 358, 328, 2FA, 2D0, 2A6, 280,
'  C-2 to B-2 : 428, 404, 381, 360, 339, 320, 302, 285, 269, 254, 240, 226
'  C-3 to B-3 : 214, 202, 190, 180, 170, 160, 151, 143, 135, 127, 120, 113
'kann man das irgendwie vereinfachen?
'z.B.: zuerst die Noten c-h, dann die Oktave 1-3
Public Function NoteToString(ln As Long, cn As Long) As String
Dim ni As Long, n As Long: n = Note(ln, cn)
Dim s As String
  If n = 0 Then
    NoteToString = "       " 'PadLeft(CStr(n), 3)
    Exit Function
  End If
'der Select Case reicht leider nicht, da auch FineTuning-Werte existieren
'zu jedem NotenWert existieren zusätzliche 15 verschiedene Finetuning-Werte
'-8...-1...+1...+7
'Nein, selbst das reicht nicht aus, die erweiterte Tabelle ist demnach auch
'nicht ausreichend
Dim i As Long, j As Long
  For i = 0 To 150
    For j = 0 To 1
      If j Then n = n - i Else n = n + i
      Select Case n
      Case 1712, 856, 428, 214, 107: s = "c-"
      Case 1616, 808, 404, 202, 101: s = "c#" 'oder auch "db", kommt auf die tonart an!
      Case 1525, 762, 381, 190, 95: s = "d-"
      Case 1440, 720, 360, 180, 90: s = "d#"
      Case 1357, 678, 339, 170, 85: s = "e-"
      Case 1281, 640, 320, 160, 80: s = "f-"
      Case 1209, 604, 302, 151, 76: s = "f#"
      Case 1141, 570, 285, 143, 71: s = "g-"
      Case 1077, 538, 269, 135, 67: s = "g#"
      Case 1017, 508, 254, 127, 64: s = "a-"
      Case 961, 480, 240, 120, 60: s = "a#"
      Case 907, 453, 226, 113, 57: s = "h-"
      End Select
      Select Case n
      Case 1712, 1616, 1525, 1440, 1357, 1281, 1209, 1141, 1077, 1017, 961, 907
        s = s & "0"
      Case 856, 808, 762, 720, 678, 640, 604, 570, 538, 508, 480, 453
        s = s & "1"
      Case 428, 404, 381, 360, 339, 320, 302, 285, 269, 254, 240, 226
        s = s & "2"
      Case 214, 202, 190, 180, 170, 160, 151, 143, 135, 127, 120, 113
        s = s & "3"
      Case 107, 101, 95, 90, 85, 80, 76, 71, 67, 64, 60, 57
        s = s & "4"
      'Case Else
      '  s = PadLeft(CStr(n), 3)
      End Select
      If Len(s) Then
        If i = 0 Then
          s = s & "    "
        Else
          If i < 0 Then
            s = s & "-"
          Else
            s = s & "+"
          End If
          s = s & PadRight(CStr(i), 3)
        End If
        Exit For
      End If
    Next
    If Len(s) Then Exit For
  Next
  'hmm wenn n ungleich Null ist, was ist es dann für eine Note ?
  If Len(s) = 0 Then s = PadRight(CStr(n), 7) '"       "
  NoteToString = s '& ""
  'NoteToString = PadLeft(Hex$(Note(ln, cn)), 2, "0")
End Function
'
'
'               Byte  1   Byte  2   Byte  3   Byte 4
'              --------- --------- --------- ---------
'              7654-3210 7654-3210 7654-3210 7654-3210
'              wwww XXXX xxxxxxxxx yyyy ZZZZ zzzzzzzzz
'
'              ZZZZzzzzzzzz (12 bits) : effect and argument
'               Bit number:         $CBA987654321
'               Mentioned above as:  ZZZZzzzzzzzz
'               We will use:         ZZZZxxxxyyyy
'
'There are two types of effects,  standard and extended.  All  effects use
'the ZZZZ portion to declare the effect number.  Standard effects  use the
'xxxx and yyyy portions as one or two arguments, either as an  8-bit value
'when taken together in the form  xxxxyyyy or as 2 nibbles xxxx  and yyyy.
'Extended effects  have the  ZZZZ effect  number $E.   They  use the  xxxx
'portion to  declare the  extended effect  number and  the  only the  yyyy
'portion as an argument.
Public Property Get Effect(ln As Long, cn As Long) As ModEffects
  With mPat(ln).l(cn): Effect = (&HF And .B3)
    If Effect = &HE Then
      Effect = &HE0 + (&HF0 And .B4) / &H10
    End If
  End With
  'With mPat(ln).L(cn): Effect = .B4: End With
End Property
Public Property Let Effect(ln As Long, cn As Long, LngVal As ModEffects)
  'n.y.i. not yet implemented
End Property


'#   Effect name                          Uses Arguments as
'-----------------------------------------------------------------------
'0   Arpeggio                             xxxx  yyyy
'1   Slide Up                             xxxxyyyy
'2   Slide Down                           xxxxyyyy
'3   Tone Portamento                      xxxxyyyy
'4   Vibrato                              xxxx  yyyy
'5   Tone Portamento + Volume Slide       xxxx  yyyy
'6   Vibrato + Volume Slide               xxxx  yyyy
'7   Tremolo                              xxxx  yyyy
'8   Set Panning Position                 xxxxyyyy
'9   Set SampleOffset                     xxxxyyyy
'A   VolumeSlide                          xxxx  yyyy
'B   Position Jump                        xxxxyyyy
'C   Set Volume                           xxxxyyyy
'D   Pattern Break                        xxxxyyyy
'E   *Extended Effects                    see below
'F   Set Speed                            xxxxyyyy
Private Function GetEffectArg1(ln As Long, cn As Long, Optional ByVal e As Long = -1) As Long
  If e < 0 Then e = Effect(ln, cn)
  With mPat(ln).l(cn)
    Select Case e
    Case 0, 4, 5, 6, 7, &HA '
      GetEffectArg1 = ((&HF0 And .B4) / &H10) '+ ((&HF And .B4) * &H10)
    Case 1, 2, 3, 8, 9, &HB, &HC, &HD, &HF
      GetEffectArg1 = .B4 '((&HF0 And .B4) / &H10)
    Case &HE
      GetEffectArg1 = (&HF And .B4)
    End Select
  End With
End Function
Public Property Get EffectArg1(ln As Long, cn As Long) As Long
  EffectArg1 = GetEffectArg1(ln, cn)
End Property
Public Property Let EffectArg1(ln As Long, cn As Long, LngVal As Long)
  'n.y.i. not yet implemented
End Property

Private Function GetEffectArg2(ln As Long, cn As Long, Optional ByVal e As Long = -1) As Long
  If e < 0 Then e = Effect(ln, cn)
  With mPat(ln).l(cn)
    Select Case e
    Case 0, 4, 5, 6, 7, &HA
      GetEffectArg2 = (&HF And .B4) '/ &H10)
    End Select
  End With
End Function
Public Property Get EffectArg2(ln As Long, cn As Long) As Long
  EffectArg2 = GetEffectArg2(ln, cn)
End Property
Public Property Let EffectArg2(ln As Long, cn As Long, LngVal As Long)
  'n.y.i. not yet implemented
End Property
Private Function EffectArgsToString(ln As Long, cn As Long, Optional ByVal e As Long = -1) As String
Dim ea1 As Long: ea1 = GetEffectArg1(ln, cn, e)
Dim ea2 As Long: ea2 = GetEffectArg2(ln, cn, e)
Dim s As String
  If ea1 >= 0 Then
    s = PadLeft(CStr(ea1), 2)
  Else
    s = "  "
  End If
  If ea2 >= 0 Then
    EffectArgsToString = s & " " & PadLeft(CStr(ea2), 2)
  Else
    EffectArgsToString = s & "   " '& PadLeft((ea2), 2)
  End If
End Function

'##############################'   Private Procedures   '##############################'
Private Function SampleNrToString(ln As Long, cn As Long)
  SampleNrToString = PadLeft(CStr(SampleNr(ln, cn)), 2)
End Function
Private Function ModNoteToString(ln As Long, cn As Long) As String
  ModNoteToString = Chr$(124) & SampleNrToString(ln, cn) & " " & NoteToString(ln, cn) _
                        & " " & EffectToString(ln, cn)  '& "|"
End Function
Private Function LineToString(lineNr As Long) As String
  Dim i As Long: For i = 0 To mChannels - 1
    LineToString = LineToString & ModNoteToString(lineNr, i)
  Next
End Function
Private Function EffectToString(ln As Long, cn As Long) As String
Dim e As ModEffects: e = Effect(ln, cn)
Dim sea As String: sea = EffectArgsToString(ln, cn, e)
Dim ea1 As Long:   ea1 = EffectArg1(ln, cn)
Dim ea2 As Long:   ea2 = EffectArg2(ln, cn)
Dim s As String
  If e = 0 And ea1 = 0 Then
    s = "   "
  Else
    s = PadLeft(CStr(Hex$(e)), 2) & " "
  End If
  Select Case e
  Case EfNone
      If ea1 = 0 Then
                         s = s & "             "
      Else
                         s = s & "Arpeggio     "
      End If
  Case PortamentoUp:     s = s & "Portam Up    "
  Case PortamentoDown:   s = s & "Portam Dwn   "
  Case TonePortamento:   s = s & "TonePortam   "
  Case Vibrato:          s = s & "Vibrato      "
  Case TonePVolSlide:    s = s & "TonePVolSlide"
  Case VibraVolSlide:    s = s & "VibraVolSlide"
  Case Tremolo:          s = s & "Tremolo      "
  Case FinePanning:      s = s & "FinePanning  "
  Case SampleOffset:     s = s & "SampleOffset "
  Case VolumeSlide:      s = s & "VolumeSlide  "
  Case PositionJump:     s = s & "PositionJump "
  Case SetVolume:        s = s & "SetVolume    "
  Case PatternBreak:     s = s & "PatternBreak "
  Case MiscCmds:         s = s & "MiscCmds     "
  Case SetSpeed:         s = s & "SetSpeed     "
  'Extended E -commands
  'The E command has been altered to
  'contain more effects than one.
  Case FilterOnOff:      s = s & "FilterOnOff  "
  Case FineslideUp:      s = s & "FineslideUp  "
  Case FineslideDown:    s = s & "FineslideDwn "
  Case GlissandoControl: s = s & "GlissandoCtl "
  Case VibratoControl:   s = s & "VibratoCtl   "
  Case SetFinetune:      s = s & "SetFinetune  "
  Case Patternloop:      s = s & "Patternloop  "
  Case TremoloControl:   s = s & "TremoloCtrl  "
  '* NOT USED * = &HE8&
  Case &HE8&:            s = s & " * UNUSED *  "
  Case RetrigNote:       s = s & "RetrigNote   "
  Case FineVolUp:        s = s & "FineVolUp    "
  Case FineVolDown:      s = s & "FineVolDwn   "
  Case NoteCut:          s = s & "NoteCut      "
  Case NoteDelay:        s = s & "NoteDelay    "
  Case PatternDelay:     s = s & "PatternDelay "
  Case PatternDelay:     s = s & "PatternDelay "
  Case Else:             s = s & " * Undef. *  "
  End Select
  EffectToString = s & sea
End Function

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("ModPattern", aProcName, Addinfo)
End Sub

