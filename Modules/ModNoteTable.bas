Attribute VB_Name = "ModNoteTable"
Option Explicit '2007_05_17 Zeilen: 96
Private Type TFT 'Finetunes
'0 to 7 == -4..0..+3
  FT(0 To 7) As Integer
End Type
Private Type TNT 'Halfnotes
  NT(0 To 11) As TFT
End Type
'Octaves
Private OT(0 To 4) As TNT
'OK find ich persönlich zwar nicht so schön, aber weil es die Empfehlung ist, und weil
'es wirklich sehr schlecht implementierte Datenreihen sind, deswegen hier als Tabelle
'Ufff
Public Sub InitNoteTable()
'Elemente im VB-Array3D sind in der Reihenfolge: Spalte, Zeile, Tabelle gespeichert
  With OT(0)
    With .NT(0):  .FT(0) = 1762: .FT(1) = 1750: .FT(2) = 1736: .FT(3) = 1724: .FT(4) = 1712: .FT(5) = 1700: .FT(6) = 1688: .FT(7) = 1676: End With
    With .NT(1):  .FT(0) = 1664: .FT(1) = 1652: .FT(2) = 1640: .FT(3) = 1628: .FT(4) = 1616: .FT(5) = 1604: .FT(6) = 1592: .FT(7) = 1582: End With
    With .NT(2):  .FT(0) = 1570: .FT(1) = 1558: .FT(2) = 1548: .FT(3) = 1536: .FT(4) = 1524: .FT(5) = 1514: .FT(6) = 1504: .FT(7) = 1492: End With
    With .NT(3):  .FT(0) = 1482: .FT(1) = 1472: .FT(2) = 1460: .FT(3) = 1450: .FT(4) = 1440: .FT(5) = 1430: .FT(6) = 1418: .FT(7) = 1408: End With
    With .NT(4):  .FT(0) = 1398: .FT(1) = 1388: .FT(2) = 1378: .FT(3) = 1368: .FT(4) = 1356: .FT(5) = 1348: .FT(6) = 1340: .FT(7) = 1330: End With
    With .NT(5):  .FT(0) = 1320: .FT(1) = 1310: .FT(2) = 1302: .FT(3) = 1292: .FT(4) = 1280: .FT(5) = 1274: .FT(6) = 1264: .FT(7) = 1256: End With
    With .NT(6):  .FT(0) = 1246: .FT(1) = 1238: .FT(2) = 1228: .FT(3) = 1220: .FT(4) = 1208: .FT(5) = 1202: .FT(6) = 1194: .FT(7) = 1184: End With
    With .NT(7):  .FT(0) = 1176: .FT(1) = 1168: .FT(2) = 1160: .FT(3) = 1150: .FT(4) = 1140: .FT(5) = 1134: .FT(6) = 1126: .FT(7) = 1118: End With
    With .NT(8):  .FT(0) = 1110: .FT(1) = 1102: .FT(2) = 1094: .FT(3) = 1086: .FT(4) = 1076: .FT(5) = 1070: .FT(6) = 1064: .FT(7) = 1056: End With
    With .NT(9):  .FT(0) = 1048: .FT(1) = 1040: .FT(2) = 1032: .FT(3) = 1026: .FT(4) = 1016: .FT(5) = 1010: .FT(6) = 1004: .FT(7) = 996:  End With
    With .NT(10): .FT(0) = 988:  .FT(1) = 982:  .FT(2) = 974:  .FT(3) = 968:  .FT(4) = 960:  .FT(5) = 954:  .FT(6) = 948:  .FT(7) = 940:  End With
    With .NT(11): .FT(0) = 934:  .FT(1) = 926:  .FT(2) = 920:  .FT(3) = 914:  .FT(4) = 906:  .FT(5) = 900:  .FT(6) = 894:  .FT(7) = 888:  End With
  End With
  With OT(1)
    With .NT(0):  .FT(0) = 881: .FT(1) = 875: .FT(2) = 868: .FT(3) = 862: .FT(4) = 856: .FT(5) = 850: .FT(6) = 844: .FT(7) = 838: End With
    With .NT(1):  .FT(0) = 832: .FT(1) = 826: .FT(2) = 820: .FT(3) = 814: .FT(4) = 808: .FT(5) = 802: .FT(6) = 796: .FT(7) = 791: End With
    With .NT(2):  .FT(0) = 785: .FT(1) = 779: .FT(2) = 774: .FT(3) = 768: .FT(4) = 762: .FT(5) = 757: .FT(6) = 752: .FT(7) = 746: End With
    With .NT(3):  .FT(0) = 741: .FT(1) = 736: .FT(2) = 730: .FT(3) = 725: .FT(4) = 720: .FT(5) = 715: .FT(6) = 709: .FT(7) = 704: End With
    With .NT(4):  .FT(0) = 699: .FT(1) = 694: .FT(2) = 689: .FT(3) = 684: .FT(4) = 678: .FT(5) = 674: .FT(6) = 670: .FT(7) = 665: End With
    With .NT(5):  .FT(0) = 660: .FT(1) = 655: .FT(2) = 651: .FT(3) = 646: .FT(4) = 640: .FT(5) = 637: .FT(6) = 632: .FT(7) = 628: End With
    With .NT(6):  .FT(0) = 623: .FT(1) = 619: .FT(2) = 614: .FT(3) = 610: .FT(4) = 604: .FT(5) = 601: .FT(6) = 597: .FT(7) = 592: End With
    With .NT(7):  .FT(0) = 588: .FT(1) = 584: .FT(2) = 580: .FT(3) = 575: .FT(4) = 570: .FT(5) = 567: .FT(6) = 563: .FT(7) = 559: End With
    With .NT(8):  .FT(0) = 555: .FT(1) = 551: .FT(2) = 547: .FT(3) = 543: .FT(4) = 538: .FT(5) = 535: .FT(6) = 532: .FT(7) = 528: End With
    With .NT(9):  .FT(0) = 524: .FT(1) = 520: .FT(2) = 516: .FT(3) = 513: .FT(4) = 508: .FT(5) = 505: .FT(6) = 502: .FT(7) = 498: End With
    With .NT(10): .FT(0) = 494: .FT(1) = 491: .FT(2) = 487: .FT(3) = 484: .FT(4) = 480: .FT(5) = 477: .FT(6) = 474: .FT(7) = 470: End With
    With .NT(11): .FT(0) = 467: .FT(1) = 463: .FT(2) = 460: .FT(3) = 457: .FT(4) = 453: .FT(5) = 450: .FT(6) = 447: .FT(7) = 444: End With
  End With
  With OT(2)
    With .NT(0):  .FT(0) = 441: .FT(1) = 437: .FT(2) = 434: .FT(3) = 431: .FT(4) = 428: .FT(5) = 425: .FT(6) = 422: .FT(7) = 419: End With
    With .NT(1):  .FT(0) = 416: .FT(1) = 413: .FT(2) = 410: .FT(3) = 407: .FT(4) = 404: .FT(5) = 401: .FT(6) = 398: .FT(7) = 395: End With
    With .NT(2):  .FT(0) = 392: .FT(1) = 390: .FT(2) = 387: .FT(3) = 384: .FT(4) = 381: .FT(5) = 379: .FT(6) = 376: .FT(7) = 373: End With
    With .NT(3):  .FT(0) = 370: .FT(1) = 368: .FT(2) = 365: .FT(3) = 363: .FT(4) = 360: .FT(5) = 357: .FT(6) = 355: .FT(7) = 352: End With
    With .NT(4):  .FT(0) = 350: .FT(1) = 347: .FT(2) = 345: .FT(3) = 342: .FT(4) = 339: .FT(5) = 337: .FT(6) = 335: .FT(7) = 332: End With
    With .NT(5):  .FT(0) = 330: .FT(1) = 328: .FT(2) = 325: .FT(3) = 323: .FT(4) = 320: .FT(5) = 318: .FT(6) = 316: .FT(7) = 314: End With
    With .NT(6):  .FT(0) = 312: .FT(1) = 309: .FT(2) = 307: .FT(3) = 305: .FT(4) = 302: .FT(5) = 300: .FT(6) = 298: .FT(7) = 296: End With
    With .NT(7):  .FT(0) = 294: .FT(1) = 292: .FT(2) = 290: .FT(3) = 288: .FT(4) = 285: .FT(5) = 284: .FT(6) = 282: .FT(7) = 280: End With
    With .NT(8):  .FT(0) = 278: .FT(1) = 276: .FT(2) = 274: .FT(3) = 272: .FT(4) = 269: .FT(5) = 268: .FT(6) = 266: .FT(7) = 264: End With
    With .NT(9):  .FT(0) = 262: .FT(1) = 260: .FT(2) = 258: .FT(3) = 256: .FT(4) = 254: .FT(5) = 253: .FT(6) = 251: .FT(7) = 249: End With
    With .NT(10): .FT(0) = 247: .FT(1) = 245: .FT(2) = 244: .FT(3) = 242: .FT(4) = 240: .FT(5) = 239: .FT(6) = 237: .FT(7) = 235: End With
    With .NT(11): .FT(0) = 233: .FT(1) = 232: .FT(2) = 230: .FT(3) = 228: .FT(4) = 226: .FT(5) = 225: .FT(6) = 224: .FT(7) = 222: End With
  End With
  With OT(3)
    With .NT(0):  .FT(0) = 220: .FT(1) = 219: .FT(2) = 217: .FT(3) = 216: .FT(4) = 214: .FT(5) = 213: .FT(6) = 211: .FT(7) = 209: End With
    With .NT(1):  .FT(0) = 208: .FT(1) = 206: .FT(2) = 205: .FT(3) = 203: .FT(4) = 202: .FT(5) = 201: .FT(6) = 199: .FT(7) = 198: End With
    With .NT(2):  .FT(0) = 196: .FT(1) = 195: .FT(2) = 193: .FT(3) = 192: .FT(4) = 190: .FT(5) = 189: .FT(6) = 188: .FT(7) = 187: End With
    With .NT(3):  .FT(0) = 185: .FT(1) = 184: .FT(2) = 183: .FT(3) = 181: .FT(4) = 180: .FT(5) = 179: .FT(6) = 177: .FT(7) = 176: End With
    With .NT(4):  .FT(0) = 175: .FT(1) = 174: .FT(2) = 172: .FT(3) = 171: .FT(4) = 170: .FT(5) = 169: .FT(6) = 167: .FT(7) = 166: End With
    With .NT(5):  .FT(0) = 165: .FT(1) = 164: .FT(2) = 163: .FT(3) = 161: .FT(4) = 160: .FT(5) = 159: .FT(6) = 158: .FT(7) = 157: End With
    With .NT(6):  .FT(0) = 156: .FT(1) = 155: .FT(2) = 154: .FT(3) = 152: .FT(4) = 151: .FT(5) = 150: .FT(6) = 149: .FT(7) = 148: End With
    With .NT(7):  .FT(0) = 147: .FT(1) = 146: .FT(2) = 145: .FT(3) = 144: .FT(4) = 143: .FT(5) = 142: .FT(6) = 141: .FT(7) = 140: End With
    With .NT(8):  .FT(0) = 139: .FT(1) = 138: .FT(2) = 137: .FT(3) = 136: .FT(4) = 135: .FT(5) = 134: .FT(6) = 133: .FT(7) = 132: End With
    With .NT(9):  .FT(0) = 131: .FT(1) = 130: .FT(2) = 129: .FT(3) = 128: .FT(4) = 127: .FT(5) = 126: .FT(6) = 125: .FT(7) = 125: End With
    With .NT(10): .FT(0) = 123: .FT(1) = 123: .FT(2) = 122: .FT(3) = 121: .FT(4) = 120: .FT(5) = 119: .FT(6) = 118: .FT(7) = 118: End With
    With .NT(11): .FT(0) = 117: .FT(1) = 116: .FT(2) = 115: .FT(3) = 114: .FT(4) = 113: .FT(5) = 113: .FT(6) = 112: .FT(7) = 111: End With
  End With
  With OT(4)
    With .NT(0):  .FT(0) = 110: .FT(1) = 109: .FT(2) = 108: .FT(3) = 108: .FT(4) = 107: .FT(5) = 106: .FT(6) = 105: .FT(7) = 104: End With
    With .NT(1):  .FT(0) = 104: .FT(1) = 103: .FT(2) = 102: .FT(3) = 101: .FT(4) = 101: .FT(5) = 100: .FT(6) = 99:  .FT(7) = 99:  End With
    With .NT(2):  .FT(0) = 98:  .FT(1) = 97:  .FT(2) = 96:  .FT(3) = 96:  .FT(4) = 95:  .FT(5) = 94:  .FT(6) = 94:  .FT(7) = 93:  End With
    With .NT(3):  .FT(0) = 92:  .FT(1) = 92:  .FT(2) = 91:  .FT(3) = 90:  .FT(4) = 90:  .FT(5) = 89:  .FT(6) = 88:  .FT(7) = 88:  End With
    With .NT(4):  .FT(0) = 87:  .FT(1) = 87:  .FT(2) = 86:  .FT(3) = 85:  .FT(4) = 85:  .FT(5) = 84:  .FT(6) = 83:  .FT(7) = 83:  End With
    With .NT(5):  .FT(0) = 82:  .FT(1) = 82:  .FT(2) = 81:  .FT(3) = 80:  .FT(4) = 80:  .FT(5) = 79:  .FT(6) = 79:  .FT(7) = 78:  End With
    With .NT(6):  .FT(0) = 78:  .FT(1) = 77:  .FT(2) = 77:  .FT(3) = 76:  .FT(4) = 75:  .FT(5) = 75:  .FT(6) = 74:  .FT(7) = 74:  End With
    With .NT(7):  .FT(0) = 73:  .FT(1) = 73:  .FT(2) = 72:  .FT(3) = 72:  .FT(4) = 71:  .FT(5) = 71:  .FT(6) = 70:  .FT(7) = 70:  End With
    With .NT(8):  .FT(0) = 69:  .FT(1) = 69:  .FT(2) = 68:  .FT(3) = 68:  .FT(4) = 67:  .FT(5) = 67:  .FT(6) = 66:  .FT(7) = 66:  End With
    With .NT(9):  .FT(0) = 65:  .FT(1) = 65:  .FT(2) = 64:  .FT(3) = 64:  .FT(4) = 63:  .FT(5) = 63:  .FT(6) = 62:  .FT(7) = 62:  End With
    With .NT(10): .FT(0) = 61:  .FT(1) = 61:  .FT(2) = 61:  .FT(3) = 60:  .FT(4) = 60:  .FT(5) = 59:  .FT(6) = 59:  .FT(7) = 59:  End With
    With .NT(11): .FT(0) = 58:  .FT(1) = 58:  .FT(2) = 57:  .FT(3) = 57:  .FT(4) = 56:  .FT(5) = 56:  .FT(6) = 56:  .FT(7) = 55:  End With
  End With
End Sub
Public Function KeyNoteToString(ByVal aKeyNote As Long) As String
Dim o As Long, n As Long, f As Long, found As Boolean
Dim s As String
  For o = 0 To 4
    For n = 0 To 11
      For f = 0 To 7
        If OT(o).NT(n).FT(f) = aKeyNote Then found = True: Exit For
      Next: If found Then Exit For
    Next: If found Then Exit For
  Next
  If Not found Then
    'Na wenn sie da nicht dabei war, dann womöglich noch in den
    'ersten und letzen Finetunigmöglichkeiten
    Select Case aKeyNote
    Case 1814, 1800, 1788, 1774: n = 0: o = 0
    Case 55, 54: n = 11: o = 4
    End Select
    Select Case aKeyNote
    Case 1814: f = -4
    Case 1800: f = -3
    Case 1788: f = -2
    Case 1774: f = -1
    Case 55:   f = 8
    Case 54:   f = 9
    End Select
  End If
  'wie heißt die Note?
  Select Case n
  Case 0:  s = "c "
  Case 1:  s = "c#"
  Case 2:  s = "d "
  Case 3:  s = "d#"
  Case 4:  s = "e "
  Case 5:  s = "f "
  Case 6:  s = "f#"
  Case 7:  s = "g "
  Case 8:  s = "g#"
  Case 9:  s = "a "
  Case 10: s = "a#"
  Case 11: s = "h "
  End Select
  'aus welcher Oktave
  s = s & CStr(o)
  'mit welchem Finetuning?
  'nur Finetuning <> 0 anzeigen
  f = f - 4
  If f = 0 Then
    s = s & "  "
  Else
    If f > 0 Then s = s & "+"
    s = s & CStr(f)
  End If
  KeyNoteToString = s
End Function

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
'die alte Funktion aus der Klasse ModPattern rausgenommen
Private Function NoteToString(ln As Long, cn As Long) As String
Dim ni As Long, n As Long: n = Note(ln, cn)
Dim s As String
  If n = 0 Then
    NoteToString = "   " 'PadLeft(CStr(n), 3)
    Exit Function
  End If
'Oh Shit, der Select Case reicht leider nicht, da auch FineTuning-Werte existieren
'zu jedem NotenWert existieren zusätzliche 15 verschiedene Finetuning-Werte
'-8...-1...+1...+7
Dim i As Long, j As Long
  For i = 0 To 150
    For j = 1 To 2
      If j = 1 Then n = n + i Else n = n - i
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
      If Len(s) Then Exit For
    Next
    If Len(s) Then Exit For
  Next
  NoteToString = s
  'NoteToString = PadLeft(Hex$(Note(ln, cn)), 2, "0")
End Function

