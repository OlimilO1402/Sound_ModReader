VERSION 5.00
Begin VB.Form FModReader 
   Caption         =   "Mod-Reader"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnGoto 
      Caption         =   "Goto"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton BtnPoints 
      Caption         =   "Points"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton BtnSaveWave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton BtnStop 
      Caption         =   "[ ] Stop"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton BtnPlay 
      Caption         =   "|> Play"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox CmbSample 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton BtnOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox CmbFileName 
      Height          =   315
      Left            =   1560
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   8055
   End
   Begin VB.ComboBox CmbDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox TxtModFile 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   960
      Width           =   10455
   End
   Begin VB.PictureBox PBWaveForm 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   7515
      TabIndex        =   7
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "FModReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '2007_05_17 Zeilen: 163
Private mMod As ModFile
Private mCurSound As ModSound
Private Declare Function PlaySoundA Lib "winmm.dll" (ByRef pArr As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_MEMORY As Long = &H4
Private Const SND_ASYNC As Long = &H1
Private mWFNCounter As Long
Private mBolPoints As Boolean

Private Sub Form_Load()
  With CmbFileName
    .Clear
    'Call .AddItem("http://modarchive.org/index_1.php")
    Call .AddItem("https://modarchive.org/") 'changed 06.feb.2021
    'Call .AddItem(App.Path & "\BspModFiles\ACIDOFIL.MOD")
    'Call .AddItem(App.Path & "\BspModFiles\ACID_AGE.MOD")
    Call .AddItem(App.Path & "\ExampleMods\POWER.MOD")
    'Call .AddItem(App.Path & "\BspModFiles\1Lars.MOD")
    'Call .AddItem(App.Path & "\BspModFiles\1Tubell.mod")
    'Call .AddItem(App.Path & "\BspModFiles\1powerem2.mod")
    .ListIndex = 0
  End With
  BtnOpen.Caption = "Open"
  BtnGoto.Caption = "Goto"
  BtnPlay.Caption = "|> Play"
  BtnStop.Caption = Chr$(216) & " Stop" '"[ ] Stop"
  BtnSaveWave.Caption = "Save"
  
  CmbDisplay.Text = vbNullString
  CmbSample.Text = vbNullString
  PBWaveForm.BackColor = &H40&
  PBWaveForm.ForeColor = &HFF0000
End Sub
Private Sub Form_Resize()
Dim l As Single, T As Single, W As Single, H As Single
Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
  
  l = brdr: T = brdr:  W = 7 * brdr: H = 3 * brdr
  If W > 0 And H > 0 Then Call BtnOpen.Move(l, T, W, H)
  
  l = l + W: W = 5 * brdr
  If W > 0 And H > 0 Then Call BtnGoto.Move(l, T, W, H)
  
  l = l + W: W = Me.ScaleWidth - l - brdr
  If W > 0 And H > 0 Then Call CmbFileName.Move(l, T, W) ', H)
  
  l = brdr: T = T + H + brdr: W = 27 * brdr
  If W > 0 And H > 0 Then Call CmbDisplay.Move(l, T, W) ', H)
  
  l = l + W + brdr
  If W > 0 And H > 0 Then Call CmbSample.Move(l, T, W) ', H)
  
  l = l + W: W = 6 * brdr: H = 3 * brdr
  If W > 0 And H > 0 Then Call BtnPlay.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnStop.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnSaveWave.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnPoints.Move(l, T, W, H)
  
  l = 0: T = T + H:
  W = Me.ScaleWidth '- L - brdr
  H = Me.ScaleHeight - T '- brdr
  If W > 0 And H > 0 Then Call TxtModFile.Move(l, T, W, H)
  If W > 0 And H > 0 Then Call PBWaveForm.Move(l, T, W, H)
  PBWaveForm.Refresh
End Sub

Private Sub InitCmbDisplay()
Dim i As Long
  With CmbDisplay
    .Clear
    Call .AddItem("SongInfos")
    For i = 0 To mMod.Patterns.count - 1
      Call .AddItem("Pattern: " & CStr(i))
    Next
    '.ListIndex = 0
  End With
End Sub
Private Sub InitCmbSample()
Dim s As String, i As Long, c1 As Long, c2 As Long
  With CmbSample
    .Clear
    For i = 0 To mMod.MaxSamples - 1
      s = Trim$(mMod.SampleInfoBank.SampleName(i))
      If Len(s) = 0 Then
        c1 = c1 + 1
        s = " * UNNAMED " & CStr(c1) & " * "
      Else
        If IsInComboBox(CmbSample, s) Then
          c2 = c2 + 1
          s = s & " " & CStr(c2)
        End If
      End If
      Call .AddItem(s)
    Next
    '.ListIndex = 0
  End With
End Sub
Private Function IsInComboBox(aCB As ComboBox, StrVal As String) As Boolean
Dim i As Long 'v As String
  For i = 0 To aCB.ListCount - 1
    If StrComp(StrVal, aCB.List(i)) = 0 Then
      IsInComboBox = True: Exit Function
    End If
  Next
End Function

Private Sub BtnOpen_Click()
TryE: On Error GoTo CatchE
  If Left$(CmbFileName.Text, 4) = "http" Then
    'Dim rv As Double
    'hier Internet Browser Programm anpassen
    'rv = Shell("C:\Programme\Internet Explorer\iexplore.exe " & CmbFileName.Text)
    'rv = Shell(CmbFileName.Text, vbNormalFocus)
    MDefInetBrowser.Start Me.hWnd, CmbFileName.Text
  Else
    Set mMod = New_ModFile(CmbFileName.Text)
    Call InitCmbSample
    Call InitCmbDisplay
    CmbDisplay.ListIndex = 0
    TxtModFile.Text = mMod.ToString
  End If
  Exit Sub
CatchE:
  MsgBox Err.Description
End Sub

Private Sub BtnGoto_Click()
Dim D As String: D = CmbFileName.Text
  Call Shell("explorer.exe " & Left$(D, InStrRev(D, "\")), vbNormalFocus)
End Sub

Private Sub CmbFileName_OLEDragDrop(Data As DataObject, Effect As Long, _
  Button As Integer, Shift As Integer, x As Single, y As Single)
  If Data.GetFormat(vbCFFiles) Then CmbFileName.Text = Data.Files.Item(1)
End Sub

Private Sub BtnPlay_Click()
  If CmbSample.ListIndex >= 0 Then
    Dim ms As ModSound
    Set ms = mMod.Sounds.Item(CmbSample.ListIndex + 1)
    Dim rv As Long
    rv = PlaySoundA(ByVal ms.Wave.pSndBuff, 0&, SND_MEMORY Or SND_ASYNC)
  Else
    MsgBox "open file first!"
  End If
End Sub
Private Sub BtnStop_Click()
  Dim rv As Long
  rv = PlaySoundA(0&, 0&, 0&)
End Sub
Private Sub BtnSaveWave_Click()
  If CmbSample.ListIndex >= 0 Then
    Dim pfn As String: pfn = GetNextWaveFileName
    Dim msg As String: msg = "Speichere wav-Datei nach: " & vbCrLf & pfn
    Dim mr As VbMsgBoxResult: mr = MsgBox(msg, vbOKCancel)
    If mr = vbOK Then
      'Dim s As ModSound: Set s = mMod.Sounds.Item(CmbSample.ListIndex + 1)
      'Call s.Wave.SaveToFile(pfn)
      Call mMod.Sound(CmbSample.ListIndex + 1).Wave.SaveToFile(pfn)
    End If
  End If
End Sub
Private Function GetNextWaveFileName() As String
Dim fn As String, u As String
Dim i As Long
  fn = CmbSample.List(CmbSample.ListIndex)
  'von allem Unrat befreien!
  u = "@:!.,\?+*#'-´`()[]{}="""
  For i = 1 To Len(u)
    fn = Replace(fn, Mid$(u, i, 1), "_")
  Next
  mWFNCounter = mWFNCounter + 1
  GetNextWaveFileName = App.Path & "\" & "Wave" & CStr(mWFNCounter) & "_" & fn & ".wav"
End Function
Private Sub BtnPoints_Click()
'nur toggeln
  If mBolPoints Then
    BtnPoints.Caption = "Points"
  Else
    BtnPoints.Caption = "Lines"
  End If
  mBolPoints = Not mBolPoints
  PBWaveForm.Refresh
End Sub

Private Sub CmbDisplay_Click()
  TxtModFile.ZOrder 0
  If CmbDisplay.ListIndex = 0 Then
    TxtModFile.Text = mMod.ToString
  Else
    Dim p As ModPattern
    Set p = mMod.Patterns.Item(CmbDisplay.ListIndex)
    TxtModFile.Text = p.ToString
  End If
End Sub

Private Sub CmbSample_Click()
  Set mCurSound = mMod.Sounds.Item(CmbSample.ListIndex + 1)
  PBWaveForm.ZOrder 0
  PBWaveForm.Cls
  PBWaveForm.Refresh
End Sub

Private Sub PBWaveForm_Paint()
  If Not mCurSound Is Nothing Then _
    Call mCurSound.DrawToPictureBox(PBWaveForm, mBolPoints)
End Sub

