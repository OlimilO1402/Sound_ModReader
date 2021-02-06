Attribute VB_Name = "MDefInetBrowser"
Option Explicit
'Added 06.feb.2021
Private Declare Function ShellExecute Lib "shell32.dll" _
                 Alias "ShellExecuteA" ( _
                 ByVal hWnd As Long, _
                 ByVal lpOperation As String, _
                 ByVal lpFile As String, _
                 ByVal lpParameters As String, _
                 ByVal lpDirectory As String, _
                 ByVal nShowCmd As Long) As Long
                 
Private Declare Function RegOpenKey Lib "advapi32.dll" _
  Alias "RegOpenKeyA" ( _
  ByVal hKey As Long, _
  ByVal lpSubKey As String, _
  phkResult As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
  Alias "RegOpenKeyExA" ( _
  ByVal hKey As Long, ByVal lpSubKey As String, _
  ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
  Alias "RegQueryValueExA" ( _
  ByVal hKey As Long, ByVal lpValueName As String, _
  ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
 
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const REG_SZ = 1

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_READ As Long = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const ERROR_SUCCESS = 0&

Public Sub Start(ByVal hWnd As Long, httpInetPage As String)
Try: On Error GoTo Finally
    Dim hr As Long
    hr = ShellExecute(hWnd, "Open", httpInetPage, "", App.Path, 0)
Finally:
    If Err.LastDllError Then
        MsgBox Err.Description
    End If
End Sub

Public Function GetDefaultBrowser(sBrowserName As String, sBrowserPath As String, sHTTPType As String) As Boolean
   Dim sHTTP As String
   
   sHTTP = GetHTTPType()
   
   '-- Fehlerbehandlung für WinXP und früher...
   If Len(sHTTP) = 0 Then sHTTP = "HTTP"
   
   sHTTPType = sHTTP
   sBrowserPath = GetBrowserPath(sHTTP)
   sBrowserName = sHTTP
    
    '-- Hier könnte man Tipp 0199 oder so zur weiteren Fehlerbehandlung anschließen.
 '  If Len(sBrowserPath) = 0 Then
 '     sHTTPType = "File"
 '     sBrowserPath = GetDefaultBrowserFormFile(sHTTP)
 '     sBrowserName = sHTTP
 '  End If
   
   If Len(sBrowserPath) > 0 Then GetDefaultBrowser = True
End Function

Private Function GetBrowserPath(Browser As String) As String
Dim Result As Long
Dim hKey As Long
Dim dwType As Long
Dim l As Long
Dim Buffer As String

    On Error GoTo BrowserErr
    
    'Wert aus dem Feld der Registry auslesen
    Result = RegOpenKeyEx(HKEY_CLASSES_ROOT, Browser & "\shell\open\command", 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then
        Result = RegQueryValueEx(hKey, "", 0&, dwType, ByVal 0&, l)
        If Result = ERROR_SUCCESS Then
            If dwType = REG_SZ Then
                ' Wert auslesen
                Buffer = Space$(l + 1)
                Result = RegQueryValueEx(hKey, "", 0&, dwType, ByVal Buffer, l)
                 
                Buffer = Trim(Buffer)
                ' Anführungszeichen entfernen
                Buffer = Replace(Buffer, """", "")
                ' Parameter entfernen...
                Buffer = (Left(Buffer, (InStr(1, Buffer, ".exe", vbTextCompare) + 4)))
                Buffer = Trim(Left$(Buffer, l - 1))
                'BrowserName festlegen (Unten)
                Browser = GetBrowserName(Buffer)
            Else
               GoTo BrowserErr
            End If
        Else
           GoTo BrowserErr
        End If
    Else
        GoTo BrowserErr
    End If
    
    RegCloseKey hKey

    GetBrowserPath = Buffer

    Exit Function

BrowserErr:
   On Error GoTo 0
   
   RegCloseKey hKey
   GetBrowserPath = ""
    
End Function

Private Function GetHTTPType() As String
    Dim Result As Long
    Dim hKey As Long
    Dim dwType As Long
    Dim l As Long
    Dim Buffer As String
    
    'Wert aus dem Feld der Registry auslesen
    Result = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then
        Result = RegQueryValueEx(hKey, "Progid", 0&, dwType, ByVal 0&, l)
        If Result = ERROR_SUCCESS Then
           If dwType = REG_SZ Then
              Buffer = Space$(l + 1)
              Result = RegQueryValueEx(hKey, "Progid", 0&, dwType, ByVal Buffer, l)
            
              Buffer = Trim(Buffer)
              Buffer = Left(Buffer, Len(Buffer) - 1)
           End If
        Else
           Buffer = ""
        End If
    Else
        Buffer = ""
    End If
    
    RegCloseKey hKey
    
    GetHTTPType = Buffer
    
End Function

Private Function GetBrowserName(Buffer As String) As String
    If Buffer <> "" Then
       If InStr(LCase$(Buffer), "iexplore") > 0 Then
         GetBrowserName = "Microsoft Internet Explorer"
       ElseIf InStr(LCase$(Buffer), "netscape") > 0 Then
         GetBrowserName = "Netscape Communicator"
       ElseIf InStr(LCase$(Buffer), "firefox") > 0 Then
         GetBrowserName = "Mozilla Firefox"
       ElseIf InStr(LCase$(Buffer), "opera") > 0 Then
         GetBrowserName = "Opera Browser"
       ElseIf InStr(LCase$(Buffer), "chrome") > 0 Then
         GetBrowserName = "Google Chrome"
       ElseIf InStr(LCase$(Buffer), "launchwinapp") > 0 Then
         GetBrowserName = "Microsoft Edge"
       Else
         GetBrowserName = "Unbekannter Browser"
       End If
   End If

End Function


