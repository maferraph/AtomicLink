Attribute VB_Name = "API"
Option Explicit
'****************************************************************
'                    C H A M A D A S     A P I
'****************************************************************
DefInt A-Z
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Class&, ByVal Caption&)
Declare Function GetWindow Lib "user32" (ByVal hwnd, ByVal wCmd)
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd, ByVal Buf$, ByVal lBuf)
Declare Sub ShowWindow Lib "user32" Alias "ShowWindowA" (ByVal hwnd, ByVal nCmd)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function GetSystemMenu& Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long)
Declare Function DeleteMenu& Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long)

Const WM_CLOSE = &H10
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const BIF_RETURNONLYFSDIRS = &H1
Const CSIDL_DRIVES = &H11
Public Const MF_BYPOSITION& = &H400&

Type SHITEMID
     cb As Long
     abID As Byte
End Type
Type ITEMIDLIST
     mkid As SHITEMID
End Type
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Function GetWinCaption$(hwnd)
    Dim strCaption As String * 256
    Dim NumChars
    NumChars = GetWindowText(hwnd, strCaption, 256)
    GetWinCaption$ = Left$(strCaption, NumChars)
End Function
Sub FechaArj()
    Dim Wnd, lResult As Long, iResult, I
    Do While SearchWindow("Concluído - Atomic Link/arj") <> 0 Or SearchWindow("Atomic Link/arj") <> 0
        Wnd = SearchWindow("Concluído - Atomic Link/arj")
        lResult = SendMessage(Wnd, WM_CLOSE, 0, 0&)
    Loop
End Sub
Function SearchWindow(ByVal strSearch$)
    Dim hwnd
    hwnd = FindWindow(0, 0)
    hwnd = GetWindow(hwnd, GW_HWNDFIRST)
    strSearch$ = UCase$(strSearch$)
    Do While hwnd <> 0
        If InStr(UCase$(GetWinCaption(hwnd)), strSearch$) Then
            SearchWindow = hwnd
            Exit Function
        End If
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT)
    Loop
End Function
