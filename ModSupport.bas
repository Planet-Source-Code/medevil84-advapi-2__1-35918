Attribute VB_Name = "ModSupport"
Option Compare Binary

Global Const AppName = "ADVAPI2"
Global Const Section = "Settings"


Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Global findResume As Long
Global findText As String
Global findJolly As Boolean
Global findMatch As Boolean
Global findList() As Long

Public Function CreateFilter(ParamArray xa() As Variant) As String

    CreateFilter = ""
    For i = LBound(xa) To UBound(xa)
        CreateFilter = CreateFilter & xa(i) & Chr(0)
    Next i

End Function
Public Function GetVBstring(Text As String) As String
    
    Dim byteArray() As Byte, i As Long
    
    byteArray = StrConv(Text, vbFromUnicode)
    
    For i = 0 To Len(Text)
        If byteArray(i) = 0 Then
            GetVBstring = Mid(StrConv(byteArray, vbUnicode), 1, i)
            i = Len(Text) + 1
        End If
    Next i
End Function

'Get the first word in a string
Public Function PrimaParola(Text As String) As String
    Dim l As Long
    l = InStr(1, Trim(Text), Space(1), vbTextCompare)
    If (l = 0) Or (Len(Text) = 0) Then PrimaParola = Text: Exit Function
    PrimaParola = Left(Trim(Text), l - 1)
End Function
'Get the string without the first word
Public Function RestoStringa(Text As String) As String
    Dim l As Long
    l = InStr(1, Trim(Text), Space(1), vbTextCompare)
    If (l = 0) Or (Len(Text) = 0) Then RestoStringa = "": Exit Function
    RestoStringa = Mid(Trim(Text), l + 1)
End Function

Public Sub Clear(col As Collection)
    
    Do While col.Count > 0
        col.Remove 1
    Loop
    
End Sub

Sub CopyItem(lDst As ListItem, lSrc As ListItem, PublicValue As Boolean)
    Dim subItem As ListSubItem
    Dim subItem2 As ListSubItem
    
    lDst.Bold = lSrc.Bold
    lDst.Checked = lSrc.Checked
    lDst.ForeColor = lSrc.ForeColor
    lDst.Ghosted = lSrc.Ghosted
    'lDst.Height = lSrc.Height
    lDst.Icon = lSrc.Icon
    'lDst.Index = lSrc.Index
    lDst.Key = lSrc.Key
    lDst.Left = lSrc.Left
    lDst.ListSubItems.Add , "public", IIf(PublicValue, "Yes", "No")
    For Each subItem In lSrc.ListSubItems
        Set subItem2 = Nothing
        Set subItem2 = lDst.ListSubItems.Add
        subItem2.Bold = subItem.Bold
        subItem2.ForeColor = subItem.ForeColor
        'subItem2.Index = subItem.Index
        subItem2.Key = subItem.Key
        subItem2.ReportIcon = subItem.ReportIcon
        subItem2.Tag = subItem.Tag
        subItem2.Text = subItem.Text
        subItem2.ToolTipText = subItem.ToolTipText
    Next subItem
    lDst.Selected = lSrc.Selected
    lDst.SmallIcon = lSrc.SmallIcon
    lDst.Tag = lSrc.Tag
    lDst.Text = lSrc.Text
    lDst.ToolTipText = lSrc.ToolTipText
    lDst.Top = lSrc.Top
    'lDst.Width = lSrc.Width
End Sub

Public Function FileExist(File As String) As Boolean

    Dim hFile As Long
    
    FileExist = False
    hFile = CreateFile(File, ByVal 0, ByVal 0, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If Not hFile = INVALID_HANDLE_VALUE Then
        FileExist = True
        Call CloseHandle(hFile)
    End If
    
End Function

Function SearchForFiles(col As Collection, Path As String) As Boolean

    Dim hFile As Long, res As Long
    Dim fnd As WIN32_FIND_DATA

    hFile = FindFirstFile(Path, fnd)
    
    If hFile = INVALID_HANDLE_VALUE Then SearchForFiles = False: Exit Function
    
    col.Add Left(fnd.cFileName, InStr(1, fnd.cFileName, Chr(0), vbBinaryCompare) - 1)
    
    Do
        res = FindNextFile(hFile, fnd)
        If Not (res = 0) Then col.Add Left(fnd.cFileName, InStr(1, fnd.cFileName, Chr(0), vbBinaryCompare) - 1)
    Loop While (res <> 0)

    res = FindClose(hFile)

End Function

Sub CopyStrToPtr(Text As String, Pointer As Long)
    
    If Pointer = 0 Then Exit Sub
    
    'Copy value of px (pointer to local variable) into pointer (pointer to a string)

    CopyMemory ByVal Pointer, ByVal StrPtr(Text), LenB(Text)
    
End Sub


Function Exist(Text As String, Find As String, Optional Jolly As Boolean = False, Optional cmpMethod As VbCompareMethod = vbBinaryCompare)

    Exist = False
    If Jolly Then
        If cmpMethod = vbBinaryCompare Then
            If Text Like Find Then Exist = True: Exit Function
        Else
            If LCase(Text) Like LCase(Find) Then Exist = True: Exit Function
        End If
    Else
        If cmpMethod = vbBinaryCompare Then
            If Text = Find Then Exist = True: Exit Function
        Else
            If LCase(Text) = LCase(Find) Then Exist = True: Exit Function
        End If
    End If
    
End Function


Sub Tmp()



End Sub
