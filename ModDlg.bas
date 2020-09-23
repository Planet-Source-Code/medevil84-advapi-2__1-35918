Attribute VB_Name = "ModDlg"
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_SHOWHELP = &H10

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Function OpenFile(xMain As Main, api As clsAPI, View As clsAPI, Optional bar As ProgressBar, Optional pStatusText As Panel) As Boolean
    '1  Display The Open dialog box
    '2  load the file in the application.
    '2.1    Load a Database file
    '2.2    Load a Text file
    
    OpenFile = True
    
    'Punto 1
    Dim op As OPENFILENAME
    Dim strMid As String
    Dim res As Long
    
    op.lStructSize = LenB(op)
    op.hwndOwner = xMain.hWnd
    op.lpstrFilter = CreateFilter("All compatible files", "*.mdb;*.txt", "Database Files (*.MDB)", "*.MDB", "Text Files (*.txt)", "*.txt", "All Files", "*.*")
    op.lpstrCustomFilter = String(100, Chr(0))
    op.nMaxCustFilter = 100
    op.nFilterIndex = 1
    op.lpstrFile = String(256, Chr(0))
    op.nMaxFile = 256
    op.lpstrFileTitle = String(256, Chr(0))
    op.nMaxFileTitle = 256
    op.flags = OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST '+ OFN_SHOWHELP
    op.lpstrDefExt = "mdb"
    
    res = GetOpenFileName(op)
    
    If res = 0 Then OpenFile = False: Exit Function 'Not opened
    
    Call SaveSetting(AppName, Section, "LastFile", GetVBstring(op.lpstrFile))
    
    'in custom filter c'è l'estensione
    
    Clear api.cConsts
    Clear api.cDeclares
    Clear api.cTypes
    
    Clear View.cConsts
    Clear View.cDeclares
    Clear View.cTypes
    
    'Punto 2
    'Controllo dell'estensione del file
    
    strMid = LCase(Right(op.lpstrFileTitle, Len(op.lpstrFileTitle) - InStr(1, op.lpstrFileTitle, ".", vbBinaryCompare)))
    
    strMid = Left(strMid, InStr(1, strMid, Chr(0), vbBinaryCompare) - 1)
    
    
    Select Case strMid
    Case "mdb"  ' E' un database...
        Call OpenFileDatabase(GetVBstring(op.lpstrFile), op.nFileOffset, api, bar, pStatusText)
    Case "txt"  ' E' un file di testo...
        Call OpenFileText(GetVBstring(op.lpstrFile), op.nFileOffset, api, bar, pStatusText)
    Case Else
        
        If MsgBox(GetLangString(lng_AskForFileInfo), vbYesNo, GetLangString(lng_AskForFileInfoTitle)) = vbYes Then
            Call OpenFileText(GetVBstring(op.lpstrFile), op.nFileOffset, api, bar, pStatusText)
        Else
            Call OpenFileDatabase(GetVBstring(op.lpstrFile), op.nFileOffset, api, bar, pStatusText)
        End If
    
    End Select
    
End Function
Public Sub OpenFileText(filePath As String, FileOffset As Integer, api As clsAPI, Optional bar As ProgressBar, Optional pStatusText As Panel)
    'Load a Text File
    Dim hFile As Long
    Dim pBuffer As String, tBuffer As String
    Dim bToRead As Long, bRead As Long, res As Long

    On Error Resume Next
    pStatusText.Text = lng_Loading
    On Error GoTo 0

    hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If (hFile <= 0) Then Exit Sub   'Error! file does not exists
        
    On Error Resume Next
        bar.Max = GetFileSize(hFile, ByVal 0) + 1
        bar.Min = 1
    On Error GoTo 0

    pBuffer = String(1024, Chr(0))
    bToRead = 1024

    res = ReadFile(hFile, ByVal pBuffer, bToRead, bRead, ByVal 0)
    
    If res = 0 And bRead = 0 Then
        'error!
        Call MsgBox(GetLangString(lng_ErrorOpeningFile), vbOKOnly + vbApplicationModal + vbCritical, GetLangString(lng_ErrorTitle))
        CloseHandle hFile
        Exit Sub
    ElseIf res > 0 Then
    
        tBuffer = pBuffer
        Do
            pBuffer = String(1024, Chr(0))
            res = ReadFile(hFile, ByVal pBuffer, bToRead, bRead, ByVal 0)
            tBuffer = tBuffer & pBuffer
            On Error Resume Next
            bar.Value = bar.Value + (bar.Max \ 1024)
            On Error GoTo 0
        Loop While bRead <> 0
    End If
    
    CloseHandle hFile

    Main.Caption = "ADVAPI 2 - [" & Mid(filePath, FileOffset + 1) & "]"
    
    'Now we need to fill info for the loaded text
    Call FillLists(tBuffer, api, bar, pStatusText)
    On Error Resume Next
    pStatusText.Text = ""
    On Error GoTo 0
    
End Sub

Sub LoadLanguage(filePath As String)

    'Load a Text File
    Dim hFile As Long
    Dim pBuffer As String, tBuffer As String
    Dim bToRead As Long, bRead As Long, res As Long

    hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If (hFile <= 0) Then Exit Sub   'Error! file does not exists
    
    pBuffer = String(1024, Chr(0))
    bToRead = 1024

    res = ReadFile(hFile, ByVal pBuffer, bToRead, bRead, ByVal 0)
    
    If res = 0 And bRead = 0 Then
        'error!
        Call MsgBox(GetLangString(lng_ErrorOpeningFile), vbOKOnly + vbApplicationModal + vbCritical, GetLangString(lng_ErrorTitle))
        CloseHandle hFile
        Exit Sub
    ElseIf res > 0 Then
    
        tBuffer = pBuffer
        Do
            pBuffer = String(1024, Chr(0))
            res = ReadFile(hFile, ByVal pBuffer, bToRead, bRead, ByVal 0)
            tBuffer = tBuffer & pBuffer
        Loop While bRead <> 0
    End If
    
    CloseHandle hFile

    'Sets Item
    
    Dim bList() As String
    Dim okList() As String
    Dim i As Long, c As Long
    
    bList = Split(tBuffer, vbCrLf)
    
    For i = LBound(bList) To UBound(bList)
        If Not (Left(bList(i), 1) = "#" Or Len(bList(i)) = 0) Then
            c = Val(Mid(bList(i), 1, InStr(1, bList(i), ".", vbBinaryCompare) - 1))
            Select Case c
            Case 1
                lng_Yes = Replace(bList(i), CStr(c) & ".", "")
            Case 2
                lng_No = Replace(bList(i), CStr(c) & ".", "")
            Case 3
                lng_ErrorTitle = Replace(bList(i), CStr(c) & ".", "")
            Case 4
                lng_ErrorOpeningFile = Replace(bList(i), CStr(c) & ".", "")
            Case 5
                lng_AskForFileInfo = Replace(bList(i), CStr(c) & ".", "")
            Case 6
                lng_AskForFileInfoTitle = Replace(bList(i), CStr(c) & ".", "")
            Case 7
                lng_Loading = Replace(bList(i), CStr(c) & ".", "")
            Case 8
                lng_Dec = Replace(bList(i), CStr(c) & ".", "")
            Case 9
                lng_Const = Replace(bList(i), CStr(c) & ".", "")
            Case 10
                lng_Type = Replace(bList(i), CStr(c) & ".", "")
            Case 11
                lng_Name = Replace(bList(i), CStr(c) & ".", "")
            Case 12
                lng_Lib = Replace(bList(i), CStr(c) & ".", "")
            Case 13
                lng_ReturnType = Replace(bList(i), CStr(c) & ".", "")
            Case 14
                lng_Params = Replace(bList(i), CStr(c) & ".", "")
            Case 15
                lng_Value = Replace(bList(i), CStr(c) & ".", "")
            Case 16
                lng_Public = Replace(bList(i), CStr(c) & ".", "")
            Case 17
                lng_Add = Replace(bList(i), CStr(c) & ".", "")
            Case 18
                lng_AddAll = Replace(bList(i), CStr(c) & ".", "")
            Case 19
                lng_Remove = Replace(bList(i), CStr(c) & ".", "")
            Case 20
                lng_RemAll = Replace(bList(i), CStr(c) & ".", "")
            Case 21
                lng_Dep = Replace(bList(i), CStr(c) & ".", "")
            Case 22 To 40
                lng_Menu(c - 22) = Replace(bList(i), CStr(c) & ".", "")
            Case 41 To 47
                lng_ToolBarTip(c - 41) = Replace(bList(i), CStr(c) & ".", "")
            Case 48
                lng_NoItems = Replace(bList(i), CStr(c) & ".", "")
            Case 49
                lng_SearchComplete = Replace(bList(i), CStr(c) & ".", "")
            Case 50
                lng_ErrorUnknowDB = Replace(bList(i), CStr(c) & ".", "")
            Case 51
                lng_LoadingCONST = Replace(bList(i), CStr(c) & ".", "")
            Case 52
                lng_LoadingTYPE = Replace(bList(i), CStr(c) & ".", "")
            Case 53
                lng_LoadingDECL = Replace(bList(i), CStr(c) & ".", "")
            End Select
        End If
    Next i
    
    Call SetLanguageItems
End Sub

Public Sub OpenFileDatabase(filePath As String, FileOffset As Integer, api As clsAPI, Optional bar As ProgressBar, Optional StatusText As Panel)
    Dim engine As New DBEngine
    Dim prop As Property
    Dim db As Database
    Dim i As Long, j As Long

    Dim unknowDB As Boolean, oldDB As Boolean

    If Not IsMissing(StatusText) Then StatusText.Text = lng_Loading

    Set db = engine.OpenDatabase(filePath)

    oldDB = True
    unknowDB = True

    For Each prop In db.Properties
        If prop.Name = "ProjVer" Then
            If CLng(prop.Value) = 24 Then
                oldDB = True
                unknowDB = False
            ElseIf CLng(prop.Value) = 200 Then
                oldDB = False
                unknowDB = False
            End If
        End If
    Next prop

    If unknowDB Then
        Select Case MsgBox(GetLangString(lng_ErrorUnknowDB), vbYesNoCancel, lng_ErrorTitle)
        Case vbYes
            oldDB = False
        Case vbNo
            oldDB = True
        Case vbCancel
            db.Close
            Exit Sub
        End Select
    End If

    If Not IsMissing(bar) Then
        bar.Min = 1
        bar.Max = 101
    End If

    Dim d As New apiDeclares
    Dim t As New apiType
    Dim c As New apiConst
    Dim p As New apiParams
    Dim rec2 As Recordset
    Dim nx() As String, tx() As String
    Dim sText As String

    If oldDB Then
        'Open OLD APIVIEW (or APILOAD) database
        Dim rec As Recordset

        If Not IsMissing(StatusText) Then StatusText.Text = lng_LoadingTYPE

        'Add type's...
        Set rec = db.OpenRecordset("Types", dbOpenDynaset, dbReadOnly)

        rec.MoveFirst
        Do
            t.decName = rec.Fields("Name").Value
            t.idKey = "t" & CStr(rec.Fields("ID").Value)
            
            api.cTypes.Add t, t.idKey
            rec.MoveNext
            '
            If (Not IsMissing(bar)) And (rec.AbsolutePosition <> -1) Then bar.Value = rec.PercentPosition + 1
        Loop While rec.AbsolutePosition > -1
        
        rec.Close
        
        Set rec = db.OpenRecordset("TypeItems", dbOpenDynaset, dbReadOnly)
        
        Do
            Set t = Nothing
            For Each t In api.cTypes
                If Val(Mid(t.idKey, 2)) = rec.Fields("TypeID").Value Then
                    p.paramName = PrimaParola(rec.Fields("TypeItem").Value)
                    p.paramType = RestoStringa(RestoStringa(rec.Fields("TypeItem").Value))
                    p.idKey = "p" & t.decParams.Count

                    t.decParams.Add p, p.idKey
                End If
            Next t

            rec.MoveNext
            If (Not IsMissing(bar)) And (rec.AbsolutePosition <> -1) Then bar.Value = rec.PercentPosition + 1
        Loop While rec.AbsolutePosition > -1
        
        rec.Close
        
        
        If Not IsMissing(StatusText) Then StatusText.Text = lng_LoadingCONST
        'Add constants

        Set rec = db.OpenRecordset("Constants", dbOpenDynaset, dbReadOnly)
        
        rec.MoveFirst
        Do
            sText = rec.Fields("Fullname").Value
            
            Call SelectType(sText, api)
            
            rec.MoveNext
            
            If (Not IsMissing(bar)) And (rec.AbsolutePosition <> -1) Then bar.Value = rec.PercentPosition + 1
        Loop While rec.AbsolutePosition > -1
        
        rec.Close
        
        
        If Not IsMissing(StatusText) Then StatusText.Text = lng_LoadingDECL
        'Add declares
        
        Set rec = db.OpenRecordset("Declares", dbOpenDynaset, dbReadOnly)
        
        rec.MoveLast
        sText = ""
        Do
            
            sText = rec.Fields("FullName").Value & sText
            
            If rec.Fields("ChunkNum").Value = 1 Then
                Call SelectType(sText, api)
                sText = ""
            End If
            
            rec.MovePrevious
            
            If (Not IsMissing(bar)) And (rec.AbsolutePosition <> -1) Then bar.Value = (bar.Max - (rec.PercentPosition + 1)) + 1
        Loop While rec.AbsolutePosition <> 1
        
        rec.Close
    Else

        
        Set rec = db.OpenRecordset("Declares", dbOpenDynaset, dbReadOnly)
        'Set rec2 = db.OpenRecordset("DeclareParams", dbOpenDynaset, dbReadOnly)
        
        'Opening Declares
        
        rec.MoveFirst
        Do While rec.AbsolutePosition > -1
        
            Set d = Nothing
            
            d.decSub = CBool(rec("Sub").Value)
            d.decName = CStr(rec("Name").Value)
            d.decLib = CStr(rec("Lib").Value)
            d.decAlias = CStr(rec("Alias").Value)
            
            'Add params to the declare statement
            If rec("paramName").FieldSize <> 0 And rec("paramType").FieldSize <> 0 Then
                nx() = Split(CStr(rec("paramName").Value), Chr(0))
                tx() = Split(CStr(IIf(Len(rec("paramType").Value) = 0, "", rec("paramType").Value)), Chr(0))
            ElseIf rec("paramName").FieldSize <> 0 And rec("paramType").FieldSize = 0 Then
                nx() = Split(CStr(rec("paramName").Value), Chr(0))
                ReDim tx(UBound(nx))
            ElseIf rec("paramType").FieldSize = 0 And rec("paramName").FieldSize <> 0 Then
                tx() = Split(CStr(IIf(Len(rec("paramType").Value) = 0, "", rec("paramType").Value)), Chr(0))
                ReDim nx(UBound(tx))
            Else
                ReDim nx(0)
                ReDim tx(0)
            End If

            
            For j = LBound(nx) To UBound(nx) - 1
                Set p = Nothing
                p.paramName = nx(j)
                p.paramType = tx(j)
                p.idKey = "p" & d.decParams.Count
                
                d.decParams.Add p, p.idKey
            Next j
            
            d.decReturnType = CStr(rec("ReturnType").Value)
            d.idKey = "d" & api.cDeclares.Count
            
            api.cDeclares.Add d, d.idKey
            
            If Not (IsMissing(bar)) Then bar.Value = rec.PercentPosition + 1
            
            rec.MoveNext
        Loop
        
        rec.Close
        
        'Opening Types
        Set rec = db.OpenRecordset("Type", dbOpenDynaset, dbReadOnly)
        
        rec.MoveFirst
        Do While rec.AbsolutePosition > -1
        
            Set t = Nothing
            
            t.decName = CStr(rec("Name").Value)
            'Add params to the declare statement
            
            If rec("paramName").FieldSize <> 0 And rec("paramType").FieldSize <> 0 Then
                nx() = Split(CStr(rec("paramName").Value), Chr(0))
                tx() = Split(CStr(IIf(Len(rec("paramType").Value) = 0, "", rec("paramType").Value)), Chr(0))
            ElseIf rec("paramName").FieldSize <> 0 And rec("paramType").FieldSize = 0 Then
                nx() = Split(CStr(rec("paramName").Value), Chr(0))
                ReDim tx(UBound(nx))
            ElseIf rec("paramType").FieldSize = 0 And rec("paramName").FieldSize <> 0 Then
                tx() = Split(CStr(IIf(Len(rec("paramType").Value) = 0, "", rec("paramType").Value)), Chr(0))
                ReDim nx(UBound(tx))
            Else
                ReDim nx(0)
                ReDim tx(0)
            End If
            
            For j = LBound(nx) To UBound(nx) - 1
                Set p = Nothing
                p.paramName = nx(j)
                p.paramType = tx(j)
                p.idKey = "p" & t.decParams.Count
                
                t.decParams.Add p, p.idKey
            Next j
            
            t.idKey = "t" & api.cTypes.Count
            
            api.cTypes.Add t, t.idKey
            
            If Not (IsMissing(bar)) Then bar.Value = rec.PercentPosition + 1
            
            rec.MoveNext
        Loop
        
        rec.Close
        
        'Opening Constants
        
        Set rec = db.OpenRecordset("Const", dbOpenDynaset, dbReadOnly)
        
        rec.MoveFirst
        
        Clear api.cConsts
        
        Do While rec.AbsolutePosition > -1
        
            Set c = Nothing
        
            c.decName = CStr(rec("Name").Value)
            c.decType = CStr(rec("Type").Value)
            c.decValue = CStr(rec("Value").Value)
            
            c.idKey = "c" & api.cConsts.Count
            
            api.cConsts.Add c, c.idKey
            
            If Not IsMissing(bar) Then bar.Value = rec.PercentPosition + 1
        
            rec.MoveNext
        Loop
        
        rec.Close
    End If
    
    Main.Caption = "ADVAPI 2 - [" & Mid(filePath, FileOffset + 1) & "]"
    
    db.Close
End Sub
