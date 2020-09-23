VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form convert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Converting Database..."
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   375
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
End
Attribute VB_Name = "convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

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

Public Sub ConvertFile(LoadFile As Boolean, xMain As Main)

    Dim op As OPENFILENAME
    Dim SaveAs As String
    Dim res As Long
    
    op.lStructSize = LenB(op)
    op.hwndOwner = xMain.hWnd
    op.lpstrFilter = "Database Files (*.MDB)" & Chr(0) & "*.MDB"
    op.lpstrCustomFilter = String(100, Chr(0))
    op.nMaxCustFilter = 100
    op.nFilterIndex = 1
    op.lpstrFile = String(256, Chr(0))
    op.nMaxFile = 256
    op.lpstrFileTitle = String(256, Chr(0))
    op.nMaxFileTitle = 256
    op.flags = OFN_EXPLORER + OFN_PATHMUSTEXIST '+ OFN_SHOWHELP
    op.lpstrDefExt = "mdb"
    
    res = GetSaveFileName(op)
    
    SaveAs = Mid(op.lpstrFile, 1, InStr(1, op.lpstrFile, Chr(0), vbTextCompare))
    
    If res = 0 Then
        MsgBox lng_ErrorOpeningFile, vbOKOnly + vbCritical, lng_ErrorTitle
        Exit Sub
    End If 'Not opened

    If LoadFile Then
        Dim cApi As New clsAPI
        Dim vApi As New clsAPI
        Call OpenFile(xMain, cApi, vApi)
        Set vApi = Nothing

        Call CreateDB(cApi, SaveAs)
    Else
        Call CreateDB(Main.apiList, SaveAs)
    End If
End Sub

Sub CreateDB(api As clsAPI, dbFile As String)
    
    Dim eng As New DBEngine
    Dim db As Database
    Dim tb As TableDef
    Dim fld As Field
    Dim ind As Index
    Dim rec As Recordset
    
    Dim d As apiDeclares
    Dim t As apiType
    Dim c As apiConst
    Dim p As apiParams
    Dim ns As String, ts As String
    
    convert.Show vbModeless, Main
    
    pBar.Min = 1
    pBar.Max = api.cConsts.Count + api.cDeclares.Count + api.cTypes.Count + 1
    
    Set db = eng.CreateDatabase(dbFile, dbLangGeneral)
    
    db.Properties.Append db.CreateProperty("ProjVer", dbInteger, 200)
    
    'Creating Tables
    Set tb = db.CreateTableDef("Declares")
    With tb
    
        .Fields.Append .CreateField("Sub", dbBoolean)
        .Fields.Append .CreateField("Name", dbText)
        .Fields.Append .CreateField("Lib", dbText)
        .Fields.Append .CreateField("Alias", dbText)
        .Fields("Alias").AllowZeroLength = True
        .Fields.Append .CreateField("paramName", dbMemo)
        .Fields("paramName").AllowZeroLength = True
        .Fields.Append .CreateField("paramType", dbMemo)
        .Fields("paramType").AllowZeroLength = True
        .Fields.Append .CreateField("ReturnType", dbText)
        .Fields("ReturnType").AllowZeroLength = True
        
    End With
    db.TableDefs.Append tb
    
    Set tb = db.CreateTableDef("Type")
    With tb
    
        .Fields.Append .CreateField("Name", dbText)
        .Fields.Append .CreateField("paramName", dbMemo)
        .Fields("paramName").AllowZeroLength = True
        .Fields.Append .CreateField("paramType", dbMemo)
        .Fields("paramType").AllowZeroLength = True
    End With
    db.TableDefs.Append tb

    Set tb = db.CreateTableDef("Const")
    With tb
        .Fields.Append .CreateField("Name", dbText)
        .Fields.Append .CreateField("Type", dbText)
        .Fields("Type").AllowZeroLength = True
        .Fields.Append .CreateField("Value", dbText)

    End With
    db.TableDefs.Append tb
    
    
    
    'Now Add stuffs to database
    pBar.Value = 1
    
    'First: Add Declares
    Set rec = db.OpenRecordset("Declares", dbOpenDynaset)
    
    'rec.MoveFirst
    For Each d In api.cDeclares
    
        rec.AddNew
    
        rec("Sub").Value = d.decSub
        rec("Name").Value = d.decName
        rec("Lib").Value = d.decLib
        rec("Alias").Value = d.decAlias
        
        ns = ""
        ts = ""
        For Each p In d.decParams
            ns = ns & p.paramName & Chr(0)
            ts = ts & p.paramType & Chr(0)
        Next p
        
        'Debug.Print Len(ns)
        
        rec("paramName").Value = ns
        rec("paramType").Value = ts
        
        rec("ReturnType").Value = d.decReturnType
        
        rec.Update
        
        rec.Bookmark = rec.LastModified
        
        pBar.Value = pBar.Value + 1
    Next d
    
    rec.Close
    
    'Now add types
    Set rec = db.OpenRecordset("Type", dbOpenDynaset)
    
    'rec.MoveFirst
    For Each t In api.cTypes
    
        rec.AddNew
    
        rec("Name").Value = t.decName
        
        ns = ""
        ts = ""
        For Each p In t.decParams
            ns = ns & p.paramName & Chr(0)
            ts = ts & p.paramType & Chr(0)
        Next p
        
        rec("paramName").Value = ns
        rec("paramType").Value = ts
        
        rec.Update
        
        rec.Bookmark = rec.LastModified
        
        pBar.Value = pBar.Value + 1
    Next t
    
    rec.Close
    
    
    'Now Add Constants
    Set rec = db.OpenRecordset("Const", dbOpenDynaset)

    'rec.MoveFirst
    For Each c In api.cConsts
    
        rec.AddNew
    
        rec("Name").Value = c.decName
        rec("Type").Value = c.decType
        rec("Value").Value = c.decValue
        
        rec.Update
        
        rec.Bookmark = rec.LastModified
        
        pBar.Value = pBar.Value + 1
    Next c
    rec.Close
    
    'Now close database...
    db.Close
    
    convert.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
End Sub
