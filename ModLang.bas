Attribute VB_Name = "ModLang"
Global lng_AskForFileInfo As String
Global lng_AskForFileInfoTitle As String
Global lng_ErrorTitle As String
Global lng_ErrorOpeningFile As String
Global lng_ErrorUnknowDB As String
Global lng_Loading As String
Global lng_Dec As String
Global lng_Const As String
Global lng_Type As String
Global lng_No As String
Global lng_Yes As String

Global lng_Name As String
Global lng_Lib As String
Global lng_ReturnType As String
Global lng_Params As String
Global lng_Value As String
Global lng_Public As String

Global lng_Add As String
Global lng_AddAll As String
Global lng_Remove As String
Global lng_RemAll As String
Global lng_Dep As String

Global lng_Menu(18) As String
Global lng_ToolBarTip(6) As String

Global lng_NoItems As String
Global lng_SearchComplete As String

Global lng_LoadingDECL As String
Global lng_LoadingTYPE As String
Global lng_LoadingCONST As String

Public Sub InitDefaultLang()
    lng_AskForFileInfo = "Please, select yes if you want to open\nselected file as a text file."
    lng_AskForFileInfoTitle = "Unable to determine type of file"
    lng_ErrorTitle = "Error!"
    lng_ErrorOpeningFile = "Unable to open selected file!"
    lng_Loading = "Loading and elaborating file..."
    
    lng_Dec = "Declarations"
    lng_Const = "Constants"
    lng_Type = "Types"
    lng_No = "No"
    lng_Yes = "Yes"
    lng_Name = "Name"
    lng_Lib = "Lib."
    lng_ReturnType = "Return Type"
    lng_Params = "Parameters"
    lng_Value = "Value"
    lng_Public = "Public"
    
    lng_Add = "Add   "
    lng_AddAll = "Add All"
    lng_Remove = "Remove "
    lng_RemAll = "Remove All"
    lng_Dep = "Dependencies Check"
    
    lng_Menu(0) = "&File"
    lng_Menu(1) = "&Open..."
    lng_Menu(2) = "&Close"
    lng_Menu(3) = "Open &last file"
    lng_Menu(4) = "E&xit"
    lng_Menu(5) = "&Edit"
    lng_Menu(6) = "Cu&t"
    lng_Menu(7) = "&Copy"
    lng_Menu(8) = "Pa&ste"
    lng_Menu(9) = "&Options"
    lng_Menu(10) = "&Convert file"
    lng_Menu(11) = "&Current file"
    lng_Menu(12) = "&Select file"
    lng_Menu(13) = "Select Language"
    lng_Menu(14) = "&Search"
    lng_Menu(15) = "&Find"
    lng_Menu(16) = "Find Ne&xt"
    lng_Menu(17) = "&Replace"
    lng_Menu(18) = "&About"
    
    lng_ToolBarTip(0) = "Open a new file..."
    lng_ToolBarTip(1) = "Close current file"
    lng_ToolBarTip(2) = "Cut selected text"
    lng_ToolBarTip(3) = "Copy selected text"
    lng_ToolBarTip(4) = "Paste current text into Visual Basic"
    lng_ToolBarTip(5) = "Find text"
    lng_ToolBarTip(6) = "About"
    
    lng_NoItems = "No items matching your query found!"
    lng_SearchComplete = "Search Complete!"
    
    lng_ErrorUnknowDB = "Warning! Unknown database type! Please select:\n\t'YES' to try opening current database as an ADVAPI 2 database.\n\t'NO' to try opening current database as an APIVIEW database.\n\t'CANCEL' to exit and ignore current database."
    
    lng_LoadingCONST = "Loadind Constants from database..."
    lng_LoadingDECL = "Loading Declares from database..."
    lng_LoadingTYPE = "Loading Types from database..."
End Sub

Public Function GetLangString(Text As String) As String
    GetLangString = Text
    GetLangString = Replace(GetLangString, "\n", vbCrLf)
    GetLangString = Replace(GetLangString, "\t", vbTab)
End Function

Sub SetLanguageItems()
    With Main
        With .tDown
        .Buttons(1).Caption = GetLangString(lng_Add)
        .Buttons(2).Caption = GetLangString(lng_Remove)
        .Buttons(1).ButtonMenus(1).Text = GetLangString(lng_AddAll)
        .Buttons(1).ButtonMenus(2).Text = GetLangString(lng_Dep)
        .Buttons(2).ButtonMenus(1).Text = GetLangString(lng_RemAll)
        .Buttons(2).ButtonMenus(2).Text = GetLangString(lng_Dep)
        End With
        
        .mFile.Caption = GetLangString(lng_Menu(0))
        .mOpen.Caption = GetLangString(lng_Menu(1))
        .mClose.Caption = GetLangString(lng_Menu(2))
        .mOpenLast.Caption = GetLangString(lng_Menu(3))
        .mExit.Caption = GetLangString(lng_Menu(4))
        .mEdit.Caption = GetLangString(lng_Menu(5))
        .mCut.Caption = GetLangString(lng_Menu(6))
        .mCopy.Caption = GetLangString(lng_Menu(7))
        .mPaste.Caption = GetLangString(lng_Menu(8))
        .mOptions.Caption = GetLangString(lng_Menu(9))
        .mConvert.Caption = GetLangString(lng_Menu(10))
        .mccOpened.Caption = GetLangString(lng_Menu(11))
        .mccSelect.Caption = GetLangString(lng_Menu(12))
        .mSelLang.Caption = GetLangString(lng_Menu(13))
        .mSearch.Caption = GetLangString(lng_Menu(14))
        .mFind.Caption = GetLangString(lng_Menu(15))
        .mFindNext.Caption = GetLangString(lng_Menu(16))
        .mReplace.Caption = GetLangString(lng_Menu(17))
        .mAbout.Caption = GetLangString(lng_Menu(18))
        
        With .tUp
        .Buttons(1).ToolTipText = GetLangString(lng_ToolBarTip(0))
        .Buttons(2).ToolTipText = GetLangString(lng_ToolBarTip(1))
        .Buttons(4).ToolTipText = GetLangString(lng_ToolBarTip(2))
        .Buttons(5).ToolTipText = GetLangString(lng_ToolBarTip(3))
        .Buttons(6).ToolTipText = GetLangString(lng_ToolBarTip(4))
        .Buttons(8).ToolTipText = GetLangString(lng_ToolBarTip(5))
        '.Buttons(7).ToolTipText = GetLangString(lng_ToolBarTip(6))
        End With
    End With

End Sub
