VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Main 
   BackColor       =   &H8000000B&
   Caption         =   "ADVAPI 2"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3780
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSComctlLib.ImageList iTree 
      Left            =   8220
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iToolBar 
      Left            =   8220
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":101C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":112E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1240
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iTool2 
      Left            =   8190
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1352
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":16A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pSizeH 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   30
      ScaleHeight     =   60
      ScaleWidth      =   6885
      TabIndex        =   11
      Top             =   2985
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.PictureBox picCont2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   8085
      TabIndex        =   8
      Top             =   3270
      Width           =   8085
      Begin VB.PictureBox pSizeV2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   2025
         Left            =   2310
         ScaleHeight     =   2025
         ScaleWidth      =   75
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   75
      End
      Begin MSComctlLib.Toolbar tDown 
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iTool2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add   "
               Key             =   "kAdd"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "aAll"
                     Text            =   "Add All"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "aDep"
                     Text            =   "Dependencies Check"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Remove "
               Key             =   "kRemove"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "rAll"
                     Text            =   "Remove All"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Key             =   "rDep"
                     Text            =   "Dependencies Check"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Begin VB.OptionButton opPublic 
            Caption         =   "Public"
            Height          =   195
            Left            =   3540
            TabIndex        =   13
            Top             =   90
            Width           =   825
         End
         Begin VB.OptionButton opPriv 
            Caption         =   "Private"
            Height          =   195
            Left            =   2640
            TabIndex        =   12
            Top             =   90
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin MSComctlLib.ListView lView 
         Height          =   1815
         Left            =   2130
         TabIndex        =   10
         Top             =   330
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tView 
         Height          =   1695
         Left            =   30
         TabIndex        =   14
         Top             =   360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   2990
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   471
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "iTree"
         Appearance      =   1
      End
      Begin VB.Image iSizeV2 
         Height          =   2085
         Left            =   2010
         MousePointer    =   9  'Size W E
         Top             =   210
         Width           =   105
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   6690
      TabIndex        =   3
      Top             =   6000
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5940
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11783
            MinWidth        =   2893
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2893
            MinWidth        =   2893
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   688
      BandCount       =   1
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   8610
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "tUp"
      MinHeight1      =   330
      Width1          =   3135
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar tUp 
         Height          =   330
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "iToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kOpen"
               Object.ToolTipText     =   "Open a new file..."
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kClose"
               Object.ToolTipText     =   "Close current file"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kCut"
               Object.ToolTipText     =   "Cut selected text"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kCopy"
               Object.ToolTipText     =   "Copy selected text"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "kPaste"
               Object.ToolTipText     =   "Paste current text into Visual Basic"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kFind"
               Object.ToolTipText     =   "Find text"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picCont 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   8085
      TabIndex        =   4
      Top             =   450
      Width           =   8085
      Begin VB.PictureBox pSizeV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   2025
         Left            =   2250
         ScaleHeight     =   2025
         ScaleWidth      =   75
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   75
      End
      Begin MSComctlLib.ListView lApi 
         Height          =   1815
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tApi 
         Height          =   2115
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   3731
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   471
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "iTree"
         Appearance      =   1
      End
      Begin VB.Image iSizeV 
         Height          =   2085
         Left            =   1950
         MousePointer    =   9  'Size W E
         Top             =   60
         Width           =   105
      End
   End
   Begin VB.Image iSizeH 
      Height          =   105
      Left            =   0
      MousePointer    =   7  'Size N S
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   8250
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mOpenLast 
         Caption         =   "Open &last file"
         Checked         =   -1  'True
      End
      Begin VB.Menu mLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "Pa&ste"
         Enabled         =   0   'False
         Shortcut        =   ^V
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mConvert 
         Caption         =   "&Convert File"
         Begin VB.Menu mccOpened 
            Caption         =   "&Current file"
            Enabled         =   0   'False
         End
         Begin VB.Menu mccSelect 
            Caption         =   "&Select File"
         End
      End
      Begin VB.Menu mLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mSelLang 
         Caption         =   "Select Language"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mSearch 
      Caption         =   "&Search"
      Begin VB.Menu mFind 
         Caption         =   "&Find..."
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mFindNext 
         Caption         =   "Find Ne&xt"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mLine3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mReplace 
         Caption         =   "&Replace"
         Enabled         =   0   'False
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public apiList As New clsAPI
Private viewList As New clsAPI
Private FindText1 As String
Private FindText2 As String

Private Sub Form_Load()

    Dim lngFile As String

    Me.Show

    InitDefaultLang
    SetLanguageItems
    
    iSizeV.Left = 1950
    iSizeV2.Left = 1950
    iSizeH.Top = 3180
    Form_Resize
    
    mOpenLast.Checked = CBool(GetSetting(AppName, Section, "OpenLast", False))
    lngFile = GetSetting(AppName, Section, "LastFile", "")
    
    If mOpenLast.Checked = True And lngFile <> "" Then
        
        Clear apiList.cConsts
        Clear apiList.cDeclares
        Clear apiList.cTypes

        Clear viewList.cConsts
        Clear viewList.cDeclares
        Clear viewList.cTypes

        Debug.Print lngFile

        Select Case LCase(Right(lngFile, 3))
        Case "mdb"  ' E' un database...
            Call OpenFileDatabase(lngFile, InStrRev(lngFile, "\"), apiList, Main.pBar, Main.sStatus.Panels(1))
        Case "txt"  ' E' un file di testo...
            Call OpenFileText(lngFile, InStrRev(lngFile, "\"), apiList, Main.pBar, Main.sStatus.Panels(1))
        Case Else
        
            If MsgBox(GetLangString(lng_AskForFileInfo), vbYesNo, GetLangString(lng_AskForFileInfoTitle)) = vbYes Then
                Call OpenFileText(lngFile, InStrRev(lngFile, "\"), apiList, Main.pBar, Main.sStatus.Panels(1))
            Else
                Call OpenFileDatabase(lngFile, InStrRev(lngFile, "\"), apiList, Main.pBar, Main.sStatus.Panels(1))
            End If
    
        End Select
            
        'Aggiunta ai controlli di xMain
        Dim xDec As apiDeclares
        Dim f() As String, exitLoop As Boolean
        
        'Get All the library in apiList and fill f()
        ReDim f(0)
        For Each xDec In apiList.cDeclares
            exitLoop = False
            'Search if the selected library exists in the list
            For i = LBound(f) To UBound(f)
                If xDec.decLib = f(i) Then exitLoop = True: Exit For
            Next i
            
            'Add the library if not exist (exitloop eq. false)
            If exitLoop = False Then
                ReDim Preserve f(UBound(f) + 1)
                f(UBound(f) - 1) = xDec.decLib
            End If
        Next xDec
        
        'Add Default items to the tApi control
        tApi.Nodes.Clear
        tApi.Nodes.Add , , "tDec", lng_Dec, 1, 2
        tApi.Nodes.Add , , "tConst", lng_Const, 1, 2
        tApi.Nodes.Add , , "tType", lng_Type, 1, 2
        
        'Add Default items to the tView control
        tView.Nodes.Clear
        tView.Nodes.Add , , "tDec", lng_Dec, 1, 2
        tView.Nodes.Add , , "tConst", lng_Const, 1, 2
        tView.Nodes.Add , , "tType", lng_Type, 1, 2
        
        'Add Other libraries
        For i = LBound(f) To UBound(f) - 1
            tApi.Nodes.Add "tDec", tvwChild, f(i), f(i), 1, 2
            tView.Nodes.Add "tDec", tvwChild, f(i), f(i), 1, 2
        Next i
        
        Call tApi_NodeClick(tApi.Nodes("tDec"))
        
        mccOpened.Enabled = True
    End If
    
    lngFile = GetSetting(AppName, Section, "Language", "")
    If lngFile <> "" Then Call LoadLanguage(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "languages\" & lngFile & ".lng")
    
    findResume = -1
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    iSizeH.Width = Me.Width
    iSizeV.Width = picCont.Width
    iSizeV2.Width = picCont2.Width
    picCont.Move 0, 450, Me.Width - 120, iSizeH.Top - 435
    picCont2.Move 0, iSizeH.Top + 90, Me.Width - 120, Me.Height - picCont.Height - sStatus.Height - 1220
    pBar.Move Me.Width - 1980, Me.Height - 885
On Error GoTo 0
End Sub

Private Sub iSizeH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeH.Move 0, Y + iSizeH.Top - 30, Me.Width, 75
    pSizeH.Visible = True
End Sub

Private Sub iSizeH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then pSizeH.Top = Y + iSizeH.Top - 30
End Sub

Private Sub iSizeH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeH.Top = Y + iSizeH.Top - 30
    iSizeH.Top = Y + iSizeH.Top - 30
    Form_Resize
    pSizeH.Visible = False
End Sub

Private Sub iSizeV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeV.Move X, 0, 75, Me.Height
    pSizeV.Visible = True
End Sub

Private Sub iSizeV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then pSizeV.Left = X + iSizeV.Left - 30
End Sub

Private Sub iSizeV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeV.Left = X + iSizeV.Left - 30
    iSizeV.Left = X + iSizeV.Left - 30
    picCont_Resize
    pSizeV.Visible = False
End Sub

Private Sub iSizeV2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeV2.Move X, 330, 75, Me.Height - 330
    pSizeV2.Visible = True
End Sub

Private Sub iSizeV2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then pSizeV2.Left = X + iSizeV2.Left - 30
End Sub

Private Sub iSizeV2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSizeV2.Left = X + iSizeV2.Left - 30
    iSizeV2.Left = X + iSizeV2.Left - 30
    picCont2_Resize
    pSizeV2.Visible = False
End Sub

Private Sub lApi_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lApi.SortKey = ColumnHeader.Position - 1
    lApi.SortOrder = Abs(Not (CBool(lApi.SortOrder)))
    lApi.Refresh
End Sub

Private Sub lApi_DblClick()
    Call tDown_ButtonClick(tDown.Buttons("kAdd"))
End Sub

Private Sub lApi_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    FindText1 = ""
'    Debug.Print "Item Click"
End Sub

Private Sub lApi_KeyPress(KeyAscii As Integer)
'Debug.Print "KeyPress"
'On Error Resume Next
'    Static x As ListItem
'    x.Selected = False
'    FindText1 = FindText1 & Chr(KeyAscii)
'    Debug.Print FindText1
'    Set x = lApi.FindItem(FindText1, lvwSubItem, "pName", lvwPartial)
'    x.Selected = True
'On Error GoTo 0
End Sub

Private Sub lView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lView.SortKey = ColumnHeader.Position - 1
    lView.Refresh
End Sub

Private Sub lView_DblClick()
    Call tDown_ButtonClick(tDown.Buttons("kRemove"))
End Sub

Private Sub lView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    opPriv.Value = IIf((Item.ListSubItems(1).Text = "Yes"), False, True)
    opPublic.Value = Not opPriv.Value
End Sub

Private Sub mAbout_Click()
    about.Show vbModal, Me
End Sub

Private Sub mccOpened_Click()
    Call convert.ConvertFile(False, Me)
End Sub

Private Sub mccSelect_Click()
    Call convert.ConvertFile(True, Me)
End Sub

Private Sub mClose_Click()
    Set apiList = Nothing
    Set viewList = Nothing
    tApi.Nodes.Clear
    lApi.ListItems.Clear
    lApi.ColumnHeaders.Clear
    lView.ListItems.Clear
    lView.ColumnHeaders.Clear
    Me.Caption = "ADVAPI 2"
    mccOpened.Enabled = False
End Sub

Private Sub mCopy_Click()
    Dim d As apiDeclares
    Dim c As apiConst
    Dim t As apiType
    Dim p As apiParams
    Dim clipText As String
    Dim cAdd As String
    Dim cPriv As String

    cPriv = IIf(Main.opPublic.Value, "Public ", "Private ")

    Clipboard.Clear

    clipText = ""

    'Copy constants
    For Each c In viewList.cConsts

        cAdd = cPriv & "Const " & c.decName & IIf(c.decType <> "", " As " & c.decType, "") & " = " & c.decValue

        clipText = clipText & cAdd & vbCrLf
    Next c

    clipText = clipText & vbCrLf

    'Copy Types
    For Each t In viewList.cTypes

        cAdd = cPriv & "Type " & t.decName

        For Each p In t.decParams
            cAdd = cAdd & vbCrLf & vbTab & p.paramName & IIf(p.paramType <> "", " As " & p.paramType, "")
        Next p

        cAdd = cAdd & vbCrLf & "End Type"

        clipText = clipText & cAdd & vbCrLf
    Next t

    clipText = clipText & vbCrLf

    'Copy Declares
    For Each d In viewList.cDeclares

        cAdd = cPriv & "Declare " & IIf(d.decSub, "Sub ", "Function ") & d.decName & " Lib " & Chr(34) & d.decLib & Chr(34) & _
        IIf(d.decAlias <> "", " Alias " & d.decAlias & " (", " (")

        For Each p In d.decParams
            cAdd = cAdd & p.paramName & IIf(p.paramType <> "", " As " & p.paramType & ", ", ", ")
        Next p

        If Right(cAdd, 2) = ", " Then cAdd = Left(cAdd, Len(cAdd) - 2)

        cAdd = cAdd & ")" & IIf(d.decSub, "", " As " & d.decReturnType)

        clipText = clipText & cAdd & vbCrLf
    Next d

    clipText = clipText & vbCrLf

    Clipboard.SetText clipText
End Sub

Private Sub mCut_Click()
    mCopy_Click
    Clear viewList.cConsts
    Clear viewList.cDeclares
    Clear viewList.cTypes
    Call tView_NodeClick(tView.Nodes("tDec"))
End Sub

Private Sub mExit_Click()
    mClose_Click
    End
End Sub

Private Sub mFind_Click()
    If lApi.ColumnHeaders.Count = 0 Then Exit Sub
    Search.Show vbModeless, Me
End Sub

Private Sub mFindNext_Click()

    Dim l As ListItem, StartSearch As Boolean

    StartSearch = False
    
    'findResume = 3
    
    If findResume <> -1 Then
    
    
    For Each l In Main.lApi.ListItems
        l.Selected = False
        If StartSearch Then
            For i = LBound(findList) To UBound(findList) - 1
                If findList(i) <> -1 Then
                    If Exist(l.SubItems(findList(i) - 1), findText, findJolly, IIf(findMatch, vbBinaryCompare, vbTextCompare)) Then
                        l.Selected = True
                        l.EnsureVisible
                        findResume = l.Index
                        Exit Sub
                    End If
                End If
            Next i
            
            findResume = -1
        Else
            If l.Index = findResume Then StartSearch = True
        End If
    Next l
    End If

    MsgBox lng_SearchComplete, vbOKOnly, App.Title
End Sub

Private Sub mOpen_Click()
    Dim ret As Boolean
    ret = OpenFile(Main, apiList, viewList, pBar, sStatus.Panels(1))
    
    If ret = True Then
        
        'Aggiunta ai controlli di xMain
    
        Dim xDec As apiDeclares
        Dim f() As String, exitLoop As Boolean
        
        
        'Get All the library in apiList and fill f()
        ReDim f(0)
        For Each xDec In apiList.cDeclares
            exitLoop = False
            'Search if the selected library exists in the list
            For i = LBound(f) To UBound(f)
                If xDec.decLib = f(i) Then exitLoop = True: Exit For
            Next i
            
            'Add the library if not exist (exitloop eq. false)
            If exitLoop = False Then
                ReDim Preserve f(UBound(f) + 1)
                f(UBound(f) - 1) = xDec.decLib
            End If
        Next xDec
        
        'Add Default items to the tApi control
        tApi.Nodes.Clear
        tApi.Nodes.Add , , "tDec", lng_Dec, 1, 2
        tApi.Nodes.Add , , "tConst", lng_Const, 1, 2
        tApi.Nodes.Add , , "tType", lng_Type, 1, 2
        
        'Add Default items to the tView control
        tView.Nodes.Clear
        tView.Nodes.Add , , "tDec", lng_Dec, 1, 2
        tView.Nodes.Add , , "tConst", lng_Const, 1, 2
        tView.Nodes.Add , , "tType", lng_Type, 1, 2
        
        'Add Other libraries
        For i = LBound(f) To UBound(f) - 1
            tApi.Nodes.Add "tDec", tvwChild, f(i), f(i), 1, 2
            tView.Nodes.Add "tDec", tvwChild, f(i), f(i), 1, 2
        Next i
        
        Call tApi_NodeClick(tApi.Nodes("tDec"))
        
        mccOpened.Enabled = True
        
        pBar.Value = pBar.Min
    End If
End Sub

Private Sub mOpenLast_Click()
    If mOpenLast.Checked = True Then
        mOpenLast.Checked = False
        Call SaveSetting(AppName, Section, "OpenLast", False)
    Else
        mOpenLast.Checked = True
        Call SaveSetting(AppName, Section, "OpenLast", True)
    End If
End Sub

Private Sub mSelLang_Click()
    lng.Show vbModal, Me
End Sub

Private Sub opPriv_Click()
    Dim li As ListItem
    For Each li In lView.ListItems
        li.ListSubItems(1) = lng_No
    Next li
End Sub

Private Sub opPublic_Click()
    Dim li As ListItem
    For Each li In lView.ListItems
        li.ListSubItems(1) = lng_Yes
    Next li
End Sub

Private Sub picCont_Resize()
On Error Resume Next
    tApi.Move 0, 0, iSizeV.Left + 15, picCont.Height
    lApi.Move iSizeV.Left + 75, 0, picCont.Width - iSizeV.Left - 90, picCont.Height
On Error GoTo 0
End Sub

Private Sub picCont2_Resize()
On Error Resume Next
    tView.Move 0, 330, iSizeV2.Left + 15, picCont2.Height - 300
    lView.Move iSizeV2.Left + 75, 330, picCont2.Width - iSizeV2.Left - 90, picCont2.Height - 300
On Error GoTo 0
End Sub

Private Sub tApi_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim d As apiDeclares
    Dim c As apiConst
    Dim t As apiType
    Dim li As ListItem

    Node.EnsureVisible

    'Clear All
    
    lApi.ListItems.Clear
    lApi.ColumnHeaders.Clear

    'Add shared Column Headres in the ColumnHeader collection
    lApi.ColumnHeaders.Add , , "N°", 450

    Select Case Node.Key
    Case "tDec"     'User clicked the declare node
    
        'Add Column's Header
        lApi.ColumnHeaders.Add , , "Sub", 450
        lApi.ColumnHeaders.Add , "pName", lng_Name, 1700
        lApi.ColumnHeaders.Add , , lng_Lib, 1200
        lApi.ColumnHeaders.Add , , "Alias", 1500
        lApi.ColumnHeaders.Add , , lng_Params, 3000
        lApi.ColumnHeaders.Add , , lng_ReturnType, 1200
    
        For Each d In apiList.cDeclares
            Set li = lApi.ListItems.Add(, d.idKey, String(Len(CStr(apiList.cDeclares.Count)) - Len(CStr(lApi.ListItems.Count)), "0") & CStr(lApi.ListItems.Count))
            
            li.ListSubItems.Add , , IIf(d.decSub, lng_Yes, lng_No)
            li.ListSubItems.Add , , d.decName
            li.ListSubItems.Add , , d.decLib
            li.ListSubItems.Add , , d.decAlias
            li.ListSubItems.Add , , GetParamString(d.decParams)
            li.ListSubItems.Add , , d.decReturnType
            
            li.Tag = d.idKey
        Next d
       
    Case "tConst"   'User clicked the const node
    
        'Add Column's Header
        lApi.ColumnHeaders.Add , "pName", lng_Name, 1700
        lApi.ColumnHeaders.Add , , lng_Type, 1100
        lApi.ColumnHeaders.Add , , lng_Value, 2000
    
        For Each c In apiList.cConsts
            Set li = lApi.ListItems.Add(, c.idKey, String(Len(CStr(apiList.cConsts.Count)) - Len(CStr(lApi.ListItems.Count)), "0") & CStr(lApi.ListItems.Count))

            li.ListSubItems.Add , , c.decName
            li.ListSubItems.Add , , c.decType
            li.ListSubItems.Add , , c.decValue
            
            li.Tag = c.idKey
        Next c
        
    Case "tType"    'User clicked the type node
    
        'Add Column's Header
        lApi.ColumnHeaders.Add , "pName", lng_Name, 1700
        lApi.ColumnHeaders.Add , , lng_Params, 3000
        
        For Each t In apiList.cTypes
            Set li = lApi.ListItems.Add(, t.idKey, String(Len(CStr(apiList.cTypes.Count)) - Len(CStr(lApi.ListItems.Count)), "0") & CStr(lApi.ListItems.Count))
                        
            li.ListSubItems.Add , , t.decName
            li.ListSubItems.Add , , GetParamString(t.decParams)
        
            li.Tag = t.idKey
        Next t
    
    Case Else       'User clicked a library in the declare node
    
        'Add Column's Header
        lApi.ColumnHeaders.Add , , "Sub", 450
        lApi.ColumnHeaders.Add , , lng_Name, 1700
        lApi.ColumnHeaders.Add , , lng_Lib, 1200
        lApi.ColumnHeaders.Add , , "Alias", 1500
        lApi.ColumnHeaders.Add , , lng_Params, 3000
        lApi.ColumnHeaders.Add , , lng_ReturnType, 1200
    
        For Each d In apiList.cDeclares
        
            If LCase(d.decLib) = LCase(Node.Key) Then
        
                Set li = lApi.ListItems.Add(, d.idKey, String(Len(CStr(apiList.cDeclares.Count)) - Len(CStr(lApi.ListItems.Count)), "0") & CStr(lApi.ListItems.Count))
                
                li.ListSubItems.Add , , IIf(opPriv.Value, lng_No, lng_Yes)
                li.ListSubItems.Add , , d.decName
                li.ListSubItems.Add , , d.decLib
                li.ListSubItems.Add , , d.decAlias
                li.ListSubItems.Add , , GetParamString(d.decParams)
                li.ListSubItems.Add , , d.decReturnType
                
                li.Tag = d.idKey
            End If
        Next d
    
    End Select

    lApi.Refresh
    Call tView_NodeClick(tView.Nodes(Node.Key))
    
End Sub

Private Function GetParamString(col As Collection) As String
    Dim p As apiParams, t As String
    
    t = ""
    If col.Count = 0 Then Exit Function
    For Each p In col
        t = t & p.paramName & IIf(p.paramType <> "", " As " & p.paramType, "") & ", "
    Next p
    GetParamString = Left$(t, Len(t) - 2)
End Function

Private Sub tDown_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim li As ListItem, ls As ListItem
    
    Select Case Button.Key
    Case "kAdd"
    
        If lApi.ListItems.Count = 0 Then Exit Sub
        
        Select Case Left(lApi.SelectedItem.Key, 1)
        Case "d"
            Dim dDec As apiDeclares
            
            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each dDec In viewList.cDeclares
                        If dDec.idKey = ls.Tag Then Exit Sub
                    Next dDec

                    Set dDec = apiList.cDeclares(ls.Tag)

                    'If it's not in then add it
                    viewList.cDeclares.Add dDec, ls.Tag
                End If
            Next ls
            
        Case "c"
            Dim cConst As apiConst
                        
            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each cConst In viewList.cConsts
                        If cConst.idKey = ls.Tag Then Exit Sub
                    Next cConst
                    
                    Set cConst = apiList.cConsts(ls.Tag)

                    'If it's not in then add it
                    viewList.cConsts.Add cConst, ls.Tag
                End If
            Next ls

        Case "t"
        
            Dim cType As apiType
                        
            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each cType In viewList.cTypes
                        If cType.idKey = ls.Tag Then Exit Sub
                    Next cType

                    Set cType = apiList.cTypes(ls.Tag)
                    
                    'If it's not in then add it
                    viewList.cTypes.Add cType, ls.Tag
                End If
            Next ls
        
        End Select
    Case "kRemove"
        
        For Each li In lView.ListItems
            If li.Selected Then
                Select Case Left(li.Tag, 1)
                Case "d"
                    viewList.cDeclares.Remove li.Tag
                Case "c"
                    viewList.cConsts.Remove li.Tag
                Case "t"
                    viewList.cTypes.Remove li.Tag
                End Select
            End If
        Next li
    End Select
            
    'Refresh Current View
    Select Case Left(lApi.ListItems(1).Key, 1)
    Case "c"
        Call tView_NodeClick(tView.Nodes("tConst"))
    Case "t"
        Call tView_NodeClick(tView.Nodes("tType"))
    Case Else
        Call tView_NodeClick(tView.Nodes("tDec"))
    End Select
End Sub

Private Sub tDown_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim ls As ListItem, ext As Boolean
    
    Dim dDec As apiDeclares
    Dim cConst As apiConst
    Dim cType As apiType
        
    Select Case ButtonMenu.Key
    Case "aAll" 'Add all items selected

        If lApi.ListItems.Count = 0 Then Exit Sub
        
        Select Case Left(lApi.SelectedItem.Key, 1)
        Case "d"
            
            For Each ls In lApi.ListItems

                'Check if the item it's already in list
                ext = False
                For Each dDec In viewList.cDeclares
                    If dDec.idKey = ls.Tag Then ext = True
                Next dDec

                If Not ext Then
                    Set dDec = apiList.cDeclares(ls.Tag)

                    'If it's not in then add it
                    viewList.cDeclares.Add dDec, ls.Tag
                End If
            Next ls
            
        Case "c"
                        
            For Each ls In lApi.ListItems

                'Check if the item it's already in list
                ext = False
                For Each cConst In viewList.cConsts
                    If cConst.idKey = ls.Tag Then ext = True
                Next cConst

                If Not ext Then
                    Set cConst = apiList.cConsts(ls.Tag)

                    'If it's not in then add it
                    viewList.cConsts.Add cConst, ls.Tag
                End If
            Next ls

        Case "t"
                        
            For Each ls In lApi.ListItems

                'Check if the item it's already in list
                ext = False
                For Each cType In viewList.cTypes
                    If cType.idKey = ls.Tag Then ext = True
                Next cType

                If Not ext Then
                    Set cType = apiList.cTypes(ls.Tag)

                    'If it's not in then add it
                    viewList.cTypes.Add cType, ls.Tag
                End If
            Next ls
        
        End Select


    Case "aDep"

        If lApi.ListItems.Count = 0 Then Exit Sub
        
        Select Case Left(lApi.SelectedItem.Key, 1)
        Case "d"
            
            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each dDec In viewList.cDeclares
                        If dDec.idKey = ls.Tag Then Exit Sub
                    Next dDec

                    Set dDec = apiList.cDeclares(ls.Tag)

                    'If it's not in then add it
                    viewList.cDeclares.Add dDec, ls.Tag
                    Call AddDep1(dDec, viewList, apiList)
                End If
            Next ls
            
        Case "c"
                        
            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each cConst In viewList.cConsts
                        If cConst.idKey = ls.Tag Then Exit Sub
                    Next cConst
                    
                    Set cConst = apiList.cConsts(ls.Tag)

                    'If it's not in then add it
                    viewList.cConsts.Add cConst, ls.Tag
                End If
            Next ls

        Case "t"

            For Each ls In lApi.ListItems
                If ls.Selected = True Then
                
                    'Check if the item it's already in list
                    For Each cType In viewList.cTypes
                        If cType.idKey = ls.Tag Then Exit Sub
                    Next cType

                    Set cType = apiList.cTypes(ls.Tag)
                    
                    'If it's not in then add it
                    viewList.cTypes.Add cType, ls.Tag
                End If
            Next ls
        
        End Select


    Case "rAll"
        
        If lView.ListItems.Count = 0 Then Exit Sub
        Select Case Left(lView.ListItems(1).Tag, 1)
        Case "d"
            Clear viewList.cDeclares
        Case "t"
            Clear viewList.cTypes
        Case "c"
            Clear viewList.cConsts
        End Select
    Case "rDep"
        For Each ls In lView.ListItems
            If ls.Selected Then
                Select Case Left(ls.Tag, 1)
                Case "d", "t"
                    Call RemDep(ls.Tag, viewList)
                Case "c"
                    viewList.cConsts.Remove ls.Tag
                End Select
            End If
        Next ls
    End Select
End Sub

Private Sub tUp_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "kOpen"
        mOpen_Click
    Case "kClose"
        mClose_Click
    Case "kCopy"
        mCopy_Click
    Case "kCut"
        mCut_Click
    Case "kPaste"
    Case "kFind"
        mFind_Click
    End Select
End Sub

Private Sub tView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim d As apiDeclares
    Dim c As apiConst
    Dim t As apiType
    Dim li As ListItem

    'Clear All
    
    lView.ListItems.Clear
    lView.ColumnHeaders.Clear

    'Add shared Column Headres in the ColumnHeader collection
    lView.ColumnHeaders.Add , , "N°", 450
    lView.ColumnHeaders.Add , , "Public", 450

    Select Case Node.Key
    Case "tDec"     'User clicked the declare node
    
        'Add Column's Header
        lView.ColumnHeaders.Add , , "Sub", 450
        lView.ColumnHeaders.Add , , lng_Name, 1700
        lView.ColumnHeaders.Add , , lng_Lib, 1200
        lView.ColumnHeaders.Add , , "Alias", 1500
        lView.ColumnHeaders.Add , , lng_Params, 3000
        lView.ColumnHeaders.Add , , lng_ReturnType, 1200
    
        For Each d In viewList.cDeclares
            Set li = lView.ListItems.Add(, d.idKey, String(Len(CStr(apiList.cDeclares.Count)) - Len(CStr(lView.ListItems.Count)), "0") & CStr(lView.ListItems.Count))
            
            li.ListSubItems.Add , , IIf(opPriv.Value, lng_No, lng_Yes)
            li.ListSubItems.Add , , IIf(d.decSub, lng_Yes, lng_No)
            li.ListSubItems.Add , , d.decName
            li.ListSubItems.Add , , d.decLib
            li.ListSubItems.Add , , d.decAlias
            li.ListSubItems.Add , , GetParamString(d.decParams)
            li.ListSubItems.Add , , d.decReturnType
            
            li.Tag = d.idKey
        Next d
        
        lView.Refresh
        
    Case "tConst"   'User clicked the const node
    
        'Add Column's Header
        lView.ColumnHeaders.Add , , lng_Name, 1700
        lView.ColumnHeaders.Add , , lng_Type, 1100
        lView.ColumnHeaders.Add , , lng_Value, 2000
    
        For Each c In viewList.cConsts
            Set li = lView.ListItems.Add(, c.idKey, String(Len(CStr(apiList.cConsts.Count)) - Len(CStr(lView.ListItems.Count)), "0") & CStr(lView.ListItems.Count))

            li.ListSubItems.Add , , IIf(opPriv.Value, lng_No, lng_Yes)
            li.ListSubItems.Add , , c.decName
            li.ListSubItems.Add , , c.decType
            li.ListSubItems.Add , , c.decValue
            
            li.Tag = c.idKey
        Next c
        
    Case "tType"    'User clicked the type node
    
        'Add Column's Header
        lView.ColumnHeaders.Add , , lng_Name, 1700
        lView.ColumnHeaders.Add , , lng_Params, 3000
        
        For Each t In viewList.cTypes
            Set li = lView.ListItems.Add(, t.idKey, String(Len(CStr(apiList.cTypes.Count)) - Len(CStr(lView.ListItems.Count)), "0") & CStr(lView.ListItems.Count))
                        
            li.ListSubItems.Add , , IIf(opPriv.Value, lng_No, lng_Yes)
            li.ListSubItems.Add , , t.decName
            li.ListSubItems.Add , , GetParamString(t.decParams)
        
            li.Tag = t.idKey
        Next t
    
    Case Else       'User clicked a library in the declare node
    
        'Add Column's Header
        lView.ColumnHeaders.Add , , "Sub", 450
        lView.ColumnHeaders.Add , , lng_Name, 1700
        lView.ColumnHeaders.Add , , lng_Lib, 1200
        lView.ColumnHeaders.Add , , "Alias", 1500
        lView.ColumnHeaders.Add , , lng_Params, 3000
        lView.ColumnHeaders.Add , , lng_ReturnType, 1200
    
        For Each d In viewList.cDeclares
        
            If LCase(d.decLib) = LCase(Node.Key) Then
        
                Set li = lView.ListItems.Add(, d.idKey, String(Len(CStr(apiList.cDeclares.Count)) - Len(CStr(lView.ListItems.Count)), "0") & CStr(lView.ListItems.Count))
                
                li.ListSubItems.Add , , IIf(opPriv.Value, lng_No, lng_Yes)
                li.ListSubItems.Add , , IIf(d.decSub, lng_Yes, lng_No)
                li.ListSubItems.Add , , d.decName
                li.ListSubItems.Add , , d.decLib
                li.ListSubItems.Add , , d.decAlias
                li.ListSubItems.Add , , GetParamString(d.decParams)
                li.ListSubItems.Add , , d.decReturnType
                
                li.Tag = d.idKey
            End If
        Next d
    
    End Select
    
    lView.Refresh
End Sub
