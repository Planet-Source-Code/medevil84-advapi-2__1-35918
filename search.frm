VERSION 5.00
Begin VB.Form Search 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search..."
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSearchIn 
      Caption         =   "Search In:"
      Height          =   1395
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   3375
      Begin VB.ListBox lstColumn 
         Height          =   1035
         ItemData        =   "search.frx":0000
         Left            =   120
         List            =   "search.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.CheckBox chUseJolly 
      Caption         =   "Use jolly wildcards"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin VB.CheckBox chMatchCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   660
      Width           =   2145
   End
   Begin VB.TextBox txFind 
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   180
      Width           =   4485
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   345
      Left            =   4860
      TabIndex        =   1
      Top             =   1710
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3810
      TabIndex        =   0
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label lbFind 
      Alignment       =   1  'Right Justify
      Caption         =   "Find:"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
    If txFind.Text = "" Then Exit Sub
    
    Dim l As ListItem, i As Long
    
    findJolly = IIf(chUseJolly.Value = vbUnchecked, False, True)
    findMatch = IIf(chMatchCase.Value = vbUnchecked, False, True)
    findText = txFind.Text
    findResume = -1
    ReDim findList(Main.lApi.ColumnHeaders.Count)
    
    For i = 0 To lstColumn.ListCount - 1
        findList(i) = IIf(lstColumn.Selected(i), lstColumn.ItemData(i), -1)
    Next i
    
    For Each l In Main.lApi.ListItems
        For i = 0 To lstColumn.ListCount - 1
            If lstColumn.Selected(i) Then
                If Exist(l.SubItems(i), txFind, chUseJolly.Value, IIf(chMatchCase.Value, vbBinaryCompare, vbTextCompare)) Then
                    l.Selected = True
                    l.EnsureVisible
                    findResume = l.Index
                    Me.Hide
                    Exit Sub
                End If
            End If
        Next i
    Next l
    MsgBox lng_NoItems, vbOKOnly, App.Title
    Me.Hide
End Sub

Private Sub Form_Load()

    Me.Caption = Replace(Main.mSearch.Caption, "&", "") & "..."
    
    Dim ch As ColumnHeader
    Dim i As Long
    
    For i = 1 To Main.lApi.ColumnHeaders.Count
        Set ch = Main.lApi.ColumnHeaders.Item(i)
        lstColumn.AddItem ch.Text
        lstColumn.ItemData(i - 1) = ch.Index
    Next i
End Sub
