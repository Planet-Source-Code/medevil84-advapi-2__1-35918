VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About..."
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   810
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "about.frx":0000
      Top             =   150
      Width           =   4125
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      MouseIcon       =   "about.frx":0036
      MousePointer    =   99  'Custom
      Picture         =   "about.frx":0478
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
   Begin VB.Label lbMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "medevil@inwind.it"
      Height          =   195
      Left            =   2190
      MouseIcon       =   "about.frx":08BA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   690
      Width           =   1290
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = Replace(Main.mAbout.Caption, "&", "") & "..."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lbMail.ForeColor = vbBlack Then
        lbMail.Font.Underline = False
        lbMail.ForeColor = vbBlack
    End If
End Sub

Private Sub lbMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lbMail.ForeColor = vbBlue Then
        lbMail.Font.Underline = True
        lbMail.ForeColor = vbBlue
    End If
End Sub

Private Sub picIcon_Click()
    Dim sText As String
    Dim exitLoop As Boolean
    Do
        sText = InputBox("Say: " & Chr(34) & "Thanks" & Chr(34), "Please Write...")
        If InStr(1, sText, "thank", vbTextCompare) Then
            MsgBox "You're a Good Boy!", vbOKOnly + vbInformation, "You're welcome"
            exitLoop = True
        Else
            MsgBox "You're a Bad Boy!" & vbCrLf & "I'll give you another chance!", vbOKOnly + vbCritical, "Did you write " & Chr(34) & "Thanks" & Chr(34) & " ?"
            exitLoop = False
        End If
    Loop While exitLoop = False
    Me.Hide
End Sub
