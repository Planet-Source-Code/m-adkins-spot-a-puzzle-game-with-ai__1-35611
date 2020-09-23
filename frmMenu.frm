VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Game Menu"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start Game"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3480
      TabIndex        =   23
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4680
      TabIndex        =   22
      Top             =   3300
      Width           =   1215
   End
   Begin VB.ComboBox cboPlayer4 
      Height          =   315
      ItemData        =   "frmMenu.frx":0442
      Left            =   4680
      List            =   "frmMenu.frx":044C
      TabIndex        =   21
      Text            =   "Computer"
      Top             =   2850
      Width           =   1215
   End
   Begin VB.ComboBox cboPlayer3 
      Height          =   315
      ItemData        =   "frmMenu.frx":0461
      Left            =   4680
      List            =   "frmMenu.frx":046B
      TabIndex        =   20
      Text            =   "Computer"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboPlayer2 
      Height          =   315
      ItemData        =   "frmMenu.frx":0480
      Left            =   4680
      List            =   "frmMenu.frx":048A
      TabIndex        =   19
      Text            =   "Computer"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.ComboBox cboPlayer1 
      Height          =   315
      ItemData        =   "frmMenu.frx":049F
      Left            =   4680
      List            =   "frmMenu.frx":04A9
      TabIndex        =   18
      Text            =   "Human"
      Top             =   1500
      Width           =   1215
   End
   Begin VB.ComboBox cboColor4 
      Height          =   315
      ItemData        =   "frmMenu.frx":04BE
      Left            =   3300
      List            =   "frmMenu.frx":04D1
      TabIndex        =   17
      Text            =   "Yellow"
      Top             =   2850
      Width           =   1215
   End
   Begin VB.ComboBox cboColor3 
      Height          =   315
      ItemData        =   "frmMenu.frx":04F6
      Left            =   3300
      List            =   "frmMenu.frx":0509
      TabIndex        =   14
      Text            =   "Green"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboColor2 
      Height          =   315
      ItemData        =   "frmMenu.frx":052E
      Left            =   3300
      List            =   "frmMenu.frx":0541
      TabIndex        =   11
      Text            =   "Red"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.ComboBox cboColor1 
      Height          =   315
      ItemData        =   "frmMenu.frx":0566
      Left            =   3300
      List            =   "frmMenu.frx":0579
      TabIndex        =   8
      Text            =   "Blue"
      Top             =   1500
      Width           =   1215
   End
   Begin VB.TextBox txtPlayer4 
      Height          =   315
      Left            =   1650
      TabIndex        =   16
      Text            =   "Player4"
      Top             =   2850
      Width           =   1515
   End
   Begin VB.TextBox txtPlayer3 
      Height          =   315
      Left            =   1650
      TabIndex        =   13
      Text            =   "Player3"
      Top             =   2400
      Width           =   1515
   End
   Begin VB.TextBox txtPlayer2 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Text            =   "Player2"
      Top             =   1950
      Width           =   1515
   End
   Begin VB.TextBox txtPlayer1 
      Height          =   315
      Left            =   1650
      TabIndex        =   7
      Text            =   "Player1"
      Top             =   1500
      Width           =   1515
   End
   Begin VB.ComboBox cboY 
      Height          =   315
      ItemData        =   "frmMenu.frx":059E
      Left            =   1650
      List            =   "frmMenu.frx":05C0
      TabIndex        =   5
      Text            =   "8"
      Top             =   1050
      Width           =   765
   End
   Begin VB.ComboBox cboX 
      Height          =   315
      ItemData        =   "frmMenu.frx":05E5
      Left            =   1650
      List            =   "frmMenu.frx":0607
      TabIndex        =   3
      Text            =   "8"
      Top             =   600
      Width           =   765
   End
   Begin VB.ComboBox cboPlayers 
      Height          =   315
      ItemData        =   "frmMenu.frx":062C
      Left            =   1650
      List            =   "frmMenu.frx":0639
      TabIndex        =   1
      Text            =   "2"
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lblPlayer4 
      Caption         =   "Player &4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   15
      Top             =   2850
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "Player &2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label lblPlayer3 
      Caption         =   "Player &3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Player &1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   1500
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "No of &Columns:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "No of &Rows:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "&No of Players:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboColor1_LostFocus()
    If cboColor1 <> "Blue" And cboColor1 <> "Red" And cboColor1 <> "Green" _
            And cboColor1 <> "Yellow" And cboColor1 <> "Happy" Then
        cboColor1.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboColor2_LostFocus()
    If cboColor2 <> "Blue" And cboColor2 <> "Red" And cboColor2 <> "Green" _
            And cboColor2 <> "Yellow" And cboColor2 <> "Happy" Then
        cboColor2.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboColor3_LostFocus()
    If cboColor3 <> "Blue" And cboColor3 <> "Red" And cboColor3 <> "Green" _
            And cboColor3 <> "Yellow" And cboColor3 <> "Happy" Then
        cboColor3.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboColor4_LostFocus()
    If cboColor4 <> "Blue" And cboColor4 <> "Red" And cboColor4 <> "Green" _
            And cboColor4 <> "Yellow" And cboColor4 <> "Happy" Then
        cboColor4.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboPlayer1_LostFocus()
    If cboPlayer1 <> "Human" And cboPlayer1 <> "Computer" Then
        cboPlayer1.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboPlayer2_LostFocus()
    If cboPlayer2 <> "Human" And cboPlayer2 <> "Computer" Then
        cboPlayer2.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboPlayer3_LostFocus()
    If cboPlayer3 <> "Human" And cboPlayer3 <> "Computer" Then
        cboPlayer3.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboPlayer4_LostFocus()
    If cboPlayer4 <> "Human" And cboPlayer4 <> "Computer" Then
        cboPlayer4.SetFocus
        MsgBox "Please select an item from the list.", vbExclamation, _
                "Invalid Selection"
    End If
End Sub

Private Sub cboPlayers_Validate(Cancel As Boolean)
    Select Case cboPlayers
        Case 2
'            lblPlayer3.Enabled = False
'            lblPlayer4.Enabled = False
'            txtPlayer3.Enabled = False
'            txtPlayer4.Enabled = False
'            cboColor3.Enabled = False
'            cboColor4.Enabled = False
'            cboPlayer3.Enabled = False
'            cboPlayer4.Enabled = False
        Case 3
'            lblPlayer3.Enabled = True
'            lblPlayer4.Enabled = False
'            txtPlayer3.Enabled = True
'            txtPlayer4.Enabled = False
'            cboColor3.Enabled = True
'            cboColor4.Enabled = False
'            cboPlayer3.Enabled = True
'            cboPlayer4.Enabled = False
        Case 4
'            lblPlayer3.Enabled = True
'            lblPlayer4.Enabled = True
'            txtPlayer3.Enabled = True
'            txtPlayer4.Enabled = True
'            cboColor3.Enabled = True
'            cboColor4.Enabled = True
'            cboPlayer3.Enabled = True
'            cboPlayer4.Enabled = True
        Case Else
            cboPlayers.SetFocus
            MsgBox "Please select an item from the list.", vbExclamation, _
                    "Invalid Selection"
    End Select
End Sub

Private Sub cboX_LostFocus()
    If IsNumeric(cboX) Then
        cboX = CLng(cboX)
        If cboX >= 3 And cboX <= 12 Then Exit Sub
    End If
    cboX.SetFocus
    MsgBox "Please select an item from the list.", vbExclamation, _
            "Invalid Selection"
End Sub

Private Sub cboY_LostFocus()
    If IsNumeric(cboY) Then
        cboY = CLng(cboY)
        If cboY >= 3 And cboY <= 12 Then Exit Sub
    End If
    cboY.SetFocus
    MsgBox "Please select an item from the list.", vbExclamation, _
            "Invalid Selection"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Set ggamGame = New CGame
    With ggamGame
        .MaxX = CByte(cboX)
        .MaxY = CByte(cboY)
        .Players = CByte(cboPlayers)
        gstrPlayerNames(1) = txtPlayer1
        gstrPlayerImg(1) = cboColor1
        gstrPlayerNames(2) = txtPlayer2
        gstrPlayerImg(2) = cboColor2
        If .Players > 2 Then
            gstrPlayerNames(3) = txtPlayer3
            gstrPlayerImg(3) = cboColor3
        End If
        If .Players > 3 Then
            gstrPlayerNames(4) = txtPlayer4
            gstrPlayerImg(4) = cboColor4
        End If
        .IsHuman(1) = (cboPlayer1 = "Human")
        .IsHuman(2) = (cboPlayer2 = "Human")
        .IsHuman(3) = (cboPlayer3 = "Human")
        .IsHuman(4) = (cboPlayer4 = "Human")
    End With
    Me.Hide
    frmGrid.Show
End Sub

Private Sub Form_Terminate()
    Set ggamGame = Nothing
End Sub

Private Sub txtPlayer1_GotFocus()
    With txtPlayer1
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPlayer2_GotFocus()
    With txtPlayer2
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPlayer3_GotFocus()
    With txtPlayer3
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPlayer4_GotFocus()
    With txtPlayer4
        .SelLength = Len(.Text)
    End With
End Sub
