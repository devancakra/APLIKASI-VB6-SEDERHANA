VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Program ListBox Quick Search"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   16725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "[ Program ListBox Quick Search ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6000
      TabIndex        =   0
      Top             =   2760
      Width           =   7455
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF80FF&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   5295
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Yang Dicari :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   6015
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BY : DEVAN CAKRA MUDRA WIJAYA (18081010013)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   7440
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
    List1.ListIndex = i
    If List1.Text = Text1.Text Then
        MsgBox " Ditemukan pada index - " & i, vbInformation + vbOKOnly, "founded"
        Exit Sub
    End If
    Next i
    MsgBox " Tidak ditemukan", vbExclamation + vbOKOnly, "not founded"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 9
        List1.AddItem "ListItem" & i
    Next i
End Sub

Private Sub Command2_Click()
    End
End Sub
