VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Program ListBox Auto Search"
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
      Caption         =   "[ Program ListBox Auto Search ]"
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
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   5295
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   5295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
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
Private Sub Form_Load()
    List1.AddItem "Adit"
    List1.AddItem "Devan"
    List1.AddItem "Heri"
    List1.AddItem "Adinda"
    List1.AddItem "Helna"
    List1.AddItem "Ade"
    List1.AddItem "Putri"
    List1.AddItem "Aldo"
    List1.AddItem "Miranda"
    List1.AddItem "Keysha"
    List1.AddItem "Nurmala"
    List1.AddItem "Tian"
    List1.AddItem "Agung"
    List1.AddItem "Rini"
    List1.AddItem "Alyssa"
    List1.AddItem "Maya"
    List1.AddItem "Della"
    List1.AddItem "Wahyu"
    List1.AddItem "Putu"
    List1.AddItem "Gedhe"
    List1.AddItem "Sunu"
    List1.AddItem "Chilla"
    List1.AddItem "Dhanny"
    List1.AddItem "Alfareza"
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Text1_Change()
    List1.ListIndex = sendmessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
End Sub
