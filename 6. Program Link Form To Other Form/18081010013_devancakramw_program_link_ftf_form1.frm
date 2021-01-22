VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Program Link Form To Other Form"
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
      Caption         =   "[ Program Link Form To Other Form ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6480
      TabIndex        =   0
      Top             =   3120
      Width           =   6375
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Quit"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "HALAMAN 2"
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "HALAMAN 1"
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Klik tombol dibawah ini untuk pindah ke halaman lain"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   5880
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    End
End Sub
