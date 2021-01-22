VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Program Input Nama Submit"
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
      Caption         =   "[ Program Input Nama Submit ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   6000
      TabIndex        =   0
      Top             =   2640
      Width           =   7455
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Submit"
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   6735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Input :"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Nama Anda"
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
         TabIndex        =   2
         Top             =   840
         Width           =   2295
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
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   5880
      TabIndex        =   6
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
    Label3.Caption = Text1.Text
End Sub
