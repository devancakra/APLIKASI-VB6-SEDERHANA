VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF80FF&
   Caption         =   "Program Option Button Auto Output"
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
      Caption         =   "[ Program Option Button Auto Output ]"
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
      Top             =   2400
      Width           =   7455
      Begin VB.Frame Frame3 
         Caption         =   "Minuman"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3720
         TabIndex        =   7
         Top             =   480
         Width           =   3375
         Begin VB.OptionButton Option6 
            Caption         =   "Es Teh"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1335
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Es Jeruk"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Teh Anget"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Makanan"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton Option3 
            Caption         =   "Mie Ayam"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Gado-Gado"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nasi Goreng"
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Minuman Yang Dipilih"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Makanan Yang Dipilih"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1920
         TabIndex        =   3
         Top             =   3480
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1920
         TabIndex        =   1
         Top             =   2520
         Width           =   5175
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
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   7200
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************MAKANAN******************'
Private Sub Option1_Click()
    Label3.Caption = Option1.Caption
End Sub

Private Sub Option2_Click()
    Label3.Caption = Option2.Caption
End Sub

Private Sub Option3_Click()
    Label3.Caption = Option3.Caption
End Sub

'*****************MINUMAN******************'
Private Sub Option4_Click()
    Label5.Caption = Option4.Caption
End Sub

Private Sub Option5_Click()
    Label5.Caption = Option5.Caption
End Sub

Private Sub Option6_Click()
    Label5.Caption = Option6.Caption
End Sub
