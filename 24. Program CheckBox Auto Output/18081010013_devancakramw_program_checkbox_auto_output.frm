VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Program CheckBox Auto Output"
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
      Caption         =   "[ Program CheckBox Auto Output ]"
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
         Begin VB.CheckBox Check6 
            Caption         =   " Es Teh Jumbo"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Check5 
            Caption         =   " Ocean Blue"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox Check4 
            Caption         =   " Milshake"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1095
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
         Begin VB.CheckBox Check3 
            Caption         =   " Rica-Rica"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   " Steak"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   " Soto Ayam"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1335
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
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         TabIndex        =   3
         Top             =   3480
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
Private Sub Check1_Click()
    Label3.Caption = Check1.Caption
End Sub

Private Sub Check2_Click()
    Label3.Caption = Check2.Caption
End Sub

Private Sub Check3_Click()
    Label3.Caption = Check3.Caption
End Sub


'*****************MINUMAN******************'
Private Sub Check4_Click()
    Label5.Caption = Check4.Caption
End Sub

Private Sub Check5_Click()
    Label5.Caption = Check5.Caption
End Sub

Private Sub Check6_Click()
    Label5.Caption = Check6.Caption
End Sub
