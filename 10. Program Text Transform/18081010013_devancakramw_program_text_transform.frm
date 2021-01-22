VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Program Text Transform"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   16725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "[ Program Text Transform ]"
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
      Top             =   2040
      Width           =   7455
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Quit"
         Height          =   615
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Submit"
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pilihan"
         Height          =   3735
         Left            =   4920
         TabIndex        =   4
         Top             =   480
         Width           =   2175
         Begin VB.CheckBox Check2 
            Caption         =   "Tebal"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3120
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Miring"
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   2160
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Merah"
            Height          =   400
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Biru"
            Height          =   400
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   360
         TabIndex        =   9
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Result :"
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
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Input"
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
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   6840
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************FONT ITALIC************'
Private Sub Check1_Click()
    Label3.FontItalic = Check1.Value
End Sub

'**************FONT BOLD************'
Private Sub Check2_Click()
    Label3.FontBold = Check2.Value
End Sub

'**************FONT COLOR BLUE************'
Private Sub Option1_Click()
    Label3.ForeColor = vbBlue
End Sub

'**************FONT COLOR RED************'
Private Sub Option2_Click()
    Label3.ForeColor = vbRed
End Sub

'**************FONT SUBMIT************'
Private Sub Command1_Click()
    Label3.Caption = Text1.Text
End Sub

'**************EXIT************'
Private Sub Command2_Click()
    End
End Sub
