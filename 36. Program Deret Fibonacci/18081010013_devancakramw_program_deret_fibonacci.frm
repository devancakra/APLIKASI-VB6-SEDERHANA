VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Program Deret Fibonacci"
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
      Caption         =   "[ Program Deret Fibonacci ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   6000
      TabIndex        =   0
      Top             =   2280
      Width           =   7455
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   5655
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
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
         Height          =   855
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Clear"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Submit"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   1215
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Banyak Deret :"
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
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Deret 2 :"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deret 1 :"
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
         TabIndex        =   2
         Top             =   600
         Width           =   3975
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5880
      TabIndex        =   3
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
    Dim a, b, c, n As Integer
    a = Text1.Text
    b = Text2.Text
    n = Text3.Text
    
    List1.AddItem a
    List1.AddItem b
    
    For i = 0 To n
        c = Val(a) + Val(b)
        List1.AddItem c
        a = b
        b = c
    Next i
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    List1.Clear
End Sub

Private Sub Command3_Click()
    End
End Sub
