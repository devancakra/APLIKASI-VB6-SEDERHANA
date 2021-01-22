VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Program 4 Operasi Sederhana"
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
      Caption         =   "[ Program 4 Operasi Sederhana ]"
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
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FF00&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4200
         Width           =   6735
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   2640
         TabIndex        =   7
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan bilangan 2"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2160
         TabIndex        =   4
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Operasi :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan bilangan 1"
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
      TabIndex        =   5
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
    Dim bil1 As Integer
    Dim bil2 As Integer
    Dim result As Integer
    bil1 = Text1.Text
    bil2 = Text2.Text
    result = Val(bil1) + Val(bil2)
    Label3.Caption = result
End Sub

Private Sub Command2_Click()
    Dim bil1 As Integer
    Dim bil2 As Integer
    Dim result As Integer
    bil1 = Text1.Text
    bil2 = Text2.Text
    result = Val(bil1) - Val(bil2)
    Label3.Caption = result
End Sub

Private Sub Command3_Click()
    Dim bil1 As Integer
    Dim bil2 As Integer
    Dim result As Integer
    bil1 = Text1.Text
    bil2 = Text2.Text
    result = Val(bil1) / Val(bil2)
    Label3.Caption = result
End Sub

Private Sub Command4_Click()
    Dim bil1 As Integer
    Dim bil2 As Integer
    Dim result4 As Integer
    bil1 = Text1.Text
    bil2 = Text2.Text
    result = Val(bil1) * Val(bil2)
    Label3.Caption = result
End Sub

Private Sub Command5_Click()
    Text1.Text = ""
    Text2.Text = ""
    Label3.Caption = ""
End Sub
