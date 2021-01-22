VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Program Validasi"
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
      Caption         =   "[ Program Validasi ]"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6000
      TabIndex        =   0
      Top             =   2280
      Width           =   7455
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   6975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Validasi"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Nama Lengkap Anda"
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
    If Text1.Text = "Devan Cakra Mudra Wijaya" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "DEVAN CAKRA MUDRA WIJAYA" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "devan cakra mudra wijaya" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "DEVAN" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "Devan" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "devan" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "DEVAN CAKRA" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "devan cakra" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "Devan Cakra" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "DEVAN CAKRA MUDRA" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "Devan Cakra Mudra" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    ElseIf Text1.Text = "devan cakra mudra" Then
        MsgBox "Anda Mahasiswa UPN V Jatim", vbInformation
    Else
        MsgBox "Anda Bukan Mahasiswa UPN V Jatim / Tidak Terdaftar Di Sistem", vbCritical
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub
