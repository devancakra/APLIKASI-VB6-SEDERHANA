VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "18081010013_devancakramw_program_combobox_auto_output_sederhana.frx":0000
         Left            =   360
         List            =   "18081010013_devancakramw_program_combobox_auto_output_sederhana.frx":0002
         TabIndex        =   4
         Text            =   "Pilih Menu Makanan"
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Makanan Yang Dipilih :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   360
         TabIndex        =   1
         Top             =   2640
         Width           =   6735
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
Private Sub Combo1_Click()
    If Combo1.Text = "Nasi Goreng" Then
        Label3.Caption = "Nasi Goreng"
    ElseIf Combo1.Text = "Soto Ayam" Then
        Label3.Caption = "Soto Ayam"
    ElseIf Combo1.Text = "Soto Daging" Then
        Label3.Caption = "Soto Daging"
    ElseIf Combo1.Text = "Nasi Pecel" Then
        Label3.Caption = "Nasi Pecel"
    ElseIf Combo1.Text = "Nasi Penyetan" Then
        Label3.Caption = "Nasi Penyetan"
    ElseIf Combo1.Text = "Nasi Geprek" Then
        Label3.Caption = "Nasi Geprek"
    ElseIf Combo1.Text = "Nasi Bebek" Then
        Label3.Caption = "Nasi Bebek"
    ElseIf Combo1.Text = "Mie Ayam" Then
        Label3.Caption = "Mie Ayam"
    ElseIf Combo1.Text = "Gado-Gado" Then
        Label3.Caption = "Gado-Gado"
    ElseIf Combo1.Text = "Sayur Bening" Then
        Label3.Caption = "Sayur Bening"
    Else
        Label3.Caption = "Tidak Tersedia!!"
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Nasi Goreng"
    Combo1.AddItem "Soto Ayam"
    Combo1.AddItem "Soto Daging"
    Combo1.AddItem "Nasi Pecel"
    Combo1.AddItem "Nasi Penyetan"
    Combo1.AddItem "Nasi Geprek"
    Combo1.AddItem "Nasi Bebek"
    Combo1.AddItem "Mie Ayam"
    Combo1.AddItem "Gado-Gado"
    Combo1.AddItem "Sayur Bening"
End Sub

