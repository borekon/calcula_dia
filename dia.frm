VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcula el dia"
   ClientHeight    =   1755
   ClientLeft      =   3675
   ClientTop       =   3255
   ClientWidth     =   4680
   Icon            =   "dia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "Sobre.."
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox año 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox mes 
      Height          =   315
      ItemData        =   "dia.frx":030A
      Left            =   1560
      List            =   "dia.frx":0332
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox dia 
      Height          =   315
      ItemData        =   "dia.frx":039B
      Left            =   240
      List            =   "dia.frx":03FC
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim diaref As Single
Dim mesref As Single
Dim anyref As Date
Dim n As Double
Dim fecha As String
Private Sub Command1_Click()
  anyref = 1996
  diaref = 1
  mesref = 1
  n = 1
  If (año.Text < anyref) Then
    While (año.Text < anyref)
      anyref = anyref - 4
      n = (n) + 2
    Wend
    End If
    If (año.Text > anyref) Then
      While (año.Text >= anyref + 4)
    anyref = anyref + 4
      n = n + 5
  Wend
  End If
  If (año.Text - anyref = 1) Then n = n + 2
  If (año.Text - anyref = 2) Then n = n + 3
  If (año.Text - anyref = 3) Then n = n + 4
      
    
  If (año.Text - anyref = 0) Then
    If (mes.ListIndex = 0) Then n = n + dia.Text - 1
    If (mes.ListIndex = 1) Then n = n + 3 + dia.Text - 1
    If (mes.ListIndex = 2) Then n = n + dia.Text + 4 - 1
    If (mes.ListIndex = 3) Then n = n + dia.Text - 1
    If (mes.ListIndex = 4) Then n = n + dia.Text + 1
    If (mes.ListIndex = 5) Then n = n + dia.Text + 4
    If (mes.ListIndex = 6) Then n = n + dia.Text - 1
    If (mes.ListIndex = 7) Then n = n + dia.Text + 2
    If (mes.ListIndex = 8) Then n = n + dia.Text + 5
    If (mes.ListIndex = 9) Then n = n + dia.Text
    If (mes.ListIndex = 10) Then n = n + dia.Text + 3
    If (mes.ListIndex = 11) Then n = n + dia.Text + 5
    ElseIf (año.Text - anyref <> 0) Then
     If (mes.ListIndex = 0) Then n = n + dia.Text - 1
     If (mes.ListIndex = 1) Then n = n + 3 + dia.Text - 1
     If (mes.ListIndex = 2) Then n = n + dia.Text + 2
     If (mes.ListIndex = 3) Then n = n + dia.Text - 5
     If (mes.ListIndex = 4) Then n = n + dia.Text
     If (mes.ListIndex = 5) Then n = n + dia.Text + 3
     If (mes.ListIndex = 6) Then n = n + dia.Text + 5
     If (mes.ListIndex = 7) Then n = n + dia.Text + 1
     If (mes.ListIndex = 8) Then n = n + dia.Text + 4
     If (mes.ListIndex = 9) Then n = n + dia.Text - 1
     If (mes.ListIndex = 10) Then n = n + dia.Text + 2
     If (mes.ListIndex = 11) Then n = n + dia.Text + 4
  End If
  While (n > 7)
    n = n - 7
  Wend
   If (n = 1) Then MsgBox "Lunes", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 2) Then MsgBox "Martes", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 3) Then MsgBox "Miercoles", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 4) Then MsgBox "Jueves", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 5) Then MsgBox "Viernes", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 6) Then MsgBox "Sabado", , (dia.Text & " " & mes.Text & " " & año.Text)
   If (n = 7) Then MsgBox "Domingo", , (dia.Text & " " & mes.Text & " " & año.Text)
anyref = 1996
n = 1
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
frmAbout.Show
End Sub
Private Sub Form_Load()
dia.ListIndex = 0
mes.ListIndex = 0
año.Text = 1996
End Sub
