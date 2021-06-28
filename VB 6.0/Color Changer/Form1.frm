VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Changer"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form1.frx":0000
      Left            =   720
      List            =   "Form1.frx":0019
      TabIndex        =   1
      Text            =   "Blank"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pick Your Color"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()

If Combo1.Text = "RED" Then
Form1.BackColor = &H8080FF
Label1.BackColor = &H8080FF

ElseIf Combo1.Text = "GREEN" Then
Form1.BackColor = &H80FF80
Label1.BackColor = &H80FF80

ElseIf Combo1.Text = "YELLOW" Then
Form1.BackColor = &H80FFFF
Label1.BackColor = &H80FFFF

ElseIf Combo1.Text = "BLUE" Then
Form1.BackColor = &HFF8080
Label1.BackColor = &HFF8080

ElseIf Combo1.Text = "GREY" Then
Form1.BackColor = &HE0E0E0
Label1.BackColor = &HE0E0E0

ElseIf Combo1.Text = "WHITE" Then
Form1.BackColor = &HFFFFFF
Label1.BackColor = &HFFFFFF


ElseIf Combo1.Text = "Blank" Then
Form1.BackColor = &H8000000F
Label1.BackColor = &H8000000F


End If

End Sub
