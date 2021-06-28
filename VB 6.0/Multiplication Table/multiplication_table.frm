VERSION 5.00
Begin VB.Form multiplication_table 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiplication Table"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblResult 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton btnClr 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtNum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame frmTable 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.Label lblNum 
         Caption         =   "Table NUM"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "multiplication_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y As Variant


Private Sub btnCalculate_Click()

x = txtNum.Text
y = x
If x = "" Then
MsgBox "Please Enter a Number", vbExclamation + vbOKOnly, "Msgbox"
ElseIf IsNumeric(x) = False Then
MsgBox "Please Enter Numbers Only", vbExclamation + vbOKOnly, "Msgbox"
txtNum.Text = ""
ElseIf IsNumeric(x) = True And Len(x) <= 5 Then

For x = 1 To 10
If x = 1 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 2 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 3 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 4 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 5 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 6 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 7 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 8 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 9 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(2) & "=" & Space$(2) & (y * x) & vbCrLf
ElseIf x = 10 Then
lblResult.Text = lblResult.Text & y & Space$(2) & "x" & Space$(2) & x & Space$(1) & "=" & Space$(2) & (y * x) & vbCrLf
End If
Next
btnCalculate.Enabled = False
End If

If IsNumeric(x) = True And Len(x) > 5 Then
MsgBox "Please Enter Upto 5 Numbers", vbExclamation + vbOKOnly, "Msgbox"
txtNum.Text = ""
End If
End Sub



Private Sub btnClr_Click()
txtNum.Text = ""
lblResult.Text = ""
btnCalculate.Enabled = True
End Sub


Private Sub txtNum_Change()
If txtNum.Text <> "" Then
btnClr.Enabled = True

Else
btnClr.Enabled = False
btnCalculate.Enabled = True
End If
End Sub
