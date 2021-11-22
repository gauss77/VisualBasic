VERSION 5.00
Begin VB.Form frmTxtSP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text To Speech"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTxtSP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "Speak"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtBox 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmTxtSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSpeak_Click()
strText = txtBox.Text

Set objVoice = CreateObject("SAPI.SpVoice")

If txtBox.Text = "" Then
MsgBox "Please write something in the text area", vbExclamation, "Message"
Else
objVoice.Speak strText
End If
End Sub
