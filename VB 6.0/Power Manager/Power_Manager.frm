VERSION 5.00
Begin VB.Form Power_Manager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Manager"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Shutdown 
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton btn_Restart 
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton btn_LogOff 
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label winXP 
      Alignment       =   2  'Center
      Caption         =   "Windows XP Only"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image imgComputer 
      Height          =   4815
      Left            =   120
      Picture         =   "Power_Manager.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5400
   End
End
Attribute VB_Name = "Power_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_LogOff_Click()
Shell ("Shutdown -l")
End Sub

Private Sub btn_Restart_Click()
Shell ("Shutdown -r -t 0")
End Sub

Private Sub btn_Shutdown_Click()
Shell ("Shutdown -s -t 0")
End Sub
