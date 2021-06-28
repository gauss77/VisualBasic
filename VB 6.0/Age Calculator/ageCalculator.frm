VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ageCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Age Calculator"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton btnClear 
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
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame frame1 
      Height          =   3015
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Frame frame2 
         Height          =   615
         Left            =   1080
         TabIndex        =   9
         Top             =   0
         Width           =   2055
         Begin VB.Label lblAge 
            Alignment       =   2  'Center
            Caption         =   "Your Age"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label lblDay 
         Caption         =   "Day (s)"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblMonth 
         Caption         =   "Month (s)"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblYear 
         Caption         =   "Year (s)"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame frame3 
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5775
      Begin MSComCtl2.DTPicker MonthView 
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16449537
         UpDown          =   -1  'True
         CurrentDate     =   44375
      End
      Begin VB.Label lblMDY 
         Caption         =   "(mm/dd/yy)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblBirthDay 
         Caption         =   "Pick Your Birthday"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "ageCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cd, age As Variant
Dim d, m, mm, y As Variant


Private Sub btnCalculate_Click()
cd = Now
d = MonthView.Day
m = MonthView.Month
y = MonthView.Year
age = Month(cd) - m
If age < 0 Then
age = 1
Else
age = 0
End If



mm = Month(cd) - m
If mm < 0 Then
mm = 12
Else
mm = 0
End If

If MonthView.Value > cd Then
MsgBox "Birth date should be earlier than current date", vbExclamation + vbOKOnly, "Error"
txtYear.Text = ""
txtMonth.Text = ""
txtDate.Text = ""
ElseIf MonthView.Value < cd Then
txtYear.Text = Abs((Year(cd) - y) - age)
txtMonth.Text = Abs((Month(cd) - m) + mm)
txtDate.Text = Abs(Day(cd) - d)
End If
End Sub

Private Sub btnClear_Click()
txtYear.Text = ""
txtMonth.Text = ""
txtDate.Text = ""
End Sub


Private Sub txtYear_Change()
If txtYear.Text <> "" Then
btnClear.Enabled = True
Else
btnClear.Enabled = False
End If
End Sub


