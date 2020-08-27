VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CALCULATOR By:Carmina"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdE 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H0000C000&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdS 
      BackColor       =   &H0000C000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdM 
      BackColor       =   &H0000C000&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H0000C000&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOff 
      BackColor       =   &H0000C000&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdP 
      BackColor       =   &H0000C000&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmd0 
      BackColor       =   &H0000C000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdAS 
      BackColor       =   &H0000C000&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H0000C000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H0000C000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H0000C000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H0000C000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H0000C000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H0000C000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H0000C000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H0000C000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H0000C000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Computer As Double
Dim Science As Double
Dim ComSci As Single
Dim Network As Single
Dim Keyboard  As Single
Dim Minah As Single
Dim InStr As String
Dim Operators As String

Private Sub cmd0_Click()
txtOutput.Text = "0"
End Sub

Private Sub cmd1_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "1"
End Sub

Private Sub cmd2_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "2"

End Sub

Private Sub cmd3_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "3"

End Sub

Private Sub cmd4_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "4"

End Sub

Private Sub cmd5_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "5"

End Sub

Private Sub cmd6_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "6"

End Sub

Private Sub cmd7_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "7"

End Sub

Private Sub cmd8_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "8"

End Sub

Private Sub cmd9_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
txtOutput.Text = txtOutput.Text & "9"

End Sub

Private Sub cmdA_Click()
Computer = Val(txtOutput.Text)
txtOutput.Text = ""
Operators = "+"
End Sub

Private Sub cmdAS_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
 If InStr(txtOutput.Text, "-") = 0 Then
txtOutput.Text = txtOutput.Text & "-"
  End If
End Sub

Private Sub cmdClear_Click()
txtOutput.Text = "0"
End Sub

Private Sub cmdD_Click()
Computer = Val(txtOutput.Text)
txtOutput.Text = ""
Operators = "/"
End Sub

Private Sub cmdE_Click()
Science = Val(txtOutput.Text)
txtOutput.Text = "0"
If Operators = "+" Then
  ComSci = Computer + Science
  txtOutput.Text = ComSci
 Else
 If Operators = "-" Then
 Network = Computer - Science
  txtOutput.Text = Network
  Else
  If Operators = "x" Then
 Keyboard = Computer * Science
 txtOutput.Text = Keyboard
 Else
 If Operators = "/" Then
 Minah = Computer / Science
  txtOutput.Text = Minah
  End If
  End If
  End If
  End If

End Sub

Private Sub cmdM_Click()
Computer = Val(txtOutput.Text)
txtOutput.Text = ""
Operators = "*"
End Sub

Private Sub cmdOff_Click()
MsgBox "Thank you using my calculator!! GodBless", vbInformation + vbOKOnly, "Created By:Minah"
End
End Sub

Private Sub cmdP_Click()
If txtOutput.Text = "0" Then
  txtOutput.Text = ""
  End If
 If InStr(txtOutput.Text, ".") = 0 Then
txtOutput.Text = txtOutput.Text & "."
  End If
End Sub

Private Sub cmdS_Click()
Computer = Val(txtOutput.Text)
txtOutput.Text = ""
Operators = "-"
End Sub

