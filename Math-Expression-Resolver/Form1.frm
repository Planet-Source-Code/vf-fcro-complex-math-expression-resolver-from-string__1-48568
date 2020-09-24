VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SYMBOL/NODE ENGINE with MATH EXPRESSION RESOLVER By Vanja Fuckar"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1440
      Width           =   7815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SYMBOL/NODE ENGINE 'TEXT TYPE' + MATH EXPRESSION EXAMPLE by Vanja Fuckar,EMAIL:inga@vip.hr"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Try This Expressions"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EAE1D5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EAE1D5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expression"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "THAT's pure ""INLINE CALCULATOR""! based on SYMBOL/NODE ENGINE!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MATH_C As New MathPlugin
Attribute MATH_C.VB_VarHelpID = -1

Private Sub Combo1_Click()
If Combo1.ListIndex = -1 Then Exit Sub
Text1 = Combo1
Text1_KeyPress 13
End Sub






Private Sub Form_Load()
MsgBox "Type MATH expression Like:" & vbCrLf & "5+5" & vbCrLf & "or" & vbCrLf & "Try more difficult expression from Combo", , "!*!"
Combo1.Clear
Combo1 = ""
Combo1.AddItem "sqr(19-tan(98)*tan(91)-sin(122)*(5*5-(199-12)))"
Combo1.AddItem "(2-12)*(5*123)-129+(1-199*(2*133))"
Combo1.AddItem "(10 mod 3 )/ 2^3-1981"
Combo1.AddItem "5*2/(5-12-1*(2-(29-12*2)))"
Combo1.AddItem "1-log(1/100)"
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim EXPR As Variant
Dim IsError As Boolean
EXPR = MATH_C.MathExpression(Text1, IsError)
    If IsError Then
    Text2 = "Error!"
    Else
    Text2 = EXPR
    End If
End If

End Sub
