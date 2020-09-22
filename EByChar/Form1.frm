VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encrypt By Char - By: Cory J. Geesaman"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "57"
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "This works from the Text Out To -> Text In"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "This works from the Text In -> Text Out"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text Out"
      Top             =   360
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text In"
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tIn, tLast, a, tOut, tPass
tPass = CInt(Text3.Text)
tIn = Text1.Text
tLast = tPass
tOut = ""
i = 1
Do
a = Asc(Mid(tIn, i, 1)) + tLast + tPass
If a > 255 Then a = a - ((a \ 255) * 255)
tLast = a
tOut = tOut & Chr(a)
i = i + 1
Loop Until i > Len(tIn)
Text2.Text = tOut
End Sub

Private Sub Command2_Click()
Dim tIn, tLast, a, tOut, tPass, b
tPass = CInt(Text3.Text)
tIn = Text2.Text
tLast = tPass
tOut = ""
i = 1
Do
b = Asc(Mid(tIn, i, 1))
a = Asc(Mid(tIn, i, 1)) - tLast - tPass
tLast = b
If a < 0 Then a = a + 255
tOut = tOut & Chr(a)
i = i + 1
Loop Until i > Len(tIn)
Text1.Text = tOut
End Sub
