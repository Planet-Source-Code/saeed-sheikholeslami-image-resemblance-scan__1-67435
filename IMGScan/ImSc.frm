VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Resemblance Scan"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "CLS"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS"
      Height          =   195
      Left            =   7800
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox B2 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   4200
      ScaleHeight     =   1995
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan Picture1 and Picture2"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resemblance both Picture
'programmer: SAEED SHEIKHOLESLAMI
'E -mail:: saeedsheikh1213@ yahoo.com
Dim one, two, D
Private Sub B1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
B1.PSet (X, Y)  ' for drawing on picture1
End If
End Sub
Private Sub B2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
B2.PSet (X, Y)  ' for drawing on picture2
End If
End Sub
Sub PaP(p1 As PictureBox, p2 As PictureBox, W As Long, H As Long, C As Long, DDD)
Dim X, Y, a, AP, DD1, DD2
D = 10  ' for decimal digit in percentile part. d= 1 or 10 or 100
For X = 0 To W - 15 Step C  ' StepbyStep for PixelComparable of both picture
For Y = 0 To H - 15 Step C
If p2.Point(X, Y) = p1.Point(X, Y) Then AP = AP + 1 ' if picture1 pointcolor= picture2 pointcolor |>counter=++1
a = a + 1  ' a=programme counter
Next
Next
DD1 = (AP * 100) \ a 'percentile
DD2 = Right$((AP * (100 * D)) \ a, Len(D) - 1) 'decimal part
DDD = DD1 & "." & DD2 & "%" ' wrought percent
End Sub
Private Sub Command1_Click()
PaP B1, B2, B1.Width, B2.Height, 15, one
MsgBox one
End Sub

Private Sub Command3_Click()
B1.Cls 'ClearScreen
End Sub

Private Sub Command2_Click()
B2.Cls 'ClearScreen
End Sub

Private Sub Form_Load()
B1.DrawWidth = 4 ' size pen=4 in both picture
B2.DrawWidth = 4
End Sub
