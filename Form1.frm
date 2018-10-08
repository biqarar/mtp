VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MTP"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      ToolTipText     =   " ⁄œ«œ ’›Õ« Ì »«Ìœ œ—Ê‰ œ” ê«Â Å—Ì‰ — »«‘œ "
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "H"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Å«ﬂ ﬂ—œ‰ ÃœÊ·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   7575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      ToolTipText     =   " ⁄œ«œ ’›Õ«  »«Ìœ „÷—»Ì «“ 4 »«‘‰œ"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Õ«”»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ ’›Õ«   Å—Ì‰ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      TabIndex        =   11
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "„—ﬂ“ ﬁ—¬‰ Ê ÕœÌÀ ﬂ—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   2520
      TabIndex        =   9
      Top             =   2640
      Width           =   4020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Å—Ì‰  œÊ„"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   7
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Å—Ì‰  «Ê·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   6
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " ⁄œ«œ ﬂ· ’›Õ« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8160
      TabIndex        =   5
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "‰—„ «›“«— „Õ«”»Â  — Ì» Å—Ì‰  (›ﬁÿ »—«Ì Å‘  Ê —Ê)1"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   5250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A, B, C, D, E, F, G, H As String
 Dim I As Integer
 
A = 1
B = Val(Text1.Text)
'C = B / 2
'D = C + 1
'E = Val(Text1.Text)

If B / 2 = B \ 2 Then
'Text4.Text = B / 4

'If B = 4 Then
'Text2.Text = Text2.Text & B & "," & A
'Text3.Text = Text3.Text & C & "," & D
'Exit Sub
'End If


For I = 1 To B Step 2
'Text2.Text = Text2.Text & B & "," & A & ","
'Text3.Text = Text3.Text & C & "," & D & ","
'A = A + 2
'B = B - 2
'C = C - 2
'D = D + 2
If I = 1 Then
Text2.Text = I
GoTo 1
End If
Text2.Text = Text2.Text & "," & I
1:


Text3.Text = (I + 1) & "," & Text3.Text
'If D = Val(Text1.Text) - 1 Then
'Text2.Text = Text2.Text & B & "," & A
'Text3.Text = Text3.Text & C & "," & D
'Exit Sub

'End If

2:
Next I
Else
MsgBox "⁄œœ Ê«—œ ‘œÂ »«Ìœ „÷—»Ì «“ 2 »«‘œ ", vbInformation, "‰—„ «›“«— „Õ«”»Â  —»Ì  Å—Ì‰ "

End If





End Sub

Private Sub Command2_Click()
Text4 = ""
Text2 = ""
Text3 = ""

End Sub

Private Sub Command3_Click()
MsgBox "Mail:BiQarar1408.251.3@gmail.com" & Chr(10) & "It's Free" & Chr(10) & "By:Reza Mohiti", vbInformation, "RM"

End Sub

Private Sub Text1_DblClick()
Text1.Text = ""

End Sub
