VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Õ‹—ﬂ‹… „‹œ›Ê⁄‹‹«  «·„Ê—œÌ‰ "
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   8025
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "⁄‹‹‹‹‹‹Êœ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   19
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Õ‹‹‹‹‹‹–›"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«”‹‹ œ⁄‹‹«¡"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Õ‹‹‹‹›Ÿ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Ã‹œÌ‹œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   13
         Top             =   4080
         Width           =   8535
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   11
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   9
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   7
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10920
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·»Ì‹‹‹‹‹‹«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   13575
         TabIndex        =   14
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„‹” ·„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9075
         TabIndex        =   12
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ‹„ «Ì‹’«· «·‹œ›‹⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12795
         TabIndex        =   10
         Top             =   2640
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÿ‹‹‹—Ì‹‹ﬁ… «·‹œ›‹‹‹⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8205
         TabIndex        =   8
         Top             =   1560
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„‹»·€"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   13860
         TabIndex        =   6
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«· ‹‹‹‹‹«—Ì‹‹‹Œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "√”‹‹„ «·‹„‹‹‹‹‹Ê—œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   13020
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

Private Sub Command3_Click()
On Error Resume Next
Command3.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True

Vnd("vnd_nm") = Combo1.Text
Vnd("vnd_dt") = Text1.Text
Vnd("vnd_am") = Text2.Text
Vnd("vnd_ho") = Text3.Text
Vnd("vnd_no") = Text4.Text
Vnd("vnd_rcv") = Text5.Text
Vnd("vnd_dec") = Text6.Text


MsgBox " „ «·Õ›Ÿ »‰Ã«Õ"
Vnd.Update
End Sub

Private Sub Command4_Click()
Dim v As String
v = InputBox("„‰ ›÷·ﬂ «œŒ· «”„ «·„Ê—œ", "«” œ⁄«¡")
If Trim$(v) = "" Then
MsgBox "»—Ã«¡ «œŒ«· «·«”„ ’ÕÌÕ"
Exit Sub
End If

Vnd.FindFirst "vnd_nm like '" & v & "*'"
If Vnd.NoMatch Then
MsgBox "⁄›Ê« «·«”„ €Ì— „”Ã·"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Exit Sub
End If

Vnd("vnd_nm") = Combo1.Text
Vnd("vnd_dt") = Text1.Text
Vnd("vnd_am") = Text2.Text
Vnd("vnd_ho") = Text3.Text
Vnd("vnd_no") = Text4.Text
Vnd("vnd_rcv") = Text5.Text
Vnd("vnd_dec") = Text6.Text

Command5.Enabled = True
End Sub


Private Sub Command5_Click()
On Error Resume Next
If Vnd.RecordCount = 0 Then
MsgBox "No record"
Exit Sub
End If

Dim t As String
t = MsgBox("Are you want delet", vbYesNo)
If t = vbYes Then
Tin.Delete
MsgBox " „ «·Õ–› »‰Ã«Õ"
Exit Sub
End If

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


Command5.Enabled = False
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Vn.RecordCount = 0 Then
MsgBox "·«ÌÊÃœ „Ê—œÌ‰ Õ Ì  ﬁÊ„ » ”ÃÌ· «·«’‰«›"
Me.Hide
Exit Sub
End If
Vn.MoveFirst
Combo1.Clear
Do While Not Vn.EOF
Combo1.AddItem Vn("vn_nm")
Vn.MoveNext
Loop
End Sub

Private Sub Form_Load()

End Sub


