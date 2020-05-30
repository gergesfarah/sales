VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "«·‘«‘Â «·—∆Ì”Ì…"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command10 
      Caption         =   "Œ‹‹‹‹‹‹‹‹‹‹‹—ÊÃ"
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "„œ›Ê⁄«  «·⁄„·«¡ "
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "„œ›Ê⁄«  «·„Ê—œÌ‰"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«Ã—«¡ Õ—ﬂ… «·„»Ì⁄« "
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«Ã—«¡ Õ—ﬂ… «·„‘ —Ì«  "
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Õ‹‹‹‹‹‹‹Ê· «·»—‰‹«„Ã"
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   " ﬁ‹‹‹‹‹‹‹«—Ì‹—"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»Ì«‰«  «·«’‰«›"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "»Ì«‰«  «·⁄„·«¡"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command10_Click()
End
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
Form6.Show
End Sub

Private Sub Command7_Click()
Form7.Show
End Sub

Private Sub Command8_Click()
Form8.Show
End Sub

Private Sub Command9_Click()
Form9.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("works.mdb")

Set Vn = db.OpenRecordset("Vn", 2)
Set Ctm = db.OpenRecordset("Ctm", 2)
Set Itm = db.OpenRecordset("Itm", 2)
Set Tin = db.OpenRecordset("Tin", 2)
Set Sal = db.OpenRecordset("Sal", 2)
Set Vst = db.OpenRecordset("Vst", 2)
Set Vnd = db.OpenRecordset("Vnd", 2)
End Sub
