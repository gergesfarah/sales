VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "������ ����������� ��������"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   6630
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "����������"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "���������"
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
      Left            =   7200
      TabIndex        =   18
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����������"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�������"
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
      Left            =   10800
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ������"
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
      Left            =   12600
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13935
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
         Left            =   5040
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
         Left            =   5040
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
         Left            =   9840
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
         Left            =   5040
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
         Left            =   9840
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
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   9960
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������������"
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
         Left            =   12615
         TabIndex        =   14
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
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
         Left            =   8115
         TabIndex        =   12
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���� ������ �������"
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
         Left            =   11835
         TabIndex        =   10
         Top             =   2640
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���������� ���������"
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
         Left            =   7245
         TabIndex        =   8
         Top             =   1560
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������"
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
         Left            =   12900
         TabIndex        =   6
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���������������"
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
         Left            =   7680
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����� ����������������"
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
         Left            =   11820
         TabIndex        =   2
         Top             =   360
         Width           =   1635
      End
   End
End
Attribute VB_Name = "Form9"
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

Vst.AddNew
End Sub

Private Sub Command3_Click()
On Error Resume Next
Command3.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True

Vst("vst_nm") = Combo1.Text
Vst("vst_dt") = Text1.Text
Vst("vst_am") = Text2.Text
Vst("vst_ho") = Text3.Text
Vst("vst_no") = Text4.Text
Vst("vst_rcv") = Text5.Text
Vst("vst_dec") = Text6.Text


MsgBox "�� ����� �����"
Vst.Update
End Sub

Private Sub Command4_Click()
Dim r As String
r = InputBox("�� ���� ����� ��� ������", "�������")
If Trim$(r) = "" Then
MsgBox "����� ����� ����� ����"
Exit Sub
End If

Vnd.FindFirst "vst_nm like '" & r & "*'"
If Vnd.NoMatch Then
MsgBox "����� ����� ��� ����"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Exit Sub
End If

Vst("vst_nm") = Combo1.Text
Vst("vst_dt") = Text1.Text
Vst("vst_am") = Text2.Text
Vst("vst_ho") = Text3.Text
Vst("vst_no") = Text4.Text
Vst("vst_rcv") = Text5.Text
Vst("vst_dec") = Text6.Text

Command5.Enabled = True
End Sub


Private Sub Command5_Click()
t = MsgBox("Are you want delet", vbYesNo)
If t = vbYes Then
Tin.Delete
MsgBox "�� ����� �����"
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
If Ctm.RecordCount = 0 Then
MsgBox "������ ������ ��� ��� ������"
Exit Sub
End If

Combo1.Clear
Ctm.MoveFirst
Do While Not Ctm.EOF
Combo1.AddItem Ctm("ctm_nm")
Ctm.MoveNext
Loop
End Sub

Private Sub Form_Deactivate()

End Sub


Private Sub Form_Load()

End Sub
