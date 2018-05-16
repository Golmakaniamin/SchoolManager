VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form10 
   BackColor       =   &H00709FC5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons command1 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "»” ‰ Å‰Ã—Â"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form10.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "À» "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form10.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ﬁ»÷"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ œ—Ì«› Ì"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form10.Hide
End Sub

Private Sub Form_Activate()
If Label4.Caption = "1" Then
   If Form4.List5.List(Form4.List5.ListCount - 1) = "" Then
      Form10.Text1.Text = 1
   Else
      Form10.Text1.Text = Form4.List5.List(Form4.List5.ListCount - 1) + 1
   End If
   Text2.SetFocus
End If
End Sub

Private Sub KewlButtons1_Click()
   If Label4.Caption = "1" Then
      Form4.Data1.Recordset.AddNew
      Form4.Data1.Recordset.Fields!id = Form4.List1.List(Form4.List1.ListIndex)
      Form4.Data1.Recordset.Fields!numberterm = Form4.List9.List(Form4.Combo1.ListIndex)
      Form4.Data1.Recordset.Fields!Number = Text1.Text
      Form4.Data1.Recordset.Fields!padakht = Text2.Text
      Form4.Data1.Recordset.Fields!numbergabz = Text3.Text
      Form4.Data1.Recordset.Update
      Form10.Hide
    End If
   
   If Label4.Caption = "2" Then
      Form4.Data1.Recordset.MoveFirst
      Do
         If (Form4.Data1.Recordset.Fields!id = Form4.List1.List(Form4.List1.ListIndex)) And (Form4.Data1.Recordset.Fields!numberterm = Form4.List9.List(Form4.Combo1.ListIndex)) And (Form4.Data1.Recordset.Fields!Number = Form4.List5.List(Form4.List5.ListIndex)) And (Form4.Data1.Recordset.Fields!padakht = Form4.List6.List(Form4.List6.ListIndex)) And (Form4.Data1.Recordset.Fields!numbergabz = Form4.List7.List(Form4.List7.ListIndex)) Then
            Form4.Data1.Recordset.Edit
            Form4.Data1.Recordset.Fields!Number = Text1.Text
            Form4.Data1.Recordset.Fields!padakht = Text2.Text
            Form4.Data1.Recordset.Fields!numbergabz = Text3.Text
            Form4.Data1.Recordset.Update
            Exit Do
         End If
         Form4.Data1.Recordset.MoveNext
      Loop Until Form4.Data1.Recordset.EOF = True
      Form10.Hide
    End If
   
   Form4.Text3.Text = ""
   Form4.Text3.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons1.SetFocus
End Sub
