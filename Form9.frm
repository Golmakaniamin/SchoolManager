VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form9 
   BackColor       =   &H00709FC5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form9.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List4 
      Height          =   1095
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      ItemData        =   "Form9.frx":001C
      Left            =   600
      List            =   "Form9.frx":0038
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
   Begin KewlButtonz.KewlButtons command1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2400
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
      MICON           =   "Form9.frx":00AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "„ﬁÿ⁄  Õ’Ì·Ì"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "¬Œ—Ì‰ „⁄œ·"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "‰«„ „œ—”Â"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "‘Ì› "
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Byte, w As Boolean, e As Integer

Private Sub Command1_Click()
   If Label5.Caption = "1" Then
      Form3.Data1.Recordset.AddNew
      Form3.Data1.Recordset.Fields!school = Text1.Text
      Form3.Data1.Recordset.Fields!paye = Combo2.ListIndex
      Form3.Data1.Recordset.Fields!Shift = Text3.Text
      Form3.Data1.Recordset.Fields!endaverage = Text4.Text
      Form3.Data1.Recordset.Fields!id = Form3.Label2.Caption
      Form3.Data1.Recordset.Update
      Form3.List20.AddItem Combo2.Text
      Form3.List30.AddItem Text1.Text
      Form3.List4.AddItem Text3.Text
      Form3.List5.AddItem Text4.Text
      Text1.SetFocus
   End If
   If Label5.Caption = "2" Then
      Form3.Data1.Recordset.FindFirst "id='" & Form3.Label2.Caption & "'"
      If Form3.Data1.Recordset.NoMatch = False Then
         Form3.Data1.Recordset.Edit
         Form3.Data1.Recordset.Fields!school = Text1.Text
         Form3.Data1.Recordset.Fields!paye = Combo2.ListIndex
         Form3.Data1.Recordset.Fields!Shift = Text3.Text
         Form3.Data1.Recordset.Fields!endaverage = Text4.Text
         Form3.Data1.Recordset.Update
       End If
   End If
Form9.Hide
End Sub

Private Sub Form_Activate()
List4.Visible = False
End Sub

Private Sub KewlButtons1_Click()
Form9.Hide
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  
  If (List4.Left = 600) And (List4.Top = 600) Then
    Text1.Text = List4.Text
    Text1.SetFocus
    List4.Visible = False
  End If
  
  If (List4.Left = 600) And (List4.Top = 1560) Then
    Text3.Text = List4.Text
    Text3.SetFocus
    List4.Visible = False
  End If

End If
End Sub

Private Sub Text1_Change()
List4.Visible = False
If Len(Text1.Text) > 0 Then
   List4.Left = 600
   List4.Top = 600
   List4.Visible = True
   List4.Clear
   Form3.Data1.Recordset.MoveFirst
   Do
      If Left(Form3.Data1.Recordset.Fields!school, Len(Text1.Text)) = Text1.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Form3.Data1.Recordset.Fields!school Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Form3.Data1.Recordset.Fields!school
      End If
      Form3.Data1.Recordset.MoveNext
   Loop Until Form3.Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo2.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_Change()
List4.Visible = False
If Len(Text1.Text) > 0 Then
   List4.Left = 600
   List4.Top = 1560
   List4.Visible = True
   List4.Clear
   List4.AddItem "’»Õ"
   List4.AddItem "⁄’—"
   List4.AddItem "ê—œ‘Ì"
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub
