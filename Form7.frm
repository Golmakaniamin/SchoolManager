VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form7 
   BackColor       =   &H00709FC5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1560
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "»” ‰ Å‰Ã—Â"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   8.25
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
      MICON           =   "Form7.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ :"
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sd, qw As String
Private Sub Form_Activate()
List1.Clear
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub KewlButtons1_Click()
Form7.Hide
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  List1.Clear
  Form8.Data1.Recordset.MoveFirst
  Do
     If Form8.Data1.Recordset.Fields!numberstudent = Text1.Text Then
     Select Case Right(Form8.Data1.Recordset.Fields!numberterm, 1)
       Case 0
         sd = "Å‰Ã„ «» œ«∆Ì"
       Case 1
         sd = "«Ê· —«Â‰„«ÌÌ"
       Case 2
         sd = "œÊ„ —«Â‰„«ÌÌ"
       Case 3
         sd = "”Ê„ —«Â‰„«ÌÌ"
       Case 4
         sd = "«Ê· œ»Ì—” «‰"
       Case 5
         sd = "œÊ„ œ»Ì—” «‰"
       Case 6
         sd = "”Ê„ œ»Ì—” «‰"
       Case 7
         sd = "ÅÌ‘ œ«‰‘ê«ÂÌ"
     End Select
     qw = "  —„ " + Left(Form8.Data1.Recordset.Fields!numberterm, Len(Form8.Data1.Recordset.Fields!numberterm) - 2) + " Å«ÌÂ " + sd + " " + " œ— ò·«” "
     Form6.Data1.Recordset.FindFirst "idclass='" & Form8.Data1.Recordset.Fields!numberclass & "'"
     qw = qw + Form6.Data1.Recordset.Fields!numberclass + " "
          
     Select Case Form6.Data1.Recordset.Fields!mozoe
      Case 0
        qw = qw + "—Ì«÷Ì"
      Case 1
        qw = qw + "⁄·Ê„"
      Case 2
        qw = qw + "›Ì“Ìò"
      Case 3
        qw = qw + "‘Ì„Ì"
      Case 4
        qw = qw + "œÌ‰Ì"
      Case 5
        qw = qw + "⁄—»Ì"
      Case 6
        qw = qw + "“»«‰"
     End Select
     If Form6.Data1.Recordset.Fields!se = 1 Then qw = qw + "  Å”—«‰"
     If Form6.Data1.Recordset.Fields!se = 2 Then qw = qw + "  œŒ —«‰"
     List1.AddItem qw
     End If
     Form8.Data1.Recordset.MoveNext
  Loop Until Form8.Data1.Recordset.EOF = True
End If
End Sub
