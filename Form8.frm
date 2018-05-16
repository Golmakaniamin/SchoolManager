VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   ClientHeight    =   9195
   ClientLeft      =   3495
   ClientTop       =   1920
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List6 
      Height          =   405
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List5 
      Height          =   405
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List4 
      Height          =   405
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2520
      Width           =   3015
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   495
      Left            =   600
      TabIndex        =   22
      ToolTipText     =   "Å«ò ò—œ‰  „«„Ì «›—«œ œ— «Ì‰ ò·«”"
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "<<"
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
      MICON           =   "Form8.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List7 
      Height          =   405
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   7440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "Form8.frx":001C
      PICN            =   "Form8.frx":0038
      PICH            =   "Form8.frx":0DD2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4680
      Width           =   2175
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   10080
      TabIndex        =   17
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "»«“ê‘ "
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
      MICON           =   "Form8.frx":1B6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   3000
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
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
      MICON           =   "Form8.frx":1B88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   2640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
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
      MICON           =   "Form8.frx":1BA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data3"
      RightToLeft     =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "·Ì”  œ«‰‘ ¬„Ê“«‰ Å«ÌÂ «‰ Œ«»Ì"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form8.frx":1BC0
      Left            =   7920
      List            =   "Form8.frx":1BC2
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form8.frx":1BC4
      Left            =   7920
      List            =   "Form8.frx":1BC6
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "œ«‰‘ ¬„Ê“«‰ ò·«” «‰ Œ«» ‘œÂ"
      Top             =   2520
      Width           =   3015
   End
   Begin KewlButtonz.KewlButtons KewlButtons5 
      Height          =   495
      Left            =   600
      TabIndex        =   23
      ToolTipText     =   "«÷«›Â ò—œ‰ Ìò ò·«” »Â «Ì‰ ò·«”"
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   ">>"
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
      MICON           =   "Form8.frx":1BC8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õœ«òÀ— «›—«œ œ— «Ì‰ ò·«”"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ò·«” »‰œÌ œ«‰‘ ¬„Ê“«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   735
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò·«”"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„  —„ Â« Ê Å«ÌÂ Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «›—«œ œ— «Ì‰ ò·«”"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «›—«œ œ— «Ì‰ Å«ÌÂ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form8.frx":1BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mn As Long, po(1000) As String
Dim w, t, o, i As String, e, q, qw As Integer, v, r, y As Boolean, u, p As Long, sd As String

Private Sub Combo2_Click()
If Combo1.Text = "" Then
   q = MsgBox("‰«„  —„ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbMsgBoxRight, "")
Else
   
   For e = 0 To 1000
     po(e) = ""
   Next e
   List1.Clear
   List2.Clear
   Label1.Caption = List6.List(Combo2.ListIndex)
   e = 0
   Form3.Data1.Recordset.MoveFirst
   Do
     If (Left(Form3.Data1.Recordset.Fields!id, 1) = Left(List5.List(Combo2.ListIndex), 1)) And (Form3.Data1.Recordset.Fields!paye = Right(List7.List(Combo1.ListIndex), 1)) Then
        t = Form3.Data1.Recordset.Fields!id
        po(e) = t
        e = e + 1
     End If
     Form3.Data1.Recordset.MoveNext
   Loop Until Form3.Data1.Recordset.EOF = True
   For q = 0 To e - 1
     Form2.Data1.Recordset.FindFirst "id='" & po(q) & "'"
     List2.AddItem po(q) + " " + Form2.Data1.Recordset.Fields!Name + " " + Form2.Data1.Recordset.Fields!family
   Next q
   
   List1.Clear
   Data1.Recordset.MoveFirst
   Do
      If (Data1.Recordset.Fields!numberclass = List5.List(Combo2.ListIndex)) And (Data1.Recordset.Fields!numberterm = List7.List((Combo1.ListIndex))) Then
         List1.AddItem Data1.Recordset.Fields!numberstudent
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If

For q = 0 To List1.ListCount - 1
   Form2.Data1.Recordset.FindFirst "id='" & List1.List(q) & "'"
   sd = List1.List(q)
   List1.RemoveItem q
   List1.AddItem sd + " " + Form2.Data1.Recordset.Fields!Name + " " + Form2.Data1.Recordset.Fields!family, q
Next q
End Sub

Private Sub Combo3_Click()
If Combo1.Text = "" Then
   q = MsgBox("‰«„  —„ —« «‰ Œ«» ò‰Ìœ", vbCritical + vbMsgBoxRight, "")
Else
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If (Data1.Recordset.Fields!numberclass = List5.List(Combo3.ListIndex)) And (Data1.Recordset.Fields!numberterm = List7.List((Combo1.ListIndex))) Then
         List4.AddItem Data1.Recordset.Fields!numberstudent
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True

For mn = List4.ListCount - 1 To 0 Step -1
   Data1.Recordset.AddNew
   Data1.Recordset.Fields!numberstudent = List4.List(mn)
   Data1.Recordset.Fields!numberclass = List5.List(Combo2.ListIndex)
   Data1.Recordset.Fields!numberterm = List7.List(Combo1.ListIndex)
   Data1.Recordset.Update
   Form2.Data1.Recordset.FindFirst "id='" & List4.List(mn) & "'"
   List1.AddItem List4.List(mn) + " " + Form2.Data1.Recordset.Fields!Name + " " + Form2.Data1.Recordset.Fields!family
Next mn

Combo3.Visible = False
End If
End Sub

Private Sub Command2_Click()
Form8.Hide
Form15.Show
End Sub

Private Sub Form_Activate()
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo3.Visible = False
List1.Clear
List2.Clear

Form5.Data1.Recordset.MoveFirst
Do
  If Form5.Data1.Recordset.Fields!numberterm <> 0 Then
     Select Case Right(Form5.Data1.Recordset.Fields!numberterm, 1)
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
     Combo1.AddItem "  —„ " + Left(Form5.Data1.Recordset.Fields!numberterm, Len(Form5.Data1.Recordset.Fields!numberterm) - 2) + " Å«ÌÂ " + sd
     List7.AddItem Form5.Data1.Recordset.Fields!numberterm
  End If
  Form5.Data1.Recordset.MoveNext
Loop Until Form5.Data1.Recordset.EOF = True
End Sub

Private Sub Combo1_Click()
Dim stra As String
Label1.Caption = 0
List1.Clear
List5.Clear
Combo2.Clear
Combo3.Clear
Form6.Data1.Recordset.MoveFirst
Do
  If Form6.Data1.Recordset.Fields!numberterm = List7.List((Combo1.ListIndex)) Then
     List5.AddItem Form6.Data1.Recordset.Fields!idclass
     stra = "ò·«” " + Form6.Data1.Recordset.Fields!numberclass + " "
     List6.AddItem Form6.Data1.Recordset.Fields!plase
     Select Case Form6.Data1.Recordset.Fields!mozoe
      Case 0
        stra = stra + "—Ì«÷Ì"
      Case 1
        stra = stra + "⁄·Ê„"
      Case 2
        stra = stra + "›Ì“Ìò"
      Case 3
        stra = stra + "‘Ì„Ì"
      Case 4
        stra = stra + "œÌ‰Ì"
      Case 5
        stra = stra + "⁄—»Ì"
      Case 6
        stra = stra + "“»«‰"
     End Select
     If Form6.Data1.Recordset.Fields!se = 1 Then stra = stra + "  Å”—«‰"
     If Form6.Data1.Recordset.Fields!se = 2 Then stra = stra + "  œŒ —«‰"
     Combo2.AddItem stra
     Combo3.AddItem stra
  End If
  Form6.Data1.Recordset.MoveNext
Loop Until Form6.Data1.Recordset.EOF = True
List1.Clear
List2.Clear
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Caption = List2.ListCount
Label2.Caption = List1.ListCount
End Sub

Private Sub KewlButtons1_Click()
If Combo2.Text = "" Then
   q = MsgBox("‰«„ ò·«” —« «‰ Œ«» ò‰Ìœ", vbCritical + vbMsgBoxRight, "")
Else
   If Val(Label2.Caption) < Val(Label1.Caption) Then
     
     y = False
     For u = 0 To List1.ListCount - 1
         If Left(List1.List(u), 6) = Left(List2.Text, 6) Then
            y = True
            Exit For
         End If
     Next u
     If y = False Then
        If List2.List(List2.ListIndex) <> "" Then
           Data1.Recordset.AddNew
           Data1.Recordset.Fields!numberstudent = Left(List2.List(List2.ListIndex), 6)
           Data1.Recordset.Fields!numberclass = List5.List(Combo2.ListIndex)
           Data1.Recordset.Fields!numberterm = List7.List(Combo1.ListIndex)
           Data1.Recordset.Update
           List1.AddItem List2.List(List2.ListIndex)
        End If
        List2.ListIndex = List2.ListIndex + 1
        Label2.Caption = List1.ListCount
     Else
        q = MsgBox("«Ì‰ œ«‰‘ ¬„Ê“ ﬁ»·« œ— «Ì‰ ò·«” À»  ‰«„ ‘œÂ «” ", vbCritical + vbMsgBoxRight, "")
     End If
   Else
     q = MsgBox("Ÿ—›Ì  «Ì‰ ò·«” »Â Õœ«òÀ— ŒÊœ —”ÌœÂ «” ", vbCritical + vbMsgBoxRight, "")
   End If
End If
End Sub

Private Sub KewlButtons2_Click()
If List1.List(List1.ListIndex) <> "" Then
   Data1.Recordset.MoveFirst
   Do
      If (Data1.Recordset.Fields!numberstudent = Left(List1.List(List1.ListIndex), 6) And (Data1.Recordset.Fields!numberclass = List5.List(Combo2.ListIndex)) And (Data1.Recordset.Fields!numberterm = List7.List(Combo1.ListIndex))) Then
         Data1.Recordset.Delete
         List1.RemoveItem List1.ListIndex
         Label2.Caption = List1.ListCount
         List1.ListIndex = List1.ListIndex + 1
         Exit Do
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub KewlButtons3_Click()
Form7.Show
End Sub

Private Sub KewlButtons4_Click()
For qw = List1.ListCount - 1 To 0 Step -1
   Data1.Recordset.MoveFirst
   Do
      If (Data1.Recordset.Fields!numberclass = List5.List(Combo2.ListIndex)) And (Data1.Recordset.Fields!numberterm = List7.List(Combo1.ListIndex)) Then
         Data1.Recordset.Delete
         List1.RemoveItem qw
         Exit Do
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
Next qw
End Sub

Private Sub KewlButtons5_Click()
Combo3.Visible = True
End Sub

Private Sub List3_Click()
If List3.List(0) <> "" Then
   For intw = 0 To List2.ListCount - 1
       If List2.List(intw) = List3.List(List3.ListIndex) Then
          List2.ListIndex = intw
          KewlButtons1.SetFocus
          Exit For
       End If
   Next intw
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 96 Then Text1.Text = ""
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 96 Then Text2.Text = ""
If KeyCode = 13 Then
   If List2.List(0) <> "" Then
      If (Text1.Text <> "") And (Text2.Text <> "") Then
         i = Text1.Text + "  " + Text2.Text
         For p = 0 To List2.ListCount - 1
            g = Len(List2.List(p)) - 8
            o = Right(List2.List(p), g)
            If i = o Then
               List2.ListIndex = p
               KewlButtons1.SetFocus
               Exit For
            End If
         Next p
      End If
   
      If (Text1.Text <> "") And (Text2.Text = "*") Then
         i = Text1.Text
         Dim intq As Integer
         List3.Clear
         For p = 0 To List2.ListCount - 1
            intq = InStr(List2.List(p), Text1.Text)
            If intq <> 0 Then
               List3.AddItem List2.List(p)
            End If
         Next p
      End If
   End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List2.ListCount - 1
       If Left(List2.List(q), 6) = Text3.Text Then
          List2.ListIndex = q
          KewlButtons1.SetFocus
          Exit For
       End If
   Next q
End If
End Sub

