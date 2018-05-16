VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   9210
   ClientLeft      =   1965
   ClientTop       =   585
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1260
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   1560
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1260
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "»«“ê‘ "
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
      MICON           =   "Form3.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "‰›— »⁄œ"
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
      MICON           =   "Form3.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command3 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "ÊÌ—«Ì‘"
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
      MICON           =   "Form3.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons command7 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "«ÿ·«⁄«  ÃœÌœ"
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
      MICON           =   "Form3.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      ItemData        =   "Form3.frx":0070
      Left            =   1440
      List            =   "Form3.frx":0072
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      ItemData        =   "Form3.frx":0074
      Left            =   3120
      List            =   "Form3.frx":0076
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ListBox List30 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      ItemData        =   "Form3.frx":0078
      Left            =   4920
      List            =   "Form3.frx":007A
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ListBox List20 
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      ItemData        =   "Form3.frx":007C
      Left            =   6840
      List            =   "Form3.frx":007E
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data2"
      RightToLeft     =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1260
   End
   Begin KewlButtonz.KewlButtons command8 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Õ–›"
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
      MICON           =   "Form3.frx":0080
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
      Height          =   135
      Left            =   9600
      TabIndex        =   21
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      BTYPE           =   3
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
      MICON           =   "Form3.frx":009C
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
      Height          =   135
      Left            =   9600
      TabIndex        =   22
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      BTYPE           =   3
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
      MICON           =   "Form3.frx":00B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À»  «ÿ·«⁄«   Õ’Ì·Ì"
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
      Height          =   615
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
      Height          =   375
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰ ÌÃÂ"
      Height          =   375
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   6000
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Left            =   9240
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      Height          =   375
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   375
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "òœ "
      Height          =   375
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "òœ"
      Height          =   375
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "¬Œ—Ì‰ „⁄œ·"
      Height          =   375
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ì› "
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ „œ—”Â"
      Height          =   375
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁÿ⁄  Õ’Ì·Ì"
      Height          =   375
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form3.frx":00D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Byte, q As Integer, w As Long, r As Long, t, y As Long, u As Boolean

Private Sub Command1_Click()
Text6.Text = Val(Label2.Caption) + 1
Text6.SetFocus
SendKeys ("{enter}")
End Sub

Private Sub Command2_Click()
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Form3.Hide
Form15.Show
End Sub

Private Sub Command3_Click()
Form9.Text1.Text = List30.List(List30.ListIndex)
Form9.Combo2.Text = List20.List(List20.ListIndex)
Form9.Text3.Text = List4.List(List4.ListIndex)
Form9.Text4.Text = List5.List(List5.ListIndex)
Form9.Label5.Caption = "2"
Form9.Show 1
Command3.Enabled = True
End Sub

Private Sub Command7_Click()
If Label2.Caption <> "" Then
   Form9.Label5.Caption = "1"
   If List2.List(0) <> "" Then
      List2.ListIndex = 0
      Form9.Text1.Text = List30.List(List3.ListIndex)
      Form9.Combo2.Text = List20.List(List2.ListIndex)
      Form9.Text3.Text = List4.List(List4.ListIndex)
      Form9.Text4.Text = List5.List(List5.ListIndex)
   End If
   Form9.Show vbModal
Else
   e = MsgBox("·ÿ›« ‘Œ’ „Ê—œ ‰Ÿ— ŒÊœ —« Ã” ÃÊ ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub Command8_Click()
Data1.Recordset.MoveFirst
Do
   If (Data1.Recordset.Fields!id = Label2.Caption) And (Data1.Recordset.Fields!paye = List20.Text) And (Data1.Recordset.Fields!school = List30.Text) And (Data1.Recordset.Fields!Shift = List4.Text) And (Data1.Recordset.Fields!endaverage = List5.Text) Then
       Data1.Recordset.Delete
       Data1.Refresh
       List20.RemoveItem (List20.ListIndex)
       List30.RemoveItem (List30.ListIndex)
       List4.RemoveItem (List4.ListIndex)
       List5.RemoveItem (List5.ListIndex)
       Exit Do
   End If
   Data1.Recordset.MoveNext
Loop Until Data1.Recordset.EOF = True
Command8.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Form_Activate()
Image2.Picture = LoadPicture("")
Command3.Enabled = False
Command8.Enabled = False
Text6.SetFocus
y = 0
t = 0
List2.Clear
List3.Clear
List20.Clear
List30.Clear
List5.Clear
List4.Clear
Text16.Text = ""
Form2.Data1.Recordset.MoveFirst
Do
   List2.AddItem Form2.Data1.Recordset.Fields!Name + "  " + Form2.Data1.Recordset.Fields!family
   List1.AddItem Form2.Data1.Recordset.Fields!id
   Form2.Data1.Recordset.MoveNext
Loop Until Form2.Data1.Recordset.EOF = True
If Label2.Caption <> "" Then
  Text6.Text = Label2.Caption
  Text6.SetFocus
  SendKeys ("{enter}")
End If
End Sub

Private Sub KewlButtons1_Click()
Dim id(1000), na(1000), idt, nat, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If id(intq) > id(intw) Then
         idt = id(intq)
         nat = na(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         
         id(intw) = idt
         na(intw) = nat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons3_Click()
Dim na(1000), nat, count As String
For intq = 0 To List3.ListCount - 1
    na(intq) = List3.List(intq)
Next intq
count = List3.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         nat = na(intq)
         
         na(intq) = na(intw)
         
         na(intw) = nat
      End If
   Next intw
Next intq
List3.Clear
For intq = 0 To count
   List3.AddItem na(intq)
Next intq
End Sub

Private Sub List2_Click()
Text6.Text = List1.List(List2.ListIndex)
Call Text6_KeyDown(13, 0)
End Sub

Private Sub List20_Click()
List30.ListIndex = List20.ListIndex
List4.ListIndex = List20.ListIndex
List5.ListIndex = List20.ListIndex
End Sub

Private Sub List3_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
End Sub

Private Sub List30_Click()
List20.ListIndex = List30.ListIndex
List4.ListIndex = List30.ListIndex
List5.ListIndex = List30.ListIndex
End Sub

Private Sub List4_Click()
List20.ListIndex = List4.ListIndex
List30.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List20.ListIndex = List5.ListIndex
List30.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
Command3.Enabled = True
Command8.Enabled = True
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Form2.Data1.Recordset.MoveFirst
   Form2.Data1.Recordset.FindFirst "id='" & Text6.Text & "'"
   If Form2.Data1.Recordset.NoMatch = True Then e = MsgBox("«Ì‰ òœ ÊÃÊœ ‰œ«—œ", vbCritical + vbMsgBoxRight, "")
   If Form2.Data1.Recordset.NoMatch = False Then
      Label2.Caption = Form2.Data1.Recordset.Fields!id
      Label3.Caption = Form2.Data1.Recordset.Fields!Name
      Label4.Caption = Form2.Data1.Recordset.Fields!family
      List20.Clear
      List30.Clear
      List4.Clear
      List5.Clear
      Image2.Picture = LoadPicture("")
      Data1.Recordset.MoveFirst
      Do
         If Data1.Recordset.Fields!id = Text6.Text Then
     
     Select Case Data1.Recordset.Fields!paye
       Case 0
         List20.AddItem "Å‰Ã„ «» œ«∆Ì"
       Case 1
         List20.AddItem "«Ê· —«Â‰„«ÌÌ"
       Case 2
         List20.AddItem "œÊ„ —«Â‰„«ÌÌ"
       Case 3
         List20.AddItem "”Ê„ —«Â‰„«ÌÌ"
       Case 4
         List20.AddItem "«Ê· œ»Ì—” «‰"
       Case 5
         List20.AddItem "œÊ„ œ»Ì—” «‰"
       Case 6
         List20.AddItem "”Ê„ œ»Ì—” «‰"
       Case 7
         List20.AddItem "ÅÌ‘ œ«‰‘ê«ÂÌ"
     End Select
            
            List30.AddItem Data1.Recordset.Fields!school
            List4.AddItem Data1.Recordset.Fields!Shift
            List5.AddItem Data1.Recordset.Fields!endaverage
          End If
         Data1.Recordset.MoveNext
       Loop Until Data1.Recordset.EOF = True
       t = t + 1
   End If
   Text6.Text = ""
   command7.SetFocus
   If basItemExist.ItemExist(App.Path + "DATA BASE INFORMATION\Image\" + Label2.Caption + ".jpg") = True Then
     Image2.Picture = LoadPicture(App.Path + "DATA BASE INFORMATION\Image\" + Label2.Caption + ".jpg")
   End If
End If
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Form2.Data1.Recordset.MoveFirst
   Form2.Data1.Recordset.FindFirst "id='" & Text6.Text & "'"
   If Form2.Data1.Recordset.NoMatch = True Then e = MsgBox("«Ì‰ òœ ÊÃÊœ ‰œ«—œ", vbCritical + vbMsgBoxRight, "")
   If Form2.Data1.Recordset.NoMatch = False Then
      Label2.Caption = Form2.Data1.Recordset.Fields!id
      Label3.Caption = Form2.Data1.Recordset.Fields!Name
      Label4.Caption = Form2.Data1.Recordset.Fields!family
      List20.Clear
      List30.Clear
      List4.Clear
      List5.Clear
      Data1.Recordset.MoveFirst
      Do
         If Data1.Recordset.Fields!id = Text6.Text Then
            List20.AddItem Data1.Recordset.Fields!paye
            List30.AddItem Data1.Recordset.Fields!school
            List4.AddItem Data1.Recordset.Fields!Shift
            List5.AddItem Data1.Recordset.Fields!endaverage
          End If
         Data1.Recordset.MoveNext
       Loop Until Data1.Recordset.EOF = True
       t = t + 1
   End If
   Text6.Text = ""
   command7.SetFocus
End If
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   List3.Clear
   For q = 0 To List2.ListCount - 1
       If InStr(List2.List(q), Text16.Text) <> 0 Then
          List3.AddItem List2.List(q)
       End If
   Next q
End If
End Sub

