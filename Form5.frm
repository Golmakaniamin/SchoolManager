VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00709FC5&
   BorderStyle     =   0  'None
   ClientHeight    =   4755
   ClientLeft      =   2790
   ClientTop       =   4170
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Õ–›"
      ENAB            =   0   'False
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
      MICON           =   "Form5.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00709FC5&
      Height          =   1575
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2400
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   420
         ItemData        =   "Form5.frx":001C
         Left            =   3240
         List            =   "Form5.frx":0038
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "Å‰Ã„ «» œ«∆Ì"
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "ÃœÌœ"
         Height          =   375
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "ÊÌ—«Ì‘"
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "„«‰‰œ : 1Ê2Ê3Ê"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â  —„ "
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ Å«ÌÂ"
         Height          =   375
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ‘—Ê⁄"
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ Å«Ì«‰"
         Height          =   375
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Â“Ì‰Â À»  ‰«„"
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.ListBox List5 
      Height          =   1560
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Height          =   1560
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   1560
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   1560
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1560
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   4080
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
      MICON           =   "Form5.frx":00AC
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
      Left            =   360
      TabIndex        =   27
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "À» "
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
      MICON           =   "Form5.frx":00C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ  —„ Â« :"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Â“Ì‰Â"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Å«Ì«‰"
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ‘—Ê⁄"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„  —„"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   120
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub Command2_Click()
Form5.Hide
End Sub

Private Sub Form_Activate()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear

Data1.Recordset.MoveFirst
Do
   If Data1.Recordset.Fields!numberterm <> 0 Then
      List1.AddItem Data1.Recordset.Fields!numberterm
      List2.AddItem Data1.Recordset.Fields!nameterm
      List3.AddItem Data1.Recordset.Fields!startdate
      List4.AddItem Data1.Recordset.Fields!enddate
      List5.AddItem Data1.Recordset.Fields!Money
   End If
   Data1.Recordset.MoveNext
Loop Until Data1.Recordset.EOF = True
Text1.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label11.Caption = List2.ListCount
End Sub

Private Sub KewlButtons1_Click()
If List2.ListIndex <> -1 Then
  Dim q As String
  q = List2.Text
  w = List2.ListIndex
  inta = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
  If inta = 6 Then
  '1
  Data1.Recordset.FindFirst "nameterm='" & q & "'"
  If Data1.Recordset.NoMatch = False Then
     Data1.Recordset.Delete
     Data1.Refresh
     List1.RemoveItem (w)
     List2.RemoveItem (w)
     List3.RemoveItem (w)
     List4.RemoveItem (w)
     List5.RemoveItem (w)
  End If
  
  '2
  Form6.Data1.Recordset.MoveFirst
  Do
     If Form6.Data1.Recordset.Fields!numberterm = q Then
        Form6.Data1.Recordset.Delete
        Form6.Data1.Refresh
     End If
     Form6.Data1.Recordset.MoveNext
  Loop Until Form6.Data1.Recordset.EOF = True
  
  '3
  Form8.Data1.Recordset.MoveFirst
  Do
     If Form8.Data1.Recordset.Fields!numberterm = q Then
        Form8.Data1.Recordset.Delete
        Form8.Data1.Refresh
     End If
     Form8.Data1.Recordset.MoveNext
  Loop Until Form8.Data1.Recordset.EOF = True
  
  '4
  Form4.Data1.Recordset.MoveFirst
  Do
     If Form4.Data1.Recordset.Fields!numberterm = q Then
        Form4.Data1.Recordset.Delete
        Form4.Data1.Refresh
     End If
     Form4.Data1.Recordset.MoveNext
  Loop Until Form4.Data1.Recordset.EOF = True
  
  '5
  Form4.Data2.Recordset.MoveFirst
  Do
     If Form4.Data2.Recordset.Fields!numberterm = q Then
        Form4.Data2.Recordset.Delete
        Form4.Data2.Refresh
     End If
     Form4.Data2.Recordset.MoveNext
  Loop Until Form4.Data2.Recordset.EOF = True
  
  '5
  Form3.Data1.Recordset.MoveFirst
  Do
     If Form3.Data1.Recordset.Fields!numberterm = q Then
        Form3.Data1.Recordset.Delete
        Form3.Data1.Refresh
     End If
     Form3.Data1.Recordset.MoveNext
  Loop Until Form3.Data1.Recordset.EOF = True
  
  '6
  Form21.Data1.Recordset.MoveFirst
  Do
     If Form21.Data1.Recordset.Fields!nt = q Then
        Form21.Data1.Recordset.Delete
        Form21.Data1.Refresh
     End If
     Form21.Data1.Recordset.MoveNext
  Loop Until Form21.Data1.Recordset.EOF = True
  
  '7
  Form21.Data2.Recordset.MoveFirst
  Do
     If Form21.Data2.Recordset.Fields!nt = q Then
        Form21.Data2.Recordset.Delete
        Form21.Data2.Refresh
     End If
     Form21.Data2.Recordset.MoveNext
  Loop Until Form21.Data2.Recordset.EOF = True
  
  '8
  Form21.Data3.Recordset.MoveFirst
  Do
     If Form21.Data3.Recordset.Fields!nt = q Then
        Form21.Data3.Recordset.Delete
        Form21.Data3.Refresh
     End If
     Form21.Data3.Recordset.MoveNext
  Loop Until Form21.Data3.Recordset.EOF = True
  
  '9
  Form21.Data4.Recordset.MoveFirst
  Do
     If Form21.Data4.Recordset.Fields!nt = q Then
        Form21.Data4.Recordset.Delete
        Form21.Data4.Refresh
     End If
     Form21.Data4.Recordset.MoveNext
  Loop Until Form21.Data4.Recordset.EOF = True
  
  End If
End If
End Sub

Private Sub KewlButtons2_Click()
   If Option1.Value = True Then
      Dim a As Boolean
      Text1.Text = Text1.Text + "-" + Trim(Str(Combo1.ListIndex))
      a = False
      For q = 0 To List1.ListCount - 1
        If Text1.Text = List1.List(q) Then
           For w = 0 To List2.ListCount - 1
             If Combo1.Text = List2.List(q) Then a = True: Exit For
           Next w
        End If
      Next q
      If a = False Then
        Data1.Recordset.AddNew
        Data1.Recordset.Fields!numberterm = Text1.Text
        Data1.Recordset.Fields!nameterm = Combo1.Text
        Data1.Recordset.Fields!startdate = Text3.Text
        Data1.Recordset.Fields!enddate = Text4.Text
        Data1.Recordset.Fields!Money = Text5.Text
        Data1.Recordset.Update
        List1.AddItem Text1.Text
        List2.AddItem Combo1.Text
        List3.AddItem Text3.Text
        List4.AddItem Text4.Text
        List5.AddItem Text5.Text
        Text1.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text1.SetFocus
      End If
    End If
    
   If Option2.Value = True Then
      Data1.Recordset.FindFirst "numberterm='" & List1.Text & "'"
      Data1.Recordset.Edit
      Data1.Recordset.Fields!numberterm = Text1.Text
      Data1.Recordset.Fields!nameterm = Combo1.Text
      Data1.Recordset.Fields!startdate = Text3.Text
      Data1.Recordset.Fields!enddate = Text4.Text
      Data1.Recordset.Fields!Money = Text5.Text
      Data1.Recordset.Update
      
      w = List2.ListIndex
      List1.RemoveItem (w)
      List2.RemoveItem (w)
      List3.RemoveItem (w)
      List4.RemoveItem (w)
      List5.RemoveItem (w)
      
      List1.AddItem Text1.Text, w
      List2.AddItem Combo1.Text, w
      List3.AddItem Text3.Text, w
      List4.AddItem Text4.Text, w
      List5.AddItem Text5.Text, w
      Text1.Locked = True
      Combo1.Locked = True
    End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List1.ListIndex = List5.ListIndex
List2.ListIndex = List5.ListIndex
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
End Sub

Private Sub Option1_Click()
Text1.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text1.Locked = False
Combo1.Locked = False
Text1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Text = List1.Text
Combo1 = List2.Text
Text3.Text = List3.Text
Text4.Text = List4.Text
Text5.Text = List5.Text
Text1.Locked = True
Combo1.Locked = True
Text3.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo1.SetFocus
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons2.SetFocus
End Sub
