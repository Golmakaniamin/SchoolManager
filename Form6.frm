VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form6 
   BackColor       =   &H00709FC5&
   BorderStyle     =   0  'None
   ClientHeight    =   9210
   ClientLeft      =   2580
   ClientTop       =   4335
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List8 
      Height          =   2460
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00709FC5&
      Caption         =   "Ã‰”Ì  ò·«”"
      Height          =   735
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2640
      Width           =   2055
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "œŒ —«‰"
         Height          =   300
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "Å”—«‰"
         Height          =   300
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ListBox List7 
      Height          =   2460
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00709FC5&
      Height          =   1455
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6720
      Width           =   11055
      Begin VB.ComboBox Combo1 
         Height          =   420
         ItemData        =   "Form6.frx":0000
         Left            =   2640
         List            =   "Form6.frx":0019
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "«ÿ·«⁄«  À»  „Ì ‘Êœ enter »«“œ‰ œò„Â"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "«ÿ·«⁄«  À»  „Ì ‘Êœ enter »«“œ‰ œò„Â"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "«ÿ·«⁄«  À»  „Ì ‘Êœ enter »«“œ‰ œò„Â"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "«ÿ·«⁄«  À»  „Ì ‘Êœ enter »«“œ‰ œò„Â"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "«ÿ·«⁄«  À»  „Ì ‘Êœ enter »«“œ‰ œò„Â"
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "ÊÌ—«Ì‘"
         Height          =   375
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00709FC5&
         Caption         =   "ÃœÌœ"
         Height          =   375
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò·«”"
         Height          =   255
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê÷Ê⁄ ò·«”"
         Height          =   255
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ œ»Ì—"
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Õœ«òÀ— «›—«œ"
         Height          =   255
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "—Ê“"
         Height          =   255
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "”«⁄ "
         Height          =   255
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.ListBox List6 
      Height          =   2460
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Height          =   2460
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ListBox List4 
      Height          =   2460
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   2460
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   10320
      TabIndex        =   7
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
      MICON           =   "Form6.frx":0049
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data2"
      RightToLeft     =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ListBox List2 
      Height          =   2460
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2460
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2415
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   8280
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
      MICON           =   "Form6.frx":0065
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
      Left            =   720
      TabIndex        =   38
      Top             =   8280
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
      MICON           =   "Form6.frx":0081
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê÷Ê⁄ ò·«”"
      Height          =   375
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄—›Ì ò·«” Â«"
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
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ ò·«” Â« :"
      Height          =   375
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "”«⁄ "
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "—Ê“"
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Õœ«òÀ— «›—«œ"
      Height          =   375
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ œ»Ì—"
      Height          =   375
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”   —„ Â« Ê Å«ÌÂ Â«"
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ò·«”"
      Height          =   375
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form6.frx":009D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String, q1(5, 100) As String, a, k As Integer

Private Sub Command2_Click()
Form6.Hide
End Sub

Private Sub Form_Activate()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Form5.Data1.Recordset.MoveFirst
Do
  If Form5.Data1.Recordset.Fields!numberterm <> 0 Then
     Select Case Right(Form5.Data1.Recordset.Fields!numberterm, 1)
       Case 0
         p = "Å‰Ã„ «» œ«∆Ì"
       Case 1
         p = "«Ê· —«Â‰„«ÌÌ"
       Case 2
         p = "œÊ„ —«Â‰„«ÌÌ"
       Case 3
         p = "”Ê„ —«Â‰„«ÌÌ"
       Case 4
         p = "«Ê· œ»Ì—” «‰"
       Case 5
         p = "œÊ„ œ»Ì—” «‰"
       Case 6
         p = "”Ê„ œ»Ì—” «‰"
       Case 7
         p = "ÅÌ‘ œ«‰‘ê«ÂÌ"
     End Select
     List1.AddItem "  —„ " + Left(Form5.Data1.Recordset.Fields!numberterm, Len(Form5.Data1.Recordset.Fields!numberterm) - 2) + " Å«ÌÂ  " + p
     List7.AddItem Form5.Data1.Recordset.Fields!numberterm
  End If
  Form5.Data1.Recordset.MoveNext
Loop Until Form5.Data1.Recordset.EOF = True
End Sub

Private Sub KewlButtons1_Click()
If List2.ListIndex <> -1 Then
  q = List2.Text
  s = List7.Text
  inta = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
  If inta = 6 Then
  '1
  Data1.Recordset.MoveFirst
  Do
     If (Data1.Recordset.Fields!numberterm = s) And (Data1.Recordset.Fields!numberclass = q) Then
        Data1.Recordset.Delete
        Data1.Refresh
     End If
     Data1.Recordset.MoveNext
  Loop Until Data1.Recordset.EOF = True
  List2.RemoveItem (List2.ListIndex)
  List3.RemoveItem (List3.ListIndex)
  List4.RemoveItem (List4.ListIndex)
  List5.RemoveItem (List5.ListIndex)
  List6.RemoveItem (List6.ListIndex)
  '2
 Form8.Data1.Recordset.MoveFirst
  Do
     If (Form8.Data1.Recordset.Fields!numberterm = s) And (Form8.Data1.Recordset.Fields!numberclass = q) Then
        Form8.Data1.Recordset.Delete
        Form8.Data1.Refresh
     End If
     Form8.Data1.Recordset.MoveNext
  Loop Until Form8.Data1.Recordset.EOF = True
  
  '3
  Form21.Data1.Recordset.MoveFirst
  Do
     If (Form21.Data1.Recordset.Fields!nt = s) And (Form21.Data1.Recordset.Fields!nc = q) Then
        Form21.Data1.Recordset.Delete
        Form21.Data1.Refresh
     End If
     Form21.Data1.Recordset.MoveNext
  Loop Until Form21.Data1.Recordset.EOF = True
  
  '4
  Form21.Data2.Recordset.MoveFirst
  Do
     If (Form21.Data2.Recordset.Fields!nt = s) And (Form21.Data2.Recordset.Fields!nc = q) Then
        Form21.Data2.Recordset.Delete
        Form21.Data2.Refresh
     End If
     Form21.Data2.Recordset.MoveNext
  Loop Until Form21.Data2.Recordset.EOF = True
  
  '5
  Form21.Data3.Recordset.MoveFirst
  Do
     If (Form21.Data3.Recordset.Fields!nt = s) And (Form21.Data3.Recordset.Fields!nc = q) Then
        Form21.Data3.Recordset.Delete
        Form21.Data3.Refresh
     End If
     Form21.Data3.Recordset.MoveNext
  Loop Until Form21.Data3.Recordset.EOF = True
  
  '6
  Form21.Data4.Recordset.MoveFirst
  Do
     If (Form21.Data4.Recordset.Fields!nt = s) And (Form21.Data4.Recordset.Fields!nc = q) Then
        Form21.Data4.Recordset.Delete
        Form21.Data4.Refresh
     End If
     Form21.Data4.Recordset.MoveNext
  Loop Until Form21.Data4.Recordset.EOF = True
  End If
End If
End Sub

Private Sub KewlButtons2_Click()
   Dim a As Boolean, j As String
   If Val(Text3.Text) > 25 Then Text3.Text = 25
   If Option1.Value = True Then
      a = False
      If Option3.Value = True Then k = 1
      If Option4.Value = True Then k = 2
      If a = False Then
         Data1.Recordset.AddNew
         Data1.Recordset.Fields!numberterm = List7.List(List1.ListIndex)
         Data1.Recordset.Fields!numberclass = Text1.Text
         Data1.Recordset.Fields!mozoe = Combo1.ListIndex
         Data1.Recordset.Fields!se = k
         Data1.Recordset.Fields!dabir = Text2.Text
         Data1.Recordset.Fields!plase = Text3.Text
         Data1.Recordset.Fields!Day = Text4.Text
         Data1.Recordset.Fields!hors = Text5.Text
         j = Trim(Str(k)) + "-" + Text1.Text + "-" + Trim(Str(Combo1.ListIndex))
         Data1.Recordset.Fields!idclass = j
         Data1.Recordset.Update
         Text1.SetFocus
      End If
   End If
   
   If Option2.Value = True Then
'
      Data1.Recordset.MoveFirst
      Do
         If (Data1.Recordset.Fields!numberterm = List7.List(List1.ListIndex)) And (Data1.Recordset.Fields!numberclass = List2.Text) Then
            Data1.Recordset.Edit
            Data1.Recordset.Fields!numberclass = Text1.Text
            Data1.Recordset.Fields!mozoe = Combo1.ListIndex
            Data1.Recordset.Fields!se = k
            Data1.Recordset.Fields!dabir = Text2.Text
            Data1.Recordset.Fields!plase = Text3.Text
            Data1.Recordset.Fields!Day = Text4.Text
            Data1.Recordset.Fields!hors = Text5.Text
            j = Trim(Str(k)) + "-" + Text1.Text + "-" + Trim(Str(Combo1.ListIndex))
            Data1.Recordset.Fields!idclass = j
            Data1.Recordset.Update
            Exit Do
         End If
         Data1.Recordset.MoveNext
      Loop Until Data1.Recordset.EOF = True
      Text1.SetFocus
   End If
End Sub

Private Sub List1_Click()
Dim z As Boolean
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List8.Clear
Data1.Recordset.MoveFirst
Do
  If (Data1.Recordset.Fields!numberterm = List7.List(List1.ListIndex)) Then
     
     If Option3.Value = True Then
       If Data1.Recordset.Fields!se = 1 Then
         z = True
         For q = 0 To List2.ListCount - 1
           If List2.List(q) = Data1.Recordset.Fields!numberclass Then z = False
         Next q
         If z = True Then List2.AddItem Data1.Recordset.Fields!numberclass
       End If
     End If
     
     If Option4.Value = True Then
       If Data1.Recordset.Fields!se = 2 Then
         z = True
         For q = 0 To List2.ListCount - 1
           If List2.List(q) = Data1.Recordset.Fields!numberclass Then z = False
         Next q
         If z = True Then List2.AddItem Data1.Recordset.Fields!numberclass
       End If
     End If
     
  End If
  Data1.Recordset.MoveNext
Loop Until Data1.Recordset.EOF = True
Label8.Caption = List2.ListCount
End Sub

Private Sub List2_Click()
Dim v As Integer
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List8.Clear
If Option3.Value = True Then v = 1
If Option4.Value = True Then v = 2
Data1.Recordset.MoveFirst
Do
  If (Data1.Recordset.Fields!numberclass = List2.Text) And (Data1.Recordset.Fields!se = v) And (Data1.Recordset.Fields!numberterm = List7.List(List1.ListIndex)) Then
    List3.AddItem Data1.Recordset.Fields!dabir
    List4.AddItem Data1.Recordset.Fields!plase
    List5.AddItem Data1.Recordset.Fields!Day
    List6.AddItem Data1.Recordset.Fields!hors
    Select Case Data1.Recordset.Fields!mozoe
      Case 0
        List8.AddItem "—Ì«÷Ì"
      Case 1
        List8.AddItem "⁄·Ê„"
      Case 2
        List8.AddItem "›Ì“Ìò"
      Case 3
        List8.AddItem "‘Ì„Ì"
      Case 4
        List8.AddItem "œÌ‰Ì"
      Case 5
        List8.AddItem "⁄—»Ì"
      Case 6
        List8.AddItem "“»«‰"
    End Select
  End If
  Data1.Recordset.MoveNext
Loop Until Data1.Recordset.EOF = True
End Sub

Private Sub List3_Click()
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
List8.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List3.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
List8.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
List8.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
List3.ListIndex = List6.ListIndex
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
List8.ListIndex = List6.ListIndex
End Sub

Private Sub List8_Click()
List3.ListIndex = List8.ListIndex
List4.ListIndex = List8.ListIndex
List5.ListIndex = List8.ListIndex
List6.ListIndex = List8.ListIndex
End Sub

Private Sub Option1_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Option2_Click()
If List8.ListIndex > -1 Then
  Text1.Text = List2.List(List2.ListIndex)
  Combo1.Text = List8.List(List8.ListIndex)
  Text2.Text = List3.List(List3.ListIndex)
  Text3.Text = List4.List(List4.ListIndex)
  Text4.Text = List5.List(List5.ListIndex)
  Text5.Text = List6.List(List6.ListIndex)
  Text1.SetFocus
Else
  e = MsgBox("·ÿ›« „Ê÷Ê⁄ ò·«” ŒÊœ —« „⁄Ì‰ ò‰Ìœ", vbCritical, "")
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo1.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
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

