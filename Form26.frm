VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form26 
   BorderStyle     =   0  'None
   Caption         =   "Form26"
   ClientHeight    =   9210
   ClientLeft      =   1365
   ClientTop       =   1200
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
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\3.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data2"
      RightToLeft     =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\3.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data1"
      RightToLeft     =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Height          =   360
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "·Ì”  „»«·€"
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
      MICON           =   "Form26.frx":0000
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
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "·Ì”  Õ÷Ê— €Ì«»"
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
      MICON           =   "Form26.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command9 
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form26.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ê“«—‘"
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
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„  —„ :"
      Height          =   375
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form26.frx":0054
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String, t As String

Private Sub Combo1_Click()
Dim q1(6, 50), q2(50, 24, 14), qwt(11), q3(50, 24, 24), t As String
Dim a1(5), q, w, e, r As Integer

PB1.Visible = True
PB1.Value = 0
'·Ì”  Õ÷Ê— €Ì«»

'Å«ò ò—œ‰
   Data1.Recordset.MoveFirst
   If Data1.Recordset.EOF = False Then
      Data1.Recordset.MoveFirst
      Do
         Data1.Recordset.Delete
         Data1.Refresh
         Data1.Recordset.MoveNext
      Loop Until Data1.Recordset.EOF = True
      Data1.Refresh
      Data1.Recordset.MoveFirst
      Data1.Recordset.Delete
      Data1.Refresh
   End If
      
   For q = 0 To 6
     For w = 0 To 50
       q1(q, w) = ""
     Next w
   Next q
      
   For q = 0 To 50
     For w = 0 To 24
       For e = 0 To 14
         q2(q, w, e) = ""
       Next e
     Next w
   Next q
   
   
PB1.Value = 10
'«‰ Œ«» ò·«” Â«
a1(0) = 0
Form6.Data1.Recordset.MoveFirst
Do
  If Form6.Data1.Recordset.Fields!numberterm = List1.List(Combo1.ListIndex) Then
     q1(0, a1(0)) = Form6.Data1.Recordset.Fields!numberclass
     t = ""
    Select Case Val(Form6.Data1.Recordset.Fields!mozoe)
      Case 0
        t = "—Ì«÷Ì"
      Case 1
        t = "⁄·Ê„"
      Case 2
        t = "›Ì“Ìò"
      Case 3
        t = "‘Ì„Ì"
      Case 4
        t = "œÌ‰Ì"
      Case 5
        t = "⁄—»Ì"
      Case 6
        t = "“»«‰"
    End Select
     q1(4, a1(0)) = t
     If Form6.Data1.Recordset.Fields!se = 1 Then q1(5, a1(0)) = "Å”—«‰"
     If Form6.Data1.Recordset.Fields!se = 2 Then q1(5, a1(0)) = "œŒ —«‰"
     
     q1(6, a1(0)) = Form6.Data1.Recordset.Fields!idclass
     q1(1, a1(0)) = Form6.Data1.Recordset.Fields!dabir
     q1(2, a1(0)) = Form6.Data1.Recordset.Fields!Day
     q1(3, a1(0)) = Form6.Data1.Recordset.Fields!hors
     a1(0) = a1(0) + 1
  End If
  Form6.Data1.Recordset.MoveNext
Loop Until Form6.Data1.Recordset.EOF = True

PB1.Value = 15
'«‰ Œ«» òœ œ«‰‘ ¬„Ê“«‰ Ê «ÿ·«⁄«  ò·«”Ì
a1(2) = 0
For a1(1) = 0 To a1(0) - 1
   a1(3) = 0
   Form8.Data1.Recordset.MoveFirst
   Do
      If (Form8.Data1.Recordset.Fields!numberterm = List1.List(Combo1.ListIndex)) And (Form8.Data1.Recordset.Fields!numberclass = q1(6, a1(1))) Then
        q2(a1(2), a1(3), 0) = Form8.Data1.Recordset.Fields!numberstudent
        q2(a1(2), a1(3), 6) = q1(0, a1(1))
        q2(a1(2), a1(3), 7) = q1(1, a1(1))
        q2(a1(2), a1(3), 8) = q1(2, a1(1))
        q2(a1(2), a1(3), 9) = q1(3, a1(1))
        q2(a1(2), a1(3), 12) = q1(4, a1(1))
        q2(a1(2), a1(3), 13) = q1(5, a1(1))
        q2(a1(2), a1(3), 14) = q1(6, a1(1))
        Select Case Right(List1.List(Combo1.ListIndex), 1)
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
        q2(a1(2), a1(3), 10) = p
        q2(a1(2), a1(3), 11) = Left(List1.List(Combo1.ListIndex), 1)
        a1(3) = a1(3) + 1
      End If
      Form8.Data1.Recordset.MoveNext
   Loop Until Form8.Data1.Recordset.EOF = True
   a1(2) = a1(2) + 1
Next a1(1)

PB1.Value = 40
'«‰ Œ«» ‰«„ œ«‰‘ ¬„Ê“«‰
For q = 0 To 50
  For w = 0 To 24
    Form2.Data1.Recordset.FindFirst "id='" & q2(q, w, 0) & "'"
    If Form2.Data1.Recordset.NoMatch = False Then
      q2(q, w, 1) = Form2.Data1.Recordset.Fields!Name
      q2(q, w, 2) = Form2.Data1.Recordset.Fields!family
    End If
  Next w
Next q

PB1.Value = 60
'«‰ Œ«» „œ—”Â œ«‰‘ ¬„Ê“«‰
For q = 0 To 50
  For w = 0 To 24
    Form3.Data1.Recordset.FindFirst "id='" & q2(q, w, 0) & "'"
    If Form3.Data1.Recordset.NoMatch = False Then
      q2(q, w, 3) = Form3.Data1.Recordset.Fields!Shift
      q2(q, w, 4) = Form3.Data1.Recordset.Fields!school
    End If
    q2(q, w, 5) = w + 1
  Next w
Next q

PB1.Value = 70
'„— » ò—œ‰ »— «”«” Õ—Ê› «·›»« ›Ì·œ ‰«„ Œ«‰Ê«œêÌ
For q = 0 To 50
  For w = 0 To 24
    For e = 0 To 24
      If (q2(q, e, 2) > q2(q, w, 2)) And (q2(q, w, 2) <> "") Then
        
        For r = 0 To 11
          qwt(r) = q2(q, e, r)
        Next r
            
        For r = 0 To 11
           q2(q, e, r) = q2(q, w, r)
        Next r
            
        For r = 0 To 11
           q2(q, w, r) = qwt(r)
        Next r
            
      End If
    Next e
  Next w
Next q

'⁄Ê÷ ò—œ‰ ‘„«—Â —œÌ›
For q = 0 To 50
  For w = 0 To 24
    q2(q, w, 5) = w + 1
  Next w
Next q

'”«⁄  Ê  «—ÌŒ »Â ’Ê—  ò«„·
For q = 0 To 50
  For w = 1 To 24
    q2(q, w, 8) = q2(q, 0, 8)
    q2(q, w, 9) = q2(q, 0, 8)
  Next w
Next q

PB1.Value = 85
'‰Ê‘ ‰ œ— ›«Ì·
For q = 0 To 50
  If q2(q, 0, 0) <> "" Then
    For w = 0 To 24
       Data1.Recordset.AddNew
       Data1.Recordset.Fields!rad = q2(q, w, 5)
       Data1.Recordset.Fields!id = q2(q, w, 0)
       Data1.Recordset.Fields!Name = q2(q, w, 1) + " " + q2(q, w, 2)
       Data1.Recordset.Fields!md = q2(q, w, 4)
       Data1.Recordset.Fields!shi = q2(q, w, 3)
       Data1.Recordset.Fields!idterm = q2(q, w, 11)
       Data1.Recordset.Fields!nameterm = q2(q, w, 10)
       Data1.Recordset.Fields!dabir = q2(q, w, 7)
       Data1.Recordset.Fields!Day = q2(q, w, 8)
       Data1.Recordset.Fields!Time = q2(q, w, 9)
       Data1.Recordset.Fields!Class = q2(q, w, 6)
       Data1.Recordset.Fields!mo = q2(q, w, 12)
       Data1.Recordset.Fields!se = q2(q, w, 13)
       Data1.Recordset.Update
    Next w
  End If
Next q
PB1.Value = 100


'·Ì”  „»«·€

'Å«ò ò—œ‰
   Data2.Recordset.MoveFirst
   If Data2.Recordset.EOF = False Then
      Data2.Recordset.MoveFirst
      Do
         Data2.Recordset.Delete
         Data2.Refresh
         Data2.Recordset.MoveNext
      Loop Until Data2.Recordset.EOF = True
      Data2.Refresh
      Data2.Recordset.MoveFirst
      Data2.Recordset.Delete
      Data2.Refresh
   End If
      
   For q = 0 To 50
     For w = 0 To 24
       For e = 0 To 22
         q3(q, w, e) = ""
       Next e
     Next w
   Next q

PB1.Value = 10
'«‰ ﬁ«· »—«Ì ·Ì”  „»«·€
For q = 0 To 50
  For w = 0 To 24
    For e = 0 To 11
      q3(q, w, e) = q2(q, w, e)
    Next e
    q3(q, w, 23) = q2(q, w, 12)
    q3(q, w, 24) = q2(q, w, 13)
  Next w
Next q

PB1.Value = 30
'„»·€ Â— œÊ—Â
Form5.Data1.Recordset.FindFirst "numberterm='" & List1.List(Combo1.ListIndex) & "'"
t = Form5.Data1.Recordset.Fields!Money

PB1.Value = 35
'Å—œ«“‘ „»«·€ »—«Ì Â— ›—œ
For q = 0 To 50
  For w = 0 To 24
     If q3(q, w, 0) <> "" Then
       Form4.Data1.Recordset.MoveFirst
       Do
         If (q3(q, w, 0) = Form4.Data1.Recordset.Fields!id) And (Form4.Data1.Recordset.Fields!numberterm = List1.List(Combo1.ListIndex)) Then
           If Form4.Data1.Recordset.Fields!Number = 1 Then
             q3(q, w, 12) = Form4.Data1.Recordset.Fields!padakht
             q3(q, w, 13) = Form4.Data1.Recordset.Fields!numbergabz
           End If
         
           If Form4.Data1.Recordset.Fields!Number = 2 Then
             q3(q, w, 14) = Form4.Data1.Recordset.Fields!padakht
             q3(q, w, 15) = Form4.Data1.Recordset.Fields!numbergabz
           End If
         
           If Form4.Data1.Recordset.Fields!Number = 3 Then
             q3(q, w, 16) = Form4.Data1.Recordset.Fields!padakht
             q3(q, w, 17) = Form4.Data1.Recordset.Fields!numbergabz
           End If
         End If
         Form4.Data1.Recordset.MoveNext
       Loop Until Form4.Data1.Recordset.EOF = True
       
       If q3(q, w, 12) = "" Then
         q3(q, w, 12) = 0
         q3(q, w, 13) = "-"
       End If
       
       If q3(q, w, 14) = "" Then
         q3(q, w, 14) = 0
         q3(q, w, 15) = "-"
       End If
       
       If q3(q, w, 16) = "" Then
         q3(q, w, 16) = 0
         q3(q, w, 17) = "-"
       End If
       
       Form4.Data2.Recordset.MoveFirst
       Do
          If (Form4.Data2.Recordset.Fields!id = q3(q, w, 0)) And (Form4.Data2.Recordset.Fields!numberterm) Then
            q3(q, w, 18) = Form4.Data2.Recordset.Fields!takhfif
          Else
            q3(q, w, 18) = 0
          End If
          Form4.Data2.Recordset.MoveNext
       Loop Until Form4.Data2.Recordset.EOF = True
        
       q3(q, w, 19) = (Val(t)) - (Val(q3(q, w, 18)) + Val(q3(q, w, 16)) + Val(q3(q, w, 14)) + Val(q3(q, w, 12)))
     End If
  Next w
Next q

PB1.Value = 75
'‘„«—Â Â«Ì  ·›‰
For q = 0 To 50
  For w = 0 To 24
    Form2.Data1.Recordset.FindFirst "id='" & q3(q, w, 0) & "'"
    If Form2.Data1.Recordset.NoMatch = False Then
      q3(q, w, 20) = Form2.Data1.Recordset.Fields!phone
      q3(q, w, 21) = Form2.Data1.Recordset.Fields!r1
      q3(q, w, 22) = Form2.Data1.Recordset.Fields!r2
    End If
  Next w
Next q


PB1.Value = 85
'‰Ê‘ ‰ œ— ›«Ì·
For q = 0 To 50
  If q3(q, 0, 0) <> "" Then
    For w = 0 To 24
       Data2.Recordset.AddNew
       Data2.Recordset.Fields!rad = q3(q, w, 5)
       Data2.Recordset.Fields!id = q3(q, w, 0)
       Data2.Recordset.Fields!Name = q3(q, w, 1) + " " + q2(q, w, 2)
       Data2.Recordset.Fields!md = q3(q, w, 4)
       Data2.Recordset.Fields!shi = q3(q, w, 3)
       Data2.Recordset.Fields!idterm = q3(q, w, 11)
       Data2.Recordset.Fields!nameterm = q3(q, w, 10)
       Data2.Recordset.Fields!dabir = q3(q, w, 7)
       Data2.Recordset.Fields!Day = q3(q, w, 8)
       Data2.Recordset.Fields!Time = q3(q, w, 9)
       Data2.Recordset.Fields!Class = q3(q, w, 6)
       Data2.Recordset.Fields!mo = q3(q, w, 23)
       Data2.Recordset.Fields!se = q3(q, w, 24)
       
       Data2.Recordset.Fields!mon1 = q3(q, w, 12)
       Data2.Recordset.Fields!gabz1 = q3(q, w, 13)
       Data2.Recordset.Fields!mon2 = q3(q, w, 14)
       Data2.Recordset.Fields!gabz2 = q3(q, w, 15)
       Data2.Recordset.Fields!mon3 = q3(q, w, 16)
       Data2.Recordset.Fields!gabz3 = q3(q, w, 17)
       
       Data2.Recordset.Fields!takh = q3(q, w, 18)
       Data2.Recordset.Fields!bag = q3(q, w, 19)
       Data2.Recordset.Fields!t1 = q3(q, w, 20)
       Data2.Recordset.Fields!t2 = q3(q, w, 21)
       Data2.Recordset.Fields!t3 = q3(q, w, 22)
       
       Data2.Recordset.Update
    Next w
  End If
Next q
PB1.Value = 100

PB1.Visible = False
End Sub


Private Sub Command9_Click()
Form26.Hide
Form15.Show
End Sub

Private Sub Form_Activate()
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
     Combo1.AddItem "  —„ " + Left(Form5.Data1.Recordset.Fields!numberterm, Len(Form5.Data1.Recordset.Fields!numberterm) - 2) + " Å«ÌÂ  " + p
     List1.AddItem Form5.Data1.Recordset.Fields!numberterm
  End If
  Form5.Data1.Recordset.MoveNext
Loop Until Form5.Data1.Recordset.EOF = True
End Sub

