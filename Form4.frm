VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
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
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List9 
      Height          =   360
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin KewlButtonz.KewlButtons KewlButtons7 
      Height          =   375
      Left            =   1080
      TabIndex        =   39
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "Form4.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons6 
      Height          =   375
      Left            =   2040
      TabIndex        =   38
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "Form4.frx":001C
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
      Height          =   135
      Left            =   4320
      TabIndex        =   32
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form4.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List8 
      Height          =   2160
      ItemData        =   "Form4.frx":0054
      Left            =   9840
      List            =   "Form4.frx":0056
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9840
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3840
      Width           =   1335
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   10320
      TabIndex        =   26
      Top             =   7920
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
      MICON           =   "Form4.frx":0058
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons command1 
      Height          =   375
      Left            =   3000
      TabIndex        =   25
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "«÷«›Â"
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
      MICON           =   "Form4.frx":0074
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      Height          =   4860
      ItemData        =   "Form4.frx":0090
      Left            =   8280
      List            =   "Form4.frx":0092
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   4860
      ItemData        =   "Form4.frx":0094
      Left            =   6960
      List            =   "Form4.frx":0096
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   4860
      ItemData        =   "Form4.frx":0098
      Left            =   5640
      List            =   "Form4.frx":009A
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   4860
      ItemData        =   "Form4.frx":009C
      Left            =   4320
      List            =   "Form4.frx":009E
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "Form4.frx":00A0
      Left            =   9720
      List            =   "Form4.frx":00A2
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ListBox List5 
      Height          =   2160
      ItemData        =   "Form4.frx":00A4
      Left            =   3480
      List            =   "Form4.frx":00A6
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   495
   End
   Begin VB.ListBox List6 
      Height          =   2160
      ItemData        =   "Form4.frx":00A8
      Left            =   2280
      List            =   "Form4.frx":00AA
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List7 
      Height          =   2160
      ItemData        =   "Form4.frx":00AC
      Left            =   1080
      List            =   "Form4.frx":00AE
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      RecordSource    =   "data4"
      RightToLeft     =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "data3"
      RightToLeft     =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
      Height          =   135
      Left            =   5640
      TabIndex        =   33
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form4.frx":00B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   135
      Left            =   6960
      TabIndex        =   34
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form4.frx":00CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons5 
      Height          =   135
      Left            =   8280
      TabIndex        =   35
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form4.frx":00E8
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
      Caption         =   "œ—Ì«›  „»«·€ œ«‰‘ ¬„Ê“«‰"
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
      TabIndex        =   42
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„  —„ Â« Ê Å«ÌÂ Â«"
      Height          =   375
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   11640
      X2              =   9720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line4 
      X1              =   9720
      X2              =   11640
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      X1              =   9720
      X2              =   11640
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   11640
      X2              =   11640
      Y1              =   7680
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   9720
      X2              =   9720
      Y1              =   3600
      Y2              =   7680
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ :"
      Height          =   375
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   375
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ"
      Height          =   375
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      Height          =   375
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "òœ "
      Height          =   375
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Åœ—"
      Height          =   375
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ò· „»·€ Å—œ«Œ Ì"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ œ—Ì«› Ì"
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ﬁ»÷"
      Height          =   375
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»«ﬁÌ„«‰œÂ"
      Height          =   375
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€  Œ›Ì›"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "⁄·   Œ›Ì›"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘Œ’  Œ›Ì›"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   9330
      Left            =   0
      Picture         =   "Form4.frx":0104
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w As String, p As Long
Dim sd, sa As String
Dim z As Long, x As Long, c As Long, v As Boolean
Private Sub search12()
Label6.Caption = ""
Text6.Text = ""
Text1.Text = ""
Text2.Text = ""
List5.Clear
List6.Clear
List7.Clear
Form5.Data1.Recordset.FindFirst "numberterm='" & List9.List(Combo1.ListIndex) & "'"
If Form5.Data1.Recordset.NoMatch = False Then
   Label6.Caption = Form5.Data1.Recordset.Fields!Money
End If

Form4.Data2.Recordset.MoveFirst
Do
   If (Form4.Data2.Recordset.Fields!id = List1.List(List1.ListIndex)) And (Form4.Data2.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)) Then
      Text6.Text = Form4.Data2.Recordset.Fields!takhfif
      Text1.Text = Form4.Data2.Recordset.Fields!ellat
      Text2.Text = Form4.Data2.Recordset.Fields!men
      Exit Do
   End If
   Form4.Data2.Recordset.MoveNext
Loop Until Form4.Data2.Recordset.EOF = True
Data1.Recordset.MoveFirst
Do
   If (Data1.Recordset.Fields!id = List1.List(List1.ListIndex)) And (Form4.Data1.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)) Then
      List5.AddItem Data1.Recordset.Fields!Number
      List6.AddItem Data1.Recordset.Fields!padakht
      List7.AddItem Data1.Recordset.Fields!numbergabz
   End If
   Data1.Recordset.MoveNext
Loop Until Data1.Recordset.EOF = True
p = 0
For m = 0 To List6.ListCount - 1
  p = p + List6.List(m)
Next m
p = p + Val(Text6.Text)
Label10.Caption = Val(Label6.Caption) - p
End Sub

Private Sub Command1_Click()
If Label10.Caption > 0 Then
   z = Combo1.ListIndex
   c = List1.ListIndex
   Form10.Label4.Caption = "1"
   Form10.Show vbModal
Else
   q = MsgBox("»«ﬁÌ „«‰œÂ «Ì‰ œ«‰‘ ¬„Ê“ 0 „Ì »«‘œ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Combo1_Click()
Dim l As Boolean
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Form8.Data1.Recordset.MoveFirst
Do
   If (Form8.Data1.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)) Then
      l = False
      For k = 0 To List1.ListCount - 1
        If List1.List(k) = Form8.Data1.Recordset.Fields!numberstudent Then
           l = True
           Exit For
        End If
      Next k
         If l = False Then List1.AddItem Form8.Data1.Recordset.Fields!numberstudent
   End If
   Form8.Data1.Recordset.MoveNext
Loop Until Form8.Data1.Recordset.EOF = True
   For m = 0 To List1.ListCount - 1
      Form2.Data1.Recordset.MoveFirst
      Do
         If (Form2.Data1.Recordset.Fields!id = List1.List(m)) Then
            List2.AddItem Form2.Data1.Recordset.Fields!Name
            List3.AddItem Form2.Data1.Recordset.Fields!family
            List4.AddItem Form2.Data1.Recordset.Fields!fathername
            Exit Do
         End If
         Form2.Data1.Recordset.MoveNext
      Loop Until Form2.Data1.Recordset.EOF = True
   Next m
Label19.Caption = List1.ListCount
End Sub

Private Sub Form_Activate()
qw = Combo1.ListIndex
Combo1.Clear
List9.Clear
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
     List9.AddItem Form5.Data1.Recordset.Fields!numberterm
  End If
  Form5.Data1.Recordset.MoveNext
Loop Until Form5.Data1.Recordset.EOF = True

Combo1.ListIndex = qw
KewlButtons6.Enabled = False
KewlButtons7.Enabled = False
End Sub

Private Sub Form_Load()
c = -1
End Sub

Private Sub KewlButtons1_Click()
Combo1.Clear
Form5.Data1.Recordset.MoveFirst
Do
  If Form5.Data1.Recordset.Fields!numberterm <> 0 Then
     Combo1.AddItem Form5.Data1.Recordset.Fields!nameterm
  End If
  Form5.Data1.Recordset.MoveNext
Loop Until Form5.Data1.Recordset.EOF = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (List1.List(List1.ListIndex) <> "") And (Label6.Caption <> "") Then
   Command1.Enabled = True
Else
   Command1.Enabled = False
End If
End Sub

Private Sub KewlButtons2_Click()
Dim id(500), na(500), fa(500), md(500), idt, nat, fat, mdt, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    fa(intq) = List3.List(intq)
    md(intq) = List4.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If md(intq) > md(intw) Then
         idt = id(intq)
         nat = na(intq)
         fat = fa(intq)
         mdt = md(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         fa(intq) = fa(intw)
         md(intq) = md(intw)
         
         id(intw) = idt
         na(intw) = nat
         fa(intw) = fat
         md(intw) = mdt
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
List4.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem fa(intq)
   List4.AddItem md(intq)
Next intq
End Sub

Private Sub KewlButtons3_Click()
Dim id(500), na(500), fa(500), md(500), idt, nat, fat, mdt, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    fa(intq) = List3.List(intq)
    md(intq) = List4.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If fa(intq) > fa(intw) Then
         idt = id(intq)
         nat = na(intq)
         fat = fa(intq)
         mdt = md(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         fa(intq) = fa(intw)
         md(intq) = md(intw)
         
         id(intw) = idt
         na(intw) = nat
         fa(intw) = fat
         md(intw) = mdt
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
List4.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem fa(intq)
   List4.AddItem md(intq)
Next intq
End Sub

Private Sub KewlButtons4_Click()
Dim id(500), na(500), fa(500), md(500), idt, nat, fat, mdt, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    fa(intq) = List3.List(intq)
    md(intq) = List4.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         idt = id(intq)
         nat = na(intq)
         fat = fa(intq)
         mdt = md(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         fa(intq) = fa(intw)
         md(intq) = md(intw)
         
         id(intw) = idt
         na(intw) = nat
         fa(intw) = fat
         md(intw) = mdt
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
List4.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem fa(intq)
   List4.AddItem md(intq)
Next intq

End Sub

Private Sub KewlButtons5_Click()
Dim id(500), na(500), fa(500), md(500), idt, nat, fat, mdt, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    fa(intq) = List3.List(intq)
    md(intq) = List4.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If id(intq) > id(intw) Then
         idt = id(intq)
         nat = na(intq)
         fat = fa(intq)
         mdt = md(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         fa(intq) = fa(intw)
         md(intq) = md(intw)
         
         id(intw) = idt
         na(intw) = nat
         fa(intw) = fat
         md(intw) = mdt
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
List4.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem fa(intq)
   List4.AddItem md(intq)
Next intq
End Sub

Private Sub KewlButtons6_Click()
If List5.List(0) <> "" Then
   z = Combo1.ListIndex
   c = List1.ListIndex
   Form10.Text1.Text = List5.List(List5.ListIndex)
   Form10.Text2.Text = List6.List(List6.ListIndex)
   Form10.Text3.Text = List7.List(List7.ListIndex)
   Form10.Label4.Caption = "2"
   Form10.Show vbModal
Else
   q = MsgBox("»«ﬁÌ „«‰œÂ «Ì‰ œ«‰‘ ¬„Ê“ 0 „Ì »«‘œ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons7_Click()
Form4.Data1.Recordset.MoveFirst
Do
   If (Form4.Data1.Recordset.Fields!id = Form4.List1.List(Form4.List1.ListIndex)) And (Form4.Data1.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)) And (Form4.Data1.Recordset.Fields!Number = Form4.List5.List(Form4.List5.ListIndex)) And (Form4.Data1.Recordset.Fields!padakht = Form4.List6.List(Form4.List6.ListIndex)) And (Form4.Data1.Recordset.Fields!numbergabz = Form4.List7.List(Form4.List7.ListIndex)) Then
      Form4.Data1.Recordset.Delete
      Form4.Data1.Refresh
      Exit Do
   End If
   Form4.Data1.Recordset.MoveNext
Loop Until Form4.Data1.Recordset.EOF = True
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
Call search12
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List6.ListIndex = List5.ListIndex
List7.ListIndex = List5.ListIndex
KewlButtons6.Enabled = True
KewlButtons7.Enabled = True
End Sub

Private Sub List6_Click()
List5.ListIndex = List6.ListIndex
List7.ListIndex = List6.ListIndex
End Sub

Private Sub List7_Click()
List5.ListIndex = List7.ListIndex
List6.ListIndex = List7.ListIndex
End Sub

Private Sub List8_Click()
For q = 0 To List3.ListCount - 1
  If List3.List(q) = List8.Text Then List3.ListIndex = q
Next q
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For k = 0 To List1.ListCount - 1
     If List1.List(k) = Text3.Text Then
        List1.ListIndex = k
        Exit For
     End If
   Next k
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   List8.Clear
   For q = 0 To List3.ListCount - 1
     intq = InStr(List3.List(q), Text4.Text)
     If intq <> 0 Then
        List8.AddItem List3.List(q)
      End If
   Next q
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If List1.List(List1.ListIndex) Then
      v = False
      Form4.Data2.Recordset.MoveFirst
      Do
         If Form4.Data2.Recordset.Fields!id = List1.List(List1.ListIndex) Then
            Form4.Data2.Recordset.Edit
            Form4.Data2.Recordset.Fields!takhfif = Text6.Text
            Form4.Data2.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)
            Form4.Data2.Recordset.Fields!ellat = Text1.Text
            Form4.Data2.Recordset.Fields!men = Text2.Text
            Form4.Data2.Recordset.Update
            v = True
            Exit Do
         End If
      Form4.Data2.Recordset.MoveNext
      Loop Until Form4.Data2.Recordset.EOF = True
      If v = False Then
         Form4.Data2.Recordset.AddNew
         Form4.Data2.Recordset.Fields!id = List1.List(List1.ListIndex)
         Form4.Data2.Recordset.Fields!takhfif = Text6.Text
         Form4.Data2.Recordset.Fields!numberterm = List9.List(Combo1.ListIndex)
         Form4.Data2.Recordset.Fields!ellat = Text1.Text
         Form4.Data2.Recordset.Fields!men = Text2.Text
         Form4.Data2.Recordset.Update
      End If
      p = 0
      For m = 0 To List6.ListCount - 1
        p = p + List6.List(m)
      Next m
      p = p + Val(Text6.Text)
      Label10.Caption = Val(Label6.Caption) - p
   End If
End If
End Sub

