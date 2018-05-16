VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   9210
   ClientLeft      =   1140
   ClientTop       =   945
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List4 
      Height          =   1260
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin KewlButtonz.KewlButtons KewlButtons5 
      Height          =   375
      Left            =   1080
      TabIndex        =   46
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "—› ‰ »Â »Œ‘ À»  ”«·  Õ’Ì·Ì"
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
      MICON           =   "Form2.frx":0000
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
      Height          =   375
      Left            =   3360
      TabIndex        =   45
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "Form2.frx":001C
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
      Left            =   9240
      TabIndex        =   42
      Top             =   7440
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
      MICON           =   "Form2.frx":0038
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
      Left            =   9960
      TabIndex        =   41
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "Form2.frx":0054
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
      Left            =   8160
      TabIndex        =   40
      Top             =   5400
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
      MICON           =   "Form2.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "Õ–›"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   6720
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "ÃœÌœ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   5880
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00709FC5&
      Caption         =   "ÊÌ—«Ì‘"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   9960
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   1860
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   1560
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5880
      Width           =   1695
   End
   Begin KewlButtonz.KewlButtons command2 
      Height          =   375
      Left            =   9840
      TabIndex        =   28
      Top             =   8160
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
      MICON           =   "Form2.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4680
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "DATA BASE INFORMATION\1.mdb"
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
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À»  «ÿ·«⁄«  œ«‰‘ ¬„Ê“«‰"
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
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ ò· œ«‰‘ ¬„Ê“«‰ :"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   11040
      X2              =   8040
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      X1              =   8040
      X2              =   8040
      Y1              =   2640
      Y2              =   7680
   End
   Begin VB.Line Line2 
      X1              =   11040
      X2              =   11040
      Y1              =   7680
      Y2              =   2640
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   11040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰ ÌÃÂ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "òœ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Õ’Ì·«  „«œ—"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Õ’Ì·«  Åœ—"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘€· Åœ—"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â  „«”"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Å” Ì"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Åœ—"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰‘«‰Ì"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ —«»ÿ 2"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ —«»ÿ 1"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  Ê·œ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ‘‰«”‰«„Â"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "òœ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form2.frx":00A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Long, w As Boolean, e As Integer

Private Sub Command2_Click()
Form2.Hide
End Sub

Private Sub Form_Activate()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
List4.Visible = False
List1.Clear
List2.Clear
List3.Clear
Text15.Text = ""
Text16.Text = ""
Form2.Data1.Recordset.MoveFirst
Do
   List1.AddItem Form2.Data1.Recordset.Fields!id
   List2.AddItem Form2.Data1.Recordset.Fields!Name + "  " + Form2.Data1.Recordset.Fields!family
   Form2.Data1.Recordset.MoveNext
Loop Until Form2.Data1.Recordset.EOF = True
For intr = 0 To List1.ListCount - 1
    If List1.List(intr) = 0 Then
       List1.RemoveItem (intr)
       List2.RemoveItem (intr)
       Exit For
    End If
Next intr
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label19.Caption = List1.ListCount
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
      If na(intq) > na(intw) Then
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

Private Sub KewlButtons2_Click()
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

Private Sub KewlButtons4_Click()
   Dim blnq As Boolean
   If Option1.Value = True Then
      If (Len(Text1.Text) < 6) Or (Len(Text1.Text) > 6) Then
         e = MsgBox("«Ì‰ òœ ‰«’ÕÌÕ «” ", vbCritical + vbMsgBoxRight, "")
         Text1.SetFocus
      Else
         blnq = False
         For q = 0 To List1.ListCount - 1
           If Text1.Text = List1.List(q) Then blnq = True
         Next q
         If blnq = True Then
            e = MsgBox("òœ Ê«—œ ‘œÂ  ò—«—Ì „Ì »«‘œ", vbCritical + vbMsgBoxRight, "")
         Else
            Data1.Recordset.AddNew
            Data1.Recordset.Fields!id = Text1.Text
            Data1.Recordset.Fields!Name = Text2.Text
            Data1.Recordset.Fields!family = Text3.Text
            Data1.Recordset.Fields!fathername = Text4.Text
            Data1.Recordset.Fields!shsh = Text5.Text
            Data1.Recordset.Fields!birthdate = Text6.Text
            Data1.Recordset.Fields!address = Text7.Text
            Data1.Recordset.Fields!idp = Text8.Text
            Data1.Recordset.Fields!phone = Text9.Text
            Data1.Recordset.Fields!r1 = Text10.Text
            Data1.Recordset.Fields!r2 = Text11.Text
            Data1.Recordset.Fields!fatehertask = Text12.Text
            Data1.Recordset.Fields!fd = Text13.Text
            Data1.Recordset.Fields!md = Text14.Text
            Data1.Recordset.Update
            List1.AddItem Text1.Text
            List2.AddItem Text2.Text + "  " + Text3.Text
            Text1.Text = Val(Text1.Text) + 1
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text2.SetFocus
         End If
      End If
   End If
   
   If Option2.Value = True Then
      Data1.Recordset.FindFirst "id='" & Text1.Text & "'"
      Data1.Recordset.Edit
      Data1.Recordset.Fields!Name = Text2.Text
      Data1.Recordset.Fields!family = Text3.Text
      Data1.Recordset.Fields!fathername = Text4.Text
      Data1.Recordset.Fields!shsh = Text5.Text
      Data1.Recordset.Fields!birthdate = Text6.Text
      Data1.Recordset.Fields!address = Text7.Text
      Data1.Recordset.Fields!idp = Text8.Text
      Data1.Recordset.Fields!phone = Text9.Text
      Data1.Recordset.Fields!r1 = Text10.Text
      Data1.Recordset.Fields!r2 = Text11.Text
      Data1.Recordset.Fields!fatehertask = Text12.Text
      Data1.Recordset.Fields!fd = Text13.Text
      Data1.Recordset.Fields!md = Text14.Text
      v = List2.ListIndex
      List2.RemoveItem v
      List2.AddItem Text2.Text + "  " + Text3.Text, v
      Data1.Recordset.Update
   End If
   
   If Option3.Value = True Then
      '1
      Data1.Recordset.FindFirst "id='" & Text1.Text & "'"
      Data1.Recordset.Delete
      Data1.Refresh
      List1.RemoveItem (List1.ListIndex)
      List2.RemoveItem (List2.ListIndex)
      Text1.Locked = False
      
      '2
      Form3.Data1.Recordset.MoveFirst
      Do
         If Form3.Data1.Recordset.Fields!id = Text1.Text Then
            Form3.Data1.Recordset.Delete
            Form3.Data1.Refresh
         End If
         Form3.Data1.Recordset.MoveNext
      Loop Until Form3.Data1.Recordset.EOF = True
      
      '3
      Form4.Data1.Recordset.MoveFirst
      Do
         If Form4.Data1.Recordset.Fields!id = Text1.Text Then
            Form4.Data1.Recordset.Delete
            Form4.Data1.Refresh
         End If
         Form4.Data1.Recordset.MoveNext
      Loop Until Form4.Data1.Recordset.EOF = True
      
      '4
      Form4.Data2.Recordset.MoveFirst
      Do
         If Form4.Data2.Recordset.Fields!id = Text1.Text Then
            Form4.Data2.Recordset.Delete
            Form4.Data2.Refresh
         End If
         Form4.Data2.Recordset.MoveNext
      Loop Until Form4.Data2.Recordset.EOF = True
      
      '5
      Form8.Data1.Recordset.MoveFirst
      Do
         If Form8.Data1.Recordset.Fields!numberstudent = Text1.Text Then
            Form8.Data1.Recordset.Delete
            Form8.Data1.Refresh
         End If
         Form8.Data1.Recordset.MoveNext
      Loop Until Form8.Data1.Recordset.EOF = True
   
      '6
      Form21.Data1.Recordset.MoveFirst
      Do
         If Form21.Data1.Recordset.Fields!id = Text1.Text Then
            Form21.Data1.Recordset.Delete
            Form21.Data1.Refresh
         End If
         Form21.Data1.Recordset.MoveNext
      Loop Until Form21.Data1.Recordset.EOF = True
   
      '7
      Form21.Data2.Recordset.MoveFirst
      Do
         If Form21.Data2.Recordset.Fields!id = Text1.Text Then
            Form21.Data2.Recordset.Delete
            Form21.Data2.Refresh
         End If
         Form21.Data2.Recordset.MoveNext
      Loop Until Form21.Data2.Recordset.EOF = True
   
      '8
      Form21.Data3.Recordset.MoveFirst
      Do
         If Form21.Data3.Recordset.Fields!id = Text1.Text Then
            Form21.Data3.Recordset.Delete
            Form21.Data3.Refresh
         End If
         Form21.Data3.Recordset.MoveNext
      Loop Until Form21.Data3.Recordset.EOF = True
   
      '9
      Form21.Data4.Recordset.MoveFirst
      Do
         If Form21.Data4.Recordset.Fields!id = Text1.Text Then
            Form21.Data4.Recordset.Delete
            Form21.Data4.Refresh
         End If
         Form21.Data4.Recordset.MoveNext
      Loop Until Form21.Data4.Recordset.EOF = True
      Text1.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
      Text5.Text = ""
      Text6.Text = ""
      Text7.Text = ""
      Text8.Text = ""
      Text9.Text = ""
      Text10.Text = ""
      Text11.Text = ""
      Text12.Text = ""
      Text13.Text = ""
      Text14.Text = ""
   End If
End Sub

Private Sub KewlButtons5_Click()
Form2.Hide
Form3.Show
Form3.Text6.Text = List1.Text
SendKeys ("{enter}")
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  
  If (List4.Left = 1080) And (List4.Top = 3240) Then
    Text2.Text = List4.Text
    Text2.SetFocus
    List4.Visible = False
  End If

  If (List4.Left = 4560) And (List4.Top = 3840) Then
    Text3.Text = List4.Text
    Text3.SetFocus
    List4.Visible = False
  End If

  If (List4.Left = 1080) And (List4.Top = 3840) Then
    Text4.Text = List4.Text
    Text4.SetFocus
    List4.Visible = False
  End If

  If (List4.Left = 4560) And (List4.Top = 6840) Then
    Text12.Text = List4.Text
    Text12.SetFocus
    List4.Visible = False
  End If

  If (List4.Left = 1080) And (List4.Top = 6840) Then
    Text13.Text = List4.Text
    Text13.SetFocus
    List4.Visible = False
  End If

  If (List4.Left = 4560) And (List4.Top = 7440) Then
    Text14.Text = List4.Text
    Text14.SetFocus
    List4.Visible = False
  End If

End If
End Sub

Private Sub List4_LostFocus()
List4.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Locked = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""

'
'sort
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
'
Dim blna As Boolean, a, b As Integer, s, d As Long
a = MsgBox("Ã‰”Ì  —« Ê«—œ ‰„«ÌÌœ " + Chr(10) + Chr(13) + "yes=Å”—" + Chr(10) + Chr(13) + "no=œŒ —", vbInformation + vbMsgBoxRight + vbYesNo, "")
blna = False
If a = 6 Then
   s = 100000
   For q = 0 To List1.ListCount - 1
       If Left(List1.List(q), 1) = 2 Then blna = True
       If blna = True Then b = q - 1: Exit For
   Next q
   d = List1.List(b)
Else
   s = 200000
   d = List1.List(List1.ListCount - 1)
End If

For q = s To d
  Data1.Recordset.FindFirst "id='" & q & "'"
  If Data1.Recordset.NoMatch = True Then Text1.Text = q
Next q
If Text1.Text = "" Then
   Text1.Text = d + 1
End If
Text1.SetFocus
End Sub

Private Sub Option2_Click()
If List1.ListIndex <> -1 Then
   Text1.Text = List1.Text
   Data1.Recordset.FindFirst "id='" & Text1.Text & "'"
   Text2.Text = Data1.Recordset.Fields!Name
   Text3.Text = Data1.Recordset.Fields!family
   Text4.Text = Data1.Recordset.Fields!fathername
   Text5.Text = Data1.Recordset.Fields!shsh
   Text6.Text = Data1.Recordset.Fields!birthdate
   Text7.Text = Data1.Recordset.Fields!address
   Text8.Text = Data1.Recordset.Fields!idp
   Text9.Text = Data1.Recordset.Fields!phone
   Text10.Text = Data1.Recordset.Fields!r1
   Text11.Text = Data1.Recordset.Fields!r2
   Text12.Text = Data1.Recordset.Fields!fatehertask
   Text13.Text = Data1.Recordset.Fields!fd
   Text14.Text = Data1.Recordset.Fields!md
   Text2.SetFocus
End If
End Sub

Private Sub Option3_Click()
If List1.ListIndex <> -1 Then
   Text1.Locked = True
   Text1.Text = List1.Text
   Data1.Recordset.FindFirst "id='" & Text1.Text & "'"
   Text2.Text = Data1.Recordset.Fields!Name
   Text3.Text = Data1.Recordset.Fields!family
   Text4.Text = Data1.Recordset.Fields!fathername
   Text5.Text = Data1.Recordset.Fields!shsh
   Text6.Text = Data1.Recordset.Fields!birthdate
   Text7.Text = Data1.Recordset.Fields!address
   Text8.Text = Data1.Recordset.Fields!idp
   Text9.Text = Data1.Recordset.Fields!phone
   Text10.Text = Data1.Recordset.Fields!r1
   Text11.Text = Data1.Recordset.Fields!r2
   Text12.Text = Data1.Recordset.Fields!fatehertask
   Text13.Text = Data1.Recordset.Fields!fd
   Text14.Text = Data1.Recordset.Fields!md
   Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub


Private Sub Text12_Change()
List4.Visible = False
If (Len(Text12.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 4560
   List4.Top = 6840
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If Left(Data1.Recordset.Fields!fatehertask, Len(Text12.Text)) = Text12.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!fatehertask Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!fatehertask
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text13_Change()
List4.Visible = False
If (Len(Text13.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 1080
   List4.Top = 6840
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If Left(Data1.Recordset.Fields!fd, Len(Text13.Text)) = Text13.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!fd Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!fd
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text14_Change()
List4.Visible = False
If (Len(Text14.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 4560
   List4.Top = 7440
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If Left(Data1.Recordset.Fields!md, Len(Text14.Text)) = Text14.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!md Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!md
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List1.ListCount - 1
     If Text15.Text = List1.List(q) Then List1.ListIndex = q
   Next q
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

Private Sub Text2_Change()
List4.Visible = False
If (Len(Text2.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 1080
   List4.Top = 3240
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If (Left(Data1.Recordset.Fields!Name, Len(Text2.Text)) = Text2.Text) And (Left(Data1.Recordset.Fields!id, 1) = Left(Text1.Text, 1)) Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!Name Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!Name
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub Text3_Change()
List4.Visible = False
If (Len(Text3.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 4560
   List4.Top = 3840
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If Left(Data1.Recordset.Fields!family, Len(Text3.Text)) = Text3.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!family Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!family
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub Text4_Change()
List4.Visible = False
If (Len(Text4.Text) > 0) And (Option1.Value = True) Then
   List4.Left = 1080
   List4.Top = 3840
   List4.Visible = True
   List4.Clear
   Data1.Recordset.MoveFirst
   Do
      If Left(Data1.Recordset.Fields!fathername, Len(Text4.Text)) = Text4.Text Then
        
        w = True
        For e = 0 To List4.ListCount - 1
         If List4.List(e) = Data1.Recordset.Fields!fathername Then
           w = False
         End If
        Next e
        
        If w = True Then List4.AddItem Data1.Recordset.Fields!fathername
      End If
      Data1.Recordset.MoveNext
   Loop Until Data1.Recordset.EOF = True
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub


Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text7.SetFocus
End Sub


Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text8.SetFocus
End Sub


Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text9.SetFocus
End Sub


Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text10.SetFocus
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text11.SetFocus
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text12.SetFocus
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text13.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text14.SetFocus
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If List4.Visible = True Then If KeyCode = 40 Then List4.SetFocus
If KeyCode = 13 Then
  KewlButtons4.SetFocus
End If
End Sub
