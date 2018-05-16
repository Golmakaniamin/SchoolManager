VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "B Nazanin"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   8040
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
      MICON           =   "Form11.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«„Ì‰ ê·„ò«‰Ì"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Õ·Ì· ê— ° »—‰«„Â ‰ÊÌ” Ê ÿ—«Õ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Ì” „ œ«‰‘ ¬„Ê“Ì ¬„Ê“‘ê«Â Â«Ì Œ’Ê’Ì"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê÷Ê⁄ :"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„Â‰œ” —‘ÌœÌ«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«” «œ :"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»—‰«„Â ‰ÊÌ”Ì ÅÌ‘—› Â 1"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "œ—” :"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«— »«ÿ »« „«"
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
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   9330
      Left            =   0
      Picture         =   "Form11.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form11.Hide
Form15.Show
End Sub
