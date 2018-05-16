VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Titr"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3135
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      _cx             =   4895
      _cy             =   5530
      FlashVars       =   ""
      Movie           =   "D:\Programing\Amin\PISHTAZ SOFTWARE\1.swf"
      Src             =   "D:\Programing\Amin\PISHTAZ SOFTWARE\1.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoBorder"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   420
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   4
      PasswordChar    =   "["
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÎÑæÌ"
      Height          =   375
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "æÑæÏ Èå ÓíÓÊã"
      Height          =   375
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer, w, e, r, t, y As String
Private Sub Command1_Click()
If q > 1 Then End
If Text1.Text = Text2.Text Then
   Form1.Hide
   Form15.Show
Else
   q = q + 1
   Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
'If basItemExist.ItemExist("C:\WINDOWS\system32\cws32sh.dll") = True Then
'   If basItemExist.ItemExist("C:\Program Files\Microsoft Visual Studio\cpmsh.dll") = True Then
'      If basItemExist.ItemExist("C:\Program Files\Internet Explorer\cpish.dll") = True Then
'         If basItemExist.ItemExist("C:\DATA BASE INFORMATION\KANON SOFTWARE\cdkspr.dll") = True Then
''            If basItemExist.ItemExist("D:\AMIN\KANON SOFTWARE\daks.dll") = True Then
               ShockwaveFlash1.Movie = App.Path + "\1.swf"
               filenames$ = App.Path & "\pasur.A@G"
               Open filenames$ For Input As #1
               Do While Not EOF(1)
                  Input #1, w
                  Text2.Text = w
               Loop
               Close #1
 '           Else
 '              Command1.Visible = False
 '              e = MsgBox("áØÝÇ ÈÇ ÈÑäÇãå äæíÓ ÊãÇÓ ÈíÑíÏ" & Chr(13) & Chr(10) & "55708769", vbCritical + vbMsgBoxRight, "")
 '              End
 '           End If
 '        Else
 '           Command1.Visible = False
 '           e = MsgBox("áØÝÇ ÈÇ ÈÑäÇãå äæíÓ ÊãÇÓ ÈíÑíÏ" & Chr(13) & Chr(10) & "55708769", vbCritical + vbMsgBoxRight, "")
 '           End
 '        End If
 '     Else
 '        Command1.Visible = False
 '        e = MsgBox("áØÝÇ ÈÇ ÈÑäÇãå äæíÓ ÊãÇÓ ÈíÑíÏ" & Chr(13) & Chr(10) & "55708769", vbCritical + vbMsgBoxRight, "")
 '        End
 '     End If
 '  Else
 '     Command1.Visible = False
 '     e = MsgBox("áØÝÇ ÈÇ ÈÑäÇãå äæíÓ ÊãÇÓ ÈíÑíÏ" & Chr(13) & Chr(10) & "55708769", vbCritical + vbMsgBoxRight, "")
 '     End
 '  End If
'Else
'   Command1.Visible = False
'   e = MsgBox("áØÝÇ ÈÇ ÈÑäÇãå äæíÓ ÊãÇÓ ÈíÑíÏ" & Chr(13) & Chr(10) & "55708769", vbCritical + vbMsgBoxRight, "")
'   End
'End If
Text1.Text = ""
q = 0
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 3 Then Call Command1_Click
End Sub
