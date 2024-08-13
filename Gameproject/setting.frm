VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                                                                          ÇáÇÚÏÇÏÇÊ"
   ClientHeight    =   7590
   ClientLeft      =   105
   ClientTop       =   885
   ClientWidth     =   8790
   Icon            =   "setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÏÞå ÇáÇßÓÇÁ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÚãÞ Çááæäí"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãæÇÝÞ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6960
      MouseIcon       =   "setting.frx":1CCA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÅáÛÇÁ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4080
      MouseIcon       =   "setting.frx":1FD4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ÃáÇÚÏÇÏÇÊ ÃáÇÝÊÑÇÖíå"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      MouseIcon       =   "setting.frx":22DE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÏÞå ÇáÚÑÖ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
'=================================
'##############################
'#Written by : jaber al_ani   #
'#                            #
'#     jaberani@yahoo.com     #
'##############################
'==================================
'==================================

'ÇáÇÚáÇä Úä ÇáßÇÆä ÇáÑÆíÓí
Dim DX As DirectX8
'ÇáÇÚáÇä Úä ßÇÆä ÇáÕæÊ
Dim DS As DirectSound8
'ÇáÇÚáÇä Úä ÈÝÑ ÇáÕæÊ
Dim DSBuf As DirectSoundSecondaryBuffer8
'ÇáÇÚáÇä Úä ßÇÆä ÍÇáå ÇáÕæÊ
Dim DSBDesc As DSBUFFERDESC
Private Sub Form_Load()

Label2.fontsize = 14
Label3.fontsize = 14
Label4.fontsize = 14

Me.Show

'ÇäÔÇÁ ÇáßÇÆä ÇáÑÆíÓí
Set DX = New DirectX8

'ÇäÔÇÁ ßÇÆä ÇáÕæÊ ÇáÑÆíÓí
Set DS = DX.DirectSoundCreate("")

'ÖÈØ ÎÕÇÆÕ ÇáßÇÆä ÇáÑÆíÓí ÈÇä áå ÇáÇÓÈÞíå Ýí ÇáÊäÝíÐ
DS.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY

'ÌÚá ßÇÆä ÍÇáå ÇáÕæÊ ãÓÄæá Úä ÇáÊÍßã ÈÇÑÊÝÇÚ ÇáÕæÊ ÝÞØ
DSBDesc.lFlags = DSBCAPS_CTRLVOLUME

'ÊÍãíá ÇáÕæÊ ÏÇÎá ÇáÈÝÑ
Set DSBuf = DS.CreateSoundBufferFromFile(App.Path & "\Data\Sounds\click.wav", DSBDesc)

Combo1.AddItem "1024 X 768", 0
Combo1.AddItem "800  X 600", 1

Combo2.AddItem "16", 0
Combo2.AddItem "24", 1
Combo2.AddItem "32", 2

Combo3.AddItem "ãÊæÓØå", 0
Combo3.AddItem "ÚÇáíÜå", 1
Combo3.AddItem "ãäÎÝÙå", 2

Open App.Path & "\System\Setting\Setting.dat" For Random As #1
Get #1, 1, temptext
Combo1.text = temptext
Get #1, 2, temptext
Combo2.text = temptext
Get #1, 3, temptext
Combo3.text = temptext

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HC0C0C0
Label3.ForeColor = &HC0C0C0
Label4.ForeColor = &HC0C0C0
Label2.fontsize = 14
Label3.fontsize = 14
Label4.fontsize = 14
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbRed
Label2.fontsize = 16
Label2.FontBold = True
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
Label3.fontsize = 16
Label3.FontBold = True
End Sub
Private Sub Label4_Click()
DSBuf.PLAY DSBPLAY_DEFAULT

Put #1, 1, Combo1.text
Put #1, 2, Combo2.text
Put #1, 3, Combo3.text

Form2.Hide
Form1.Show
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.fontsize = 16
Label4.FontBold = True

End Sub
Private Sub Label3_Click()
DSBuf.PLAY DSBPLAY_DEFAULT
Form2.Hide
Form1.Show
End Sub
Private Sub Label2_Click()
DSBuf.PLAY DSBPLAY_DEFAULT
Combo1.text = "1024 X 768"
Combo2.text = "16"
Combo3.text = "ãÊæÓØå"
End Sub

