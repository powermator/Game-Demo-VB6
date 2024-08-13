VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                            ÃåáÇ æÓåáÇ Èßã Ýí ÃáÚÇÈ ÃáãÔÑÞ"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "start.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "start.frx":1CCA
   ScaleHeight     =   7635
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit- - - - - - - - >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting- - - - - ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Play the game->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   6360
      Width           =   6975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   120
      X2              =   3600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6120
      MouseIcon       =   "start.frx":31D0C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÅÚÏÇÏÇÊ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6000
      MouseIcon       =   "start.frx":32016
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÔÛíá ÃááÚÈå"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5880
      MouseIcon       =   "start.frx":32320
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      Index           =   0
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   5160
   End
End
Attribute VB_Name = "Form1"
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
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
DSBuf.PLAY DSBPLAY_DEFAULT
Form1.Hide
Form2.Hide
form3.Show
End If
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
Form2.Show
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.fontsize = 16
Label4.FontBold = True
End Sub
Private Sub Label3_Click()
DSBuf.PLAY DSBPLAY_DEFAULT
Form1.Hide
Form2.Hide
form3.Show
End Sub
Private Sub Label2_Click()
DSBuf.PLAY DSBPLAY_DEFAULT
End
End Sub
