VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================== Drowing text using D3DX
'##############################
'#Written by : jaber al_ani   #
'#Date       : 2-Ap-2006      #
'#Time       : 3.00 PM        #
'##############################

Private MainFont           As D3DXFont
Private MainFontDesc       As IFont
Private TFont              As New StdFont
Private text               As String
Private color              As Long

Public Function create(size As Byte, col As Long, tex As String, bold As Byte)

text = tex
color = col

TFont.size = size
TFont.bold = bold

Set MainFontDesc = TFont
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

End Function

Public Function render(left As Integer, top As Integer, W As Integer, H As Integer, col As Long)
D3DX.DrawText MainFont, col, text, Mrect(left, top, W, H), DT_TOP Or DT_LEFT
End Function

Public Function rendertwo(temptext As String, left As Integer, top As Integer, W As Integer, H As Integer, col As Long, size As Byte)
TFont.size = size
TFont.bold = bold

Set MainFontDesc = TFont
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

D3DX.DrawText MainFont, col, temptext, Mrect(left, top, W, H), DT_TOP Or DT_LEFT
End Function


