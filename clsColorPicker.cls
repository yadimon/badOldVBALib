VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ******** Code Start ********
'This code was originally written by Terry Kreft,
'and modified by Stephen Lebans
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
' Contact Stephen@lebans.com
'

Private Type COLORSTRUC
  lStructSize As Long
  hWnd As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Const CC_SOLIDCOLOR = &H80

Private Declare Function ChooseColor _
    Lib "comdlg32.dll" Alias "ChooseColorA" _
    (pChoosecolor As COLORSTRUC) As Long


Public Function showColorPick()
  Dim X As Long, CS As COLORSTRUC, CustColor(16) As Long
  CS.lStructSize = Len(CS)
  CS.hWnd = hWndAccessApp
  CS.Flags = CC_SOLIDCOLOR
  CS.lpCustColors = String$(16 * 4, 0)
  X = ChooseColor(CS)
  If X = 0 Then
    ' ERROR - use Default White
    showColorPick = RGB(255, 255, 255) ' White
    Exit Function
  Else
    ' Normal processing
     showColorPick = CS.rgbResult
  End If
End Function
' ********* Code End ***********
