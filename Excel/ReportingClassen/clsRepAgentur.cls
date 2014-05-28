VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepAgentur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRepAgentur
' Author    : Dmitry Gorelenkov
' Date      : 09.11.2012
' Purpose   : Agentur datenhaltung
' Requires  :
' Hint      : LOW QUALITY CODE!!!
'---------------------------------------------------------------------------------------

Option Explicit
Private m_sName As String
Private m_sCode As String
Private m_lFarbe As Long
Public names As Variant
Public IDs As Variant
Public realNames As Variant

Private Sub Class_Initialize()
    m_sName = ""
    m_sCode = ""
    m_lFarbe = rgbWhite
End Sub

Public Property Get name() As String
    name = m_sName
End Property

Public Property Let name(ByVal sName As String)
    m_sName = sName
End Property

Public Property Get code() As String
    code = m_sCode
End Property

Public Property Let code(ByVal scode As String)
    m_sCode = scode
End Property

Public Property Get farbe() As Long
    farbe = m_lFarbe
End Property

Public Property Let farbe(ByVal lfarbe As Long)
    m_lFarbe = lfarbe
End Property



