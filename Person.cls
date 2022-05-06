VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id As Integer
Private m_vorname As String
Private m_nachname As String



Public Property Get id() As Integer
    id = m_id
End Property

Public Property Let id(ByVal newId As Integer)
    m_id = newId
End Property

Public Property Get Vorname() As String
    Vorname = m_vorname
End Property

Public Property Let Vorname(ByVal newVorname As String)
    m_vorname = newVorname
End Property

Public Property Get Nachname() As String
    Nachname = m_nachname
End Property

Public Property Let Nachname(ByVal newNachname As String)
    m_nachname = newNachname
End Property


Private Sub Class_Initialize()
    id = 0
    Vorname = ""
    Nachname = ""
End Sub

Public Function ToString()

    ToString = "Person[ Id=" & id & ", Vorname=" & Vorname & ", Nachname=" & Nachname & "]"

End Function

