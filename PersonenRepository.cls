VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonenRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const m_connection_string = "Provider='SQLOLEDB';Data Source='Kepler';Initial Catalog='ruv';Integrated Security='SSPI';"
Private m_connection As ADODB.connection



Private Sub Class_Initialize()
    Set m_connection = New ADODB.connection
    m_connection.Open m_connection_string
    
    If m_connection.State = 1 Then
        Debug.Print "Connected!"
    Else
        Debug.Print "Not Connected!"
    End If
End Sub


Private Sub Class_Terminate()
    m_connection.Close
End Sub

Public Sub fillSheetWithData(ByVal sheet As Integer, ByVal range As String)

    Dim recordset As New ADODB.recordset
    
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen", m_connection, adOpenStatic
        
        
    
    End With
    
    ThisWorkbook.Sheets(sheet).range(range).CopyFromRecordset recordset
    
    
    recordset.Close
    
    
End Sub

''-------------------------------------------------------------------
'' Function: Save (2009-06-28)
''-------------------------------------------------------------------
''/**
'' <summary>
'' Person persistieren
'' </summary>
'' <param name="person">zu speichernede Person</param>
'' <returns>Person</returns>
'' <remarks>
'' Eine tolle Bemerkung
'' </remarks>
''**/
''------------------------------------------------------------------
Public Function Save(ByRef person As person) As person
    Dim recordset As New ADODB.recordset
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen", m_connection, adOpenKeyset, adLockOptimistic
        
        .AddNew ' neue Zeile
        .Fields("vorname") = person.Vorname
        .Fields("nachname") = person.Nachname
        .Update ' Schreibt in die Datenbank zurück
        person.id = .Fields("id") ' Hole ID aus dem aktuellen Datensatz
        
        
        .Close
    End With
    
    Set Save = person
End Function


Public Function Update(person As person) As Boolean
    Update = False
    Dim recordset As New ADODB.recordset
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen where id=" & person.id, m_connection, adOpenKeyset, adLockOptimistic
        
        If Not .EOF Then
        
            .Fields("vorname") = person.Vorname
            .Fields("nachname") = person.Nachname
            .Update ' Schreibt in die Datenbank zurück
            Update = True
        End If
        .Close
    End With
    
    
End Function

Public Function DeleteById(id As Integer) As Boolean
    DeleteById = False
    Dim recordset As New ADODB.recordset
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen where id=" & id, m_connection, adOpenKeyset, adLockOptimistic
        
        If Not .EOF Then
            .Delete
            .Update ' Schreibt in die Datenbank zurück
            DeleteById = True
        End If
        .Close
    End With

End Function

Public Function Delete(person As person) As Boolean
   Delete = DeleteById(person.id)
End Function


Public Function FindById(id As Integer) As person
    
    Dim recordset As New ADODB.recordset
    Set FindById = Nothing
    
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen where id=" & id, m_connection, adOpenStatic
        
        If Not .EOF Then
            
            Set FindById = CreatePersonFromRecordset(recordset)
            
         End If
        
        .Close
        
    End With
End Function

Public Function FindAll() As Collection
    Dim result  As New Collection
   
    
    Dim recordset As New ADODB.recordset
    With recordset
        .CursorLocation = adUseClient
        .Open "Select * from tbl_personen", m_connection, adOpenStatic
        
        
        Do Until .EOF
            result.Add CreatePersonFromRecordset(recordset)
            
            .MoveNext
            
        Loop

        .Close
        
    End With
    
     Set FindAll = result
End Function

Private Function CreatePersonFromRecordset(ByRef recordset As ADODB.recordset) As person
    Dim result As New person
    result.id = recordset.Fields("id")
    result.Vorname = recordset.Fields("vorname")
    result.Nachname = recordset.Fields("nachname")
    
    Set CreatePersonFromRecordset = result
End Function


