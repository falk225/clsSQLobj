Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sel As String
Public From As String
Public Where As String
Public OrderBy As String
Public Insert As String
Public str_SQL As String
Public dict_rs As New Dictionary 'collection of recordsets created using addRecordset
Private db As DAO.Database
Private rs_object As DAO.Recordset

'Single recordset example:
'dim oSQL as new clsSQLobj
'oSQL.strSQL = "my sql string here"
'debug.print osql.rs.recordcount

'Multiple recordset example:
'dim oSQL as new clsSQLobj
'oSQL.strSQL = "my sql string here"
'oSQL.addRecordset "rs1"
'oSQL.strSQL = "my 2nd sql string here"
'oSQL.addRecordset "rs2"
'debug.print oSQL!rs1.recordcount & " " & oSQL!rs2.recordcount

Public Property Let strSQL(strSQL As String)
    str_SQL = strSQL
    Set rs_object = Nothing
End Property

Public Property Get strSQL() As String
    strSQL = str_SQL
End Property

Private Sub Class_Initialize()
    Set db = CurrentDb
End Sub

Public Function makeStr() As String
'Returns str_SQL or if empty creates SQL str using Sel, From, Where and OrderBy properties
Dim strOrder As String
    strOrder = ""
    
    If str_SQL = "" Then
        If Sel = "" Then
            MsgBox "Select property not set."
            Exit Function
        End If
        
        If From = "" Then
            MsgBox "From property not set."
            Exit Function
        End If
        
        If Where = "" Then
            MsgBox "Where property not set."
            Exit Function
        End If
        
        If OrderBy <> "" Then 'order by is not required, but if provided will sort results
            strOrder = " ORDER BY " & OrderBy
        End If
        
        makeStr = "SELECT " & Sel & " FROM " & From & " WHERE (" & Where & ")" & strOrder & ";"
    Else
        makeStr = str_SQL
    End If
End Function

Public Function createRecordset(Optional strSQL As String = "Unused") As DAO.Recordset
'This returns a recordset using the SQL string created by the makeStr function unless a SQL string is specified
    If strSQL = "Unused" Then strSQL = makeStr()
    Set rs_object = db.OpenRecordset(strSQL)
    Set createRecordset = rs_object
End Function

Public Sub appendToTable()
'If Insert, From, and Where have been set and strSQL is empty this command will append the
'records specified by From and Where to the table specified in Insert
    If strSQL <> "" Then
        MsgBox "strSQL not compatable with append."
        Exit Sub
    End If
       
    If From = "" Then
        MsgBox "From property not set."
        Exit Sub
    End If
    
    If Where = "" Then
        MsgBox "Where property not set."
        Exit Sub
    End If
    
    If Insert = "" Then
        MsgBox "Insert Property not set"
        Exit Sub
    End If
Dim qryAppend As QueryDef
Dim tempstr As String
    tempstr = "INSERT INTO " & Insert & " SELECT " & From & ".* FROM " & From & " WHERE (" & Where & ");"
    Set qryAppend = db.CreateQueryDef("", tempstr)
    qryAppend.Execute
    qryAppend.Close
End Sub

Public Sub deleteRows()
'If From and Where are set and strSQL is empty then the rows specified by Where will be removed
'from the table specified in From
    If strSQL <> "" Then
        MsgBox "strSQL not compatable with append."
        Exit Sub
    End If
    If From = "" Then
        MsgBox "From property not set."
        Exit Sub
    End If
    
    If Where = "" Then
        MsgBox "Where property not set."
        Exit Sub
    End If

Dim qryDelete As QueryDef
Dim tempstr As String
    tempstr = "Delete " & From & ".* " & " FROM " & From & " WHERE (" & Where & ");"
    Set qryDelete = db.CreateQueryDef("", tempstr)
    qryDelete.Execute
    qryDelete.Close
End Sub

Public Sub runUpdateQry()
'Creates QueryDef using strSQL property and executes it
If strSQL = "" Then
    MsgBox "You must assign a sql statement to the property strSQL."
    Exit Sub
End If
Dim qryUpdate As QueryDef
Set qryUpdate = db.CreateQueryDef("", strSQL)
qryUpdate.Execute
qryUpdate.Close
End Sub

Public Sub addRecordset(key As String, Optional strSQL As String = "Unused")
'Adds a recordset of key:Key to the rs collection using the function createRecordset
If strSQL = "Unused" Then strSQL = makeStr()
If key = "" Then
    MsgBox "No key was supplied, could not create recordset"
    Exit Sub
End If
dict_rs.Add key, createRecordset(strSQL)
End Sub

Public Sub rsClose(rsKey As String)
Dim i As Integer
    dict_rs(rsKey).Close
    dict_rs.Remove (rsKey)
End Sub

Public Sub closeAll()
'closes all open recordsets in dict_rs
Dim i As Integer
    For i = 0 To ArrayLen(dict_rs.Keys()) - 1
        On Error Resume Next
        dict_rs(dict_rs.Keys(i)).Close
        dict_rs.Remove (dict_rs.Keys(i))
    Next
End Sub

Public Function Item(ByVal index As Variant) As DAO.Recordset
Attribute Item.VB_UserMemId = 0
    Set Item = dict_rs.Item(index)
End Function

Public Function rs() As DAO.Recordset
'allows you to use clsSQLOBJ.rs as a rs object before you even run the query
    If rs_object Is Nothing Then
        Set rs_object = createRecordset
    End If
    Set rs = rs_object
End Function

Public Sub refresh()
    If rs_object Is Nothing Then
        Set rs_object = createRecordset
    Else
        rs_object.Close
        Set rs_object = Nothing
        Set rs_object = createRecordset
    End If
End Sub

Public Function has_rs() As Boolean
    has_rs = Not rs_object Is Nothing
End Function