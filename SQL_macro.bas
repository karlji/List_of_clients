Attribute VB_Name = "SQL_macro"
Option Explicit
Sub SQL_import()
Dim strCountry As String
Dim runSQLonWW As Boolean
Dim db_server As String, db_name As String, db_login As String, db_password As String, table_name As String
Dim h_org_code As String, h_legal_name As String, org As String, SQLStatement As String
Dim LastRowOld As Integer, LastRowNew As Integer, i As Integer
Dim Conn As New ADODB.Connection
Dim mrs As New ADODB.Recordset

strCountry = askAboutCountry()
'SQL information
db_server = ""
db_login = ""
db_password = ""
table_name = " "
h_org_code = ""
h_legal_name = ""

LastRowOld = Worksheets(1).Range(old_col_org & Worksheets(1).Rows.Count).End(xlUp).Row
LastRowNew = Worksheets(1).Range(new_col_org & Worksheets(1).Rows.Count).End(xlUp).Row



If strCountry <> "Brazil" Then
    db_name = "" 'DB for non-Brazil countries
Else
    db_name = "" 'DB for Brazil
End If

'Clearing previous checks
Worksheets(2).Range("A2:H10000").Clear


'Merging orgcodes from old codes table
For i = 12 To LastRowOld
    If i <> LastRowOld Then
         org = org & "'" & Range(old_col_org & i).Value & "',"
    Else
         org = org & "'" & Range(old_col_org & i).Value & "'"
    End If
Next i


'Downloading Orgcodes and legal names + copying to checks from old codes
SQLStatement = "SELECT " & h_org_code & ", " & h_legal_name & _
        " FROM " & table_name & _
        " WHERE " & h_org_code & " in (" & org & ")" & _
        " GROUP BY " & h_org_code & ", " & h_legal_name
        
    
Conn.ConnectionString = "driver={SQL Server};" & _
                        "server=" & db_server & ";" & _
                        "uid=" & db_login & ";" & _
                        "pwd=" & db_password & ";" & _
                        "database=" & db_name

Conn.Open
Conn.CommandTimeout = 600
If Conn.State = adStateOpen Then
    If UCase(Left(SQLStatement, 6)) = "SELECT" Then
        mrs.Open SQLStatement, Conn
        Worksheets(2).Cells(2, 1).CopyFromRecordset mrs
        mrs.Close
    Else
        Conn.Execute SQLStatement
    End If
    runSQLonWW = True
End If
Conn.Close

'setting sql related variables to nothing
Set Conn = Nothing
Set mrs = Nothing
SQLStatement = ""
org = ""

'Merging orgcodes from new codes table
For i = 12 To LastRowNew
    If i <> LastRowNew Then
        org = org & "'" & Range(new_col_org & i).Value & "',"
    Else
        org = org & "'" & Range(new_col_org & i).Value & "'"
    End If
Next i

'Downloading Orgcodes and legal names + copying to checks from new codes
SQLStatement = "SELECT " & h_org_code & ", " & h_legal_name & _
        " FROM " & table_name & _
        " WHERE " & h_org_code & " in (" & org & ")" & _
        " GROUP BY " & h_org_code & ", " & h_legal_name
        
    
Conn.ConnectionString = "driver={SQL Server};" & _
                        "server=" & db_server & ";" & _
                        "uid=" & db_login & ";" & _
                        "pwd=" & db_password & ";" & _
                        "database=" & db_name

Conn.Open
Conn.CommandTimeout = 600
If Conn.State = adStateOpen Then
    If UCase(Left(SQLStatement, 6)) = "SELECT" Then
        mrs.Open SQLStatement, Conn
        Worksheets(2).Cells(2, 4).CopyFromRecordset mrs
        mrs.Close
    Else
        Conn.Execute SQLStatement
    End If
    runSQLonWW = True
End If
Conn.Close

'setting sql related variables to nothing
Set Conn = Nothing
Set mrs = Nothing
SQLStatement = ""
org = ""

'Downloading any possible extra new codes
org = Worksheets(1).Range(new_col_org & LastRowNew).Value
org = Left(org, Len(org) - 3)
SQLStatement = "SELECT " & h_org_code & ", " & h_legal_name & _
        " FROM " & table_name & _
        " WHERE " & h_org_code & " like (" & "'" & org & "%" & "'" & ")" & _
        " GROUP BY " & h_org_code & ", " & h_legal_name
        
    
Conn.ConnectionString = "driver={SQL Server};" & _
                        "server=" & db_server & ";" & _
                        "uid=" & db_login & ";" & _
                        "pwd=" & db_password & ";" & _
                        "database=" & db_name

Conn.Open
Conn.CommandTimeout = 600
If Conn.State = adStateOpen Then
    If UCase(Left(SQLStatement, 6)) = "SELECT" Then
        mrs.Open SQLStatement, Conn
        Worksheets(2).Cells(2, 7).CopyFromRecordset mrs
        mrs.Close
    Else
        Conn.Execute SQLStatement
    End If
    runSQLonWW = True
End If
Conn.Close

'setting sql related variables to nothing
Set Conn = Nothing
Set mrs = Nothing
SQLStatement = ""
org = ""

Exit Sub

errorHandler:
    runSQLonWW = False
    Set Conn = Nothing
    Set mrs = Nothing

End Sub

