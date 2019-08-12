Attribute VB_Name = "Main_macros"
Public old_col_org, new_col_org, old_check_col, new_check_col As String
Public Sub assign_columns()
old_col_org = "D"
new_col_org = "O"
old_check_col = "B"
new_check_col = "M"
End Sub

Sub Check_Click()
Dim LastRowOld As Integer, LastRowNew As Integer, i As Integer
Dim x As Variant, y As Variant
Dim err_num As Byte, missmatch_num As Byte
Dim Today As Date

Call assign_columns
'checking username vs whitelist
Call access_check

'downloading SQL information
Call SQL_import

LastRowOld = Worksheets(1).Range(old_col_org & Worksheets(1).Rows.Count).End(xlUp).Row
LastRowNew = Worksheets(1).Range(new_col_org & Worksheets(1).Rows.Count).End(xlUp).Row

Worksheets(1).Range(old_check_col & "12:" & old_check_col & LastRowOld + 1).Clear
Worksheets(1).Range(new_check_col & "12:" & new_check_col & LastRowNew + 1).Clear

err_num = 0
missmatch_num = 0
Today = Format$(Now, "yyyy-mm-dd")

'Looping through old codes table
On Error GoTo NotFoundOld
For i = 12 To LastRowOld
    y = Worksheets(1).Range(old_col_org & i).Value
    x = Worksheets(2).Columns("A:A").Find(What:=y, SearchOrder:=xlByRows, MatchCase:=False, LookAt:=xlWhole).Row
    
    If Not x = "Not Found in the DB" Then
        x = Worksheets(2).Cells(x, 2).Value
        y = Worksheets(1).Range(old_col_org & i).Offset(, 1).Value
        
        If IsNumeric(x) And IsNumeric(y) Then
            x = CInt(x)
            y = CInt(y)
        End If
        
        If Not y = x Then
            With Worksheets(1).Range(old_check_col & i)
            .Value = x
            .Interior.Color = RGB(255, 0, 0)
            End With
            missmatch_num = missmatch_num + 1
        End If
        
    End If
    
Next i

'Looping through new codes table

On Error GoTo NotFoundNew
For i = 12 To LastRowNew
    y = Worksheets(1).Range(new_col_org & i).Value
    x = Worksheets(2).Columns("D:D").Find(What:=y, SearchOrder:=xlByRows, MatchCase:=False, LookAt:=xlWhole).Row
    
    If Not x = "Not Found in the DB" Then
        x = Worksheets(2).Cells(x, 5).Value
        y = Worksheets(1).Range(new_col_org & i).Offset(, 1).Value
        If Not y = x Then
            With Worksheets(1).Range(new_check_col & i)
            .Value = x
            .Interior.Color = RGB(255, 0, 0)
            End With
            missmatch_num = missmatch_num + 1
        End If
    End If
    
Next i

'Looking for possible new codes
On Error GoTo NotNumeric:
y = Worksheets(1).Range(new_col_org & LastRowNew).Value
y = Right(y, Len(y) - 2)
y = CInt(y)
x = Worksheets(2).Cells(Worksheets(2).Range("G1").End(xlDown).Row, 7).Value
x = Right(x, Len(x) - 2)
x = CInt(x)

If x > y Then
    y = Worksheets(1).Range(new_col_org & LastRowNew).Value
    x = Worksheets(2).Cells(Worksheets(2).Range("G1").End(xlDown).Row, 7).Value
    With Worksheets(1).Range(new_check_col & LastRowNew + 1)
        .Value = "Newer code found: " & x
        .Interior.Color = RGB(255, 0, 0)
    End With
    
    MsgBox ("Warning some new orgcodes are not indicated in this file." & vbNewLine & "Last indicated new code: " & y & _
       vbNewLine & "Largest found code in database: " & x), vbCritical
End If

'Final message
EndMessage:
If err_num > 0 Or missmatch_num > 0 Then
    MsgBox (err_num & " orgcodes not found" & vbNewLine & missmatch_num & " titles don't match"), vbCritical
    With Worksheets(1).Columns(old_check_col).EntireColumn
        .Hidden = False
        .AutoFit
    End With
    With Worksheets(1).Columns(new_check_col).EntireColumn
        .Hidden = False
        .AutoFit
    End With
    
Else
    MsgBox ("Congratulations!" & vbNewLine & "Everything is ok!"), vbInformation
End If

Worksheets(1).Range("Y7").Value = Today
Exit Sub


'Error handler for new orgcodes
NotFoundNew:
x = "Not Found in the DB"
With Worksheets(1).Range(new_check_col & i)
    .Value = x
    .Interior.Color = RGB(255, 0, 0)
End With

err_num = err_num + 1

Resume Next

'Error handler for old orgcodes
NotFoundOld:
x = "Not Found in the DB"
With Worksheets(1).Range(old_check_col & i)
    .Value = x
    .Interior.Color = RGB(255, 0, 0)
End With

err_num = err_num + 1

Resume Next

'Error handler for special last new code
NotNumeric:
MsgBox ("Last new code is special code. Please check if all new codes are in this list manually."), vbCritical
GoTo EndMessage

Resume Next

End Sub
Sub Backup_Click()
Dim strCountry As String
Dim intYear As Integer
Dim strPath As String
Dim myfilename As String
Dim Pos As Long
Dim Today As Date

'checking username vs whitelist
Call access_check

Today = Format$(Now, "yyyy-mm-dd")
strCountry = askAboutCountry()
'select path + date
intYear = Format$(Now, "yyyy")
strPath = "i:\" 'fill in path

'Save as copy with date
Worksheets(1).Range("Y8").Value = Today
ActiveWorkbook.SaveCopyAs Filename:=strPath & Format$(Now, "yyyy-mm-dd") & "_" & ActiveWorkbook.Name

MsgBox ("Back-up created")


End Sub
Function askAboutCountry()
    Dim strCountry As String
    
    strCountry = Left(Right(ThisWorkbook.Name, 7), 2)
    
    Select Case strCountry
        Case "BR"
            strCountry = "Brazil"
        Case "CL"
            strCountry = "Chile"
        Case "PE"
            strCountry = "Peru"
        Case "CO"
            strCountry = "Colombia"
        Case "EC"
            strCountry = "Ecuador"
        Case Else
            MsgBox ("Wrong file name. File name have to contain one of following shortcuts at the end of file name: BR,CL,PE,CO,EC")
            End
    End Select
    
   askAboutCountry = strCountry
    
End Function
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Sub access_check()
Dim whitelist As Variant
Dim username As String

'whitelist of users that can use macro
whitelist = Array("user1", "user2", "user3")

username = VBA.Interaction.Environ$("Username")
If IsInArray(username, whitelist) = False Then
    MsgBox ("You don't have permission to use this macro.")
    End
End If


End Sub




