Attribute VB_Name = "register_io"
Public Function read_register(alias)
    read_register = Workbooks("__register.xlsx").Worksheets("register1").Cells(alias + 1, 3).Value2
End Function

Public Sub write_register(alias, alias_name, keystrokes)
    Set activewb = ActiveWorkbook
    keystrokes = LCase(keystrokes)
    Workbooks("__register.xlsx").Worksheets("register1").Cells(alias + 1, 2).Value = alias_name
    Workbooks("__register.xlsx").Worksheets("register1").Cells(alias + 1, 3).Value = keystrokes
    Workbooks("__register.xlsx").Save
    activewb.Activate
End Sub
Public Function stringify_single_register(alias)
    j = alias + 1
    alias = Workbooks("__register.xlsx").Worksheets("register1").Cells(j, 1).Value
    name = Workbooks("__register.xlsx").Worksheets("register1").Cells(j, 2).Value
    sendkey_str = Workbooks("__register.xlsx").Worksheets("register1").Cells(j, 3).Value
    new_line = Str(alias) & ": "
    If Len(sendkey_str) >= 1 Then
        new_line = new_line & name & " = " & sendkey_str
    End If
         
    stringify_single_register = new_line
End Function

Public Function stringify_register()
    s = "--- Alias Definitions Below ---"
    s = s & vbCrLf
    length_cutoff = 40
    For i = 1 To 9
        new_line = stringify_single_register(i)
        If Len(new_line) >= length_cutoff Then
            new_line = Mid(new_line, 1, length_cutoff)
            new_line = new_line & "..."
        End If
        s = s & new_line & vbCrLf
    Next i
    
    stringify_register = s
End Function

