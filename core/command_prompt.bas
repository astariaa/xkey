Attribute VB_Name = "command_prompt"
Public Sub open_command_prompt()
    line1 = "Enter number 1 thru 9 to run alias."
    line1 = line1 & vbCrLf & "Enter e to edit an alias."
    line1 = line1 & vbCrLf & "Enter v to view an alias definition."
    line1 = line1 & vbCrLf & "Enter sr to unhide register with aliases, hr to hide."
    line1 = line1 & vbCrLf
    
    s = stringify_register()
    prompt_title = "xkey - Command Prompt"
    
    tgt_alias = InputBox(line1 & vbCrLf & s, prompt_title)
    On Error Resume Next
    If tgt_alias = "e" Then
        new_alias = InputBox("Enter number 1 thru 9 to edit." & vbCrLf & s, prompt_title)
        If new_alias <> "" Then
            new_name = InputBox("Enter name for new alias at #" & new_alias & ".", hdr)
            If new_name <> "" Then
                new_sendkeys = InputBox("Type in SendKeys keystroke string." & vbCrLf & s, prompt_title)
                Call write_register(Int(new_alias), new_name, new_sendkeys)
            End If
        End If
    End If
    
    If tgt_alias = "v" Then
        check_alias = InputBox("Enter number 1 thru 9 to view alias definition.", prompt_title)
        register_val = stringify_single_register(check_alias)
        MsgBox (register_val)
    End If
    
    If tgt_alias = "sr" Then
        Windows("__register.xlsx").Visible = True
    End If
    
    If tgt_alias = "hr" Then
        Windows("__register.xlsx").Visible = False
    End If
    
    If (Int(tgt_alias) >= 1) And (Int(tgt_alias) <= 9) Then
        Call run_register_alias(Int(tgt_alias))
    End If
    On Error GoTo 0
End Sub



