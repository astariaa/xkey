Attribute VB_Name = "execute_command"
Public Sub parse_sendkeys_input(sendkeys_str)
    delimiter = ","
    prepped_sendkeys_str = prepare_sendkeys_input(sendkeys_str)
    sendkeys_list = Split(prepped_sendkeys_str, delimiter)
    On Error Resume Next
    
    For Each s In sendkeys_list
        Application.SendKeys s
    Next
    
    On Error GoTo 0
End Sub

Public Function prepare_sendkeys_input(sendkeys_str)
    s = sendkeys_str
    s = Replace(s, "alt+", "%")
    s = Replace(s, "ctrl+", "^")
    s = Replace(s, "shift+", "+")
    s = Replace(s, "alt", "%")
    s = Replace(s, "ctrl", "^")
    s = Replace(s, "shift", "+")
    prepare_sendkeys_input = s
End Function

Public Sub run_register_alias(alias)
    send_keys_str = read_register(alias)
    Call parse_sendkeys_input(send_keys_str)
End Sub


