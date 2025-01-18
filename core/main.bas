Attribute VB_Name = "main"
Public Sub init()
    Application.ScreenUpdating = False
    On Error Resume Next
    Workbooks.Open (ThisWorkbook.Path & "\__register.xlsx")
    Windows("__register.xlsx").Visible = False
    On Error GoTo 0
    Application.OnKey "{F3}", "open_command_prompt"
    Application.ScreenUpdating = True
    Workbooks.Add
End Sub

Public Sub export_modules()
    parent_path = ThisWorkbook.Path + "\"
    ThisWorkbook.VBProject.VBComponents("ThisWorkbook").Export parent_path + "ThisWorkbook.bas"
    ThisWorkbook.VBProject.VBComponents("command_prompt").Export parent_path + "command_prompt.bas"
    ThisWorkbook.VBProject.VBComponents("execute_command").Export parent_path + "execute_command.bas"
    ThisWorkbook.VBProject.VBComponents("main").Export parent_path + "main.bas"
    ThisWorkbook.VBProject.VBComponents("register_io").Export parent_path + "register_io.bas"

End Sub
