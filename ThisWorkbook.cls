Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Target.Column = 7 And Target.Count = 1 Then
        Application.EnableEvents = False
        GreenStarMacros.HandleChange Sh, Target
        Application.EnableEvents = True
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    GreenStarMacros.InitWorkbook
    On Error GoTo 0
End Sub
