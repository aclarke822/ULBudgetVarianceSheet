Attribute VB_Name = "GitModule"
Public Sub ExportSourceFiles()
    Dim component As VBComponent
    Dim folder As String
    Dim fullPath As String
    Set project = Application.VBE.ActiveVBProject

    folder = Application.ActiveWorkbook.Path + "\"
    Debug.Print folder
    For Each component In project.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            fullPath = folder + component.Name + ToFileExtension(component.Type)
            Debug.Print fullPath
            Debug.Print component.Name
            component.Export (fullPath)
            
        End If
    Next
 
End Sub
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
        ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
        ToFileExtension = vbNullString
    End Select
 
End Function
Public Sub RemoveAllModules()
Dim project As VBProject
Set project = Application.VBE.ActiveVBProject
Dim component As VBComponent

For Each component In project.VBComponents
    If Not component.Name = "DevTools" And (component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule) Then
        project.VBComponents.Remove comp
    End If
Next

End Sub
Public Sub ImportSourceFiles()

Dim folder As String
Dim fullPath As String

folder = Application.ActiveWorkbook.Path + "\"
fullPath = Dir(folder + "*.bas")

'While fullPath <> vbNullString
    For i = 0 To 10
        Debug.Print fullPath
        Debug.Print folder + fullPath
        fullPath = Dir(folder + "*.bas")
    Next i
    
End Sub


