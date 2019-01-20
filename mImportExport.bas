Attribute VB_Name = "mImportExport"
Option Explicit

' Adapted - with much appreciation - from routines developed by Ron Debruin and available at
' https://www.rondebruin.nl/win/s9/win002.htm

Const ptImport = True
Const ptExport = False

Public Sub ExportModules()
    Dim wbSource As Excel.Workbook
    Dim bExport As Boolean
    Dim objFSO As Object
    Dim sExportPath As String
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    Set wbSource = ActiveWorkbook
    
    ' Exit if there is no active workbook
    If wbSource Is Nothing Then
        Exit Sub
    End If
    
    ' Exit early if there are no VBA components to save
    If wbSource.VBProject.VBComponents.Count = 0 Then
        MsgBox "There are no VBA components in this workbook."
        Exit Sub
    End If
    
    ' Warm & exit if the components can't be saved due to protection
    If wbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    sExportPath = FolderWithVBAProjectFiles(ptExport)
    If sExportPath = "Error" Then
        MsgBox "Export Folder doesn't exist:" & vbCr & sExportPath
        Exit Sub
    End If
    
''' opion: offer to delete all existing code modules from the export folder
'    On Error Resume Next
'        Kill FolderWithVBAProjectFiles & "\*.*"
'    On Error GoTo 0
'
    For Each cmpComponent In wbSource.VBProject.VBComponents
        bExport = True
        sFileName = cmpComponent.Name

        ' Add the correct file extension for module being exported
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = sFileName & ".cls"
            Case vbext_ct_MSForm
                sFileName = sFileName & ".frm"
            Case vbext_ct_StdModule
                sFileName = sFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ' Export the component to a text file
            cmpComponent.Export objFSO.BuildPath(sExportPath, sFileName)
            
            ''' option: remove each module from the project if you want
            ' wbSource.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent

    MsgBox "Export is completed"
End Sub

Public Sub ImportModules()
    Dim wbTarget As Workbook
    Dim objFSO As Object
    Dim objFile As Object
    Dim lImportCount As Long
    Dim sImportPath As String
    Dim sFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    Set wbTarget = ActiveWorkbook
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Exit if there is no active workbook
    If wbTarget Is Nothing Then
        Exit Sub
    End If
    
    ' Warn & exit if modules can't be imported due to protection
    If wbTarget.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to Import the code"
        Exit Sub
    End If

    ' Prevent loading routines into this workbook, to avoid unintended consequences
    If wbTarget.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
            "Not possible to import in this - the macro - workbook "
        Exit Sub
    End If

    ' User-selected path where the code modules are located.
    sImportPath = FolderWithVBAProjectFiles(ptImport)
    If sImportPath = "Error" Then
        MsgBox "Import Folder doesn't exist"
        Exit Sub
    End If

    ' Inform the user if there aren't any VBA components to import from the folder
    For Each objFile In objFSO.GetFolder(sImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            lImportCount = lImportCount + 1
        End If
    Next objFile
    
    If lImportCount = 0 Then
        MsgBox "There are no files to import from" & vbCr & _
            vbTab & sImportPath
        Exit Sub
    End If

    ' Delete all modules/Userforms from the ActiveWorkbook before importing
    Call DeleteVBAModulesAndUserForms(wbTarget)

    ' Import all the code modules - .bas, .frm, .cls - in the specified path
    Set cmpComponents = wbTarget.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(sImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
    Next objFile
    
    MsgBox "Import is complete"
End Sub

Function FolderWithVBAProjectFiles(PathType As Boolean) As String
' PathType: ptImport - location for importing VBA components
'           ptExport - location for exporting VBA components
'
    Dim WshShell As Object
    Dim fd As FileDialog
    Dim objFSO As Object
    Dim sFilePath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' If the file hasn't yet been saved, offer to save the VBA components in "My Documents"
    ' Otherwise, offer to save the components in the workbook's folder
    If Application.ActiveWorkbook.Path = "" Then
        sFilePath = WshShell.SpecialFolders("MyDocuments")
    Else
        sFilePath = Application.ActiveWorkbook.FullName
    End If
    If Right(sFilePath, 1) <> "\" Then
        sFilePath = sFilePath & "\"
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If PathType = ptExport Then
            .Title = "Folder for Saving VBA Components"
        Else
            .Title = "Folder to Import VBA Components from"
        End If
        
        .InitialFileName = sFilePath
        If .Show = -1 Then
            sFilePath = .SelectedItems.Item(1)
        End If
    End With
    
    If objFSO.FolderExists(sFilePath) = True Then
        FolderWithVBAProjectFiles = sFilePath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
End Function

Function DeleteVBAModulesAndUserForms(wb As Workbook)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = wb.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            ' Thisworkbook or worksheet module -> do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function


