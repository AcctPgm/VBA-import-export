Attribute VB_Name = "mImportExport"
Option Explicit

' Adapted - with much appreciation - from routines developed by Ron Debruin and available at
' https://www.rondebruin.nl/win/s9/win002.htm

Const ptImport = True
Const ptExport = False

Public Sub ExportVBA(control As IRibbonControl)
'
    Dim wbSource As Excel.Workbook
    Dim bExport As Boolean
    Dim lOverwrite As Long
    Dim objFSO As Object
    Dim sExportPath As String
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim Response As VbMsgBoxResult

    Dim sMsg As String
    Dim sMsgForms As String
    Dim sMsgModules As String
    Dim sMsgClasses As String
    
    Set wbSource = ActiveWorkbook
    
    ' Exit if there is no active workbook
    If wbSource Is Nothing Then
        Exit Sub
    End If
    
    ' Exit early if there are no VBA components to save
'    If wbSource.VBProject.VBComponents.Count = 0 Then
    If VBACount(wbSource) = 0 Then
        MsgBox "There are no VBA components in this workbook."
        Exit Sub
    End If
    
    ' Warm & exit if the components can't be saved due to protection
    If wbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    ' Create a file system object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the folder for saving the components
    sExportPath = VBAProjectFolder(ptExport)
    If sExportPath = "*" Then
        Exit Sub
    ElseIf sExportPath = "*Error" Then
        MsgBox "Export Folder doesn't exist:" & vbCr & sExportPath
        Exit Sub
    End If
    
    ' Create lists of components, including whether they already exist in the folder
    lOverwrite = 0
    sMsgForms = ""
    sMsgModules = ""
    sMsgClasses = ""
    
    For Each cmpComponent In wbSource.VBProject.VBComponents
        sFileName = cmpComponent.Name

        ' Add the prospective exported file name to the respective list
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = sFileName & ".cls"
                If objFSO.FileExists(objFSO.BuildPath(sExportPath, sFileName)) Then
                    sMsgClasses = sMsgClasses & "    " & "* " & sFileName & vbCr
                    lOverwrite = lOverwrite + 1
                Else
                    sMsgClasses = sMsgClasses & "    " & sFileName & vbCr
                End If
            Case vbext_ct_MSForm
                sFileName = sFileName & ".frm"
                If objFSO.FileExists(objFSO.BuildPath(sExportPath, sFileName)) Then
                    sMsgForms = sMsgForms & "    " & "* " & sFileName & vbCr
                    lOverwrite = lOverwrite + 1
                Else
                    sMsgForms = sMsgForms & "    " & sFileName & vbCr
                End If
            Case vbext_ct_StdModule
                sFileName = sFileName & ".bas"
                If objFSO.FileExists(objFSO.BuildPath(sExportPath, sFileName)) Then
                    sMsgModules = sMsgModules & "    " & "* " & sFileName & vbCr
                    lOverwrite = lOverwrite + 1
                Else
                    sMsgModules = sMsgModules & "    " & sFileName & vbCr
                End If
            Case vbext_ct_Document
                ' This is a worksheet or workbook object.
                ' It won't be exported, so nothing to check
        End Select
    Next cmpComponent
    
    ' Inform user of items that will be exported, with warning about overwrites
    sMsg = "Components that will export:" & vbCr & vbCr
    If Len(sMsgForms) > 0 Then
        sMsg = sMsg & "Forms:" & vbCr & sMsgForms & vbCr
    End If
    If Len(sMsgModules) > 0 Then
        sMsg = sMsg & "Modules:" & vbCr & sMsgModules & vbCr
    End If
    If Len(sMsgClasses) > 0 Then
        sMsg = sMsg & "Classes:" & vbCr & sMsgClasses & vbCr
    End If
    
    If lOverwrite > 0 Then
        sMsg = sMsg & "* " & lOverwrite & " item" & IIf(lOverwrite > 1, "s", "") & " will be overwritten" & _
            vbCr & vbCr & "Proceed?"
    End If
    
    Response = MsgBox(sMsg, vbYesNoCancel + IIf(lOverwrite > 0, vbDefaultButton2, vbDefaultButton1), "Confirm VBA Export")
    If Response <> vbYes Then
        Exit Sub
    End If
    
    ' Save each component
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

Public Sub ImportVBA(control As IRibbonControl)
'
    Dim wbTarget As Workbook
    Dim objFSO As Object
    Dim objFile As Object
    Dim lImportCount As Long
    Dim sImportPath As String
    Dim sFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    Dim Response As VbMsgBoxResult

    Dim sMsg As String
    Dim sMsgForms As String
    Dim sMsgModules As String
    Dim sMsgClasses As String

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
    sImportPath = VBAProjectFolder(ptImport)
    If sImportPath = "*" Then
        Exit Sub
    ElseIf sImportPath = "*Error" Then
        MsgBox "Import Folder doesn't exist:" & vbCr & sImportPath
        Exit Sub
    End If

    ' Show list of components that will be imported
    lImportCount = 0
    sMsgForms = ""
    sMsgModules = ""
    sMsgClasses = ""
    
    For Each objFile In objFSO.GetFolder(sImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Then
            sMsgClasses = sMsgClasses & "  " & objFile.Name & vbCr
            lImportCount = lImportCount + 1
        End If
        If (objFSO.GetExtensionName(objFile.Name) = "frm") Then
            sMsgForms = sMsgForms & "  " & objFile.Name & vbCr
            lImportCount = lImportCount + 1
        End If
        If (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            sMsgModules = sMsgModules & "  " & objFile.Name & vbCr
            lImportCount = lImportCount + 1
        End If
    Next objFile
    
        ' Exit if there are no VBA components in the selected folder
    If lImportCount = 0 Then
        MsgBox "There are no files to import from" & vbCr & vbTab & sImportPath
        Exit Sub
    End If
    
    sMsg = "Components that will import:" & vbCr & vbCr
    If Len(sMsgForms) > 0 Then
        sMsg = sMsg & "Forms:" & vbCr & sMsgForms & vbCr
    End If
    If Len(sMsgModules) > 0 Then
        sMsg = sMsg & "Modules:" & vbCr & sMsgModules & vbCr
    End If
    If Len(sMsgClasses) > 0 Then
        sMsg = sMsg & "Classes:" & vbCr & sMsgClasses & vbCr
    End If

    Response = MsgBox(sMsg & _
        IIf(VBACount(wbTarget) > 0, "CAUTION: Existing VBA components will be deleted!" & vbCr, "") & _
        vbCr & "Proceed?", _
        vbYesNoCancel + vbDefaultButton2, "Importing VBA Components")
    If Response <> vbYes Then
        Exit Sub
    End If

    ' Delete all modules/Userforms from the ActiveWorkbook before importing
    Call DeleteVBA(wbTarget)

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

Function VBACount(wb As Workbook)
'
    Dim cmpComponent As VBIDE.VBComponent
    
    VBACount = 0
    For Each cmpComponent In wb.VBProject.VBComponents
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                VBACount = VBACount + 1
            Case vbext_ct_MSForm
                VBACount = VBACount + 1
            Case vbext_ct_StdModule
                VBACount = VBACount + 1
            Case vbext_ct_Document
                ' This is a worksheet or workbook object, and won't be exported
        End Select
    Next cmpComponent
End Function

Function VBAProjectFolder(PathType As Boolean) As String
' PathType: ptImport - location for importing VBA components
'           ptExport - location for exporting VBA components
' Returns:  path, if a path was successfully selected
'           * if user cancelled file dialog
'           *Error if the selected path somehow doesn't exist
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
        sFilePath = Application.ActiveWorkbook.Path
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
        
        If .Show Then
            sFilePath = .SelectedItems.Item(1)
        Else
            sFilePath = "*"
        End If
    End With
    
    If sFilePath = "*" Then
        VBAProjectFolder = "*"
    ElseIf objFSO.FolderExists(sFilePath) = True Then
        VBAProjectFolder = sFilePath
    Else
        VBAProjectFolder = "*Error"
    End If
End Function

Function DeleteVBA(wb As Workbook)
'
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


