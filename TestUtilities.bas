Attribute VB_Name = "TestUtilities"
Public Sub AcadExportAllModules()
    Dim appAcad As AcadApplication
    Set appAcad = GetAcad()
    AcadExportVbaProjects appAcad
    
End Sub

Sub AcadExportVbaProjects(appAcad As AcadApplication)
    Dim objIDE As Object
    Dim projects As Object
    Dim exportFolder As String
    Dim project As Object
    
    Set objIDE = AcadApplication.VBE
    Set projects = objIDE.VBProjects
    
    For Each project In projects
        exportFolder = GetFolderName(project.fileName)
        For Each Module In project.VBComponents
            Module.Export (exportFolder & "\" & Module.Name & ".bas")
        Next Module
    Next project

End Sub




' Section AutoCAD










Public Function GetAcad() As AcadApplication
    Dim appAcad As AcadApplication
    Set appAcad = GetObject(, "AutoCAD.Application")
    Set GetAcad = appAcad
End Function

Public Sub InsertDrawingThenDelete(doc As AcadDocument, drawingToInsert As String)
    Dim XYZScale As Double
    Dim rotation As Double
    XYZScale = 1#
    rotation = 0#
    InsertDrawingAsBlockThenDelete doc, ConstantInsertionPointZeroZeroZero(), drawingToInsert, XYZScale, rotation
End Sub

Public Function SaveCurrentDwgToTempFolder(doc As AcadDocument) As String
    Dim fileName As String
    fileName = GetNewDwgFileInTempFolder()
    doc.SaveAs (fileName)
    SaveCurrentDwgToTempFolder = fileName
End Function

Public Sub InsertDrawingAsBlockThenDelete(doc As AcadDocument, insertionPt As Variant, drawingToInsert As String, XYZScale As Double, rotation As Double)
    Dim objBlockRef As AcadBlockReference
    On Error Resume Next
        Set objBlockRef = doc.ModelSpace.InsertBlock(insertionPt, drawingToInsert, XYZScale, _
        XYZScale, XYZScale, rotation)
    If Err Then
        MsgBox "Unable to insert this block. " + Err.Description
    End If
    objBlockRef.Delete
End Sub








' Section Constants










Public Function ConstantInsertionPointZeroZeroZero() As Variant
    Dim insertionPointZeroZeroZero(0 To 2) As Double
    insertionPointZeroZeroZero(0) = 0
    insertionPointZeroZeroZero(1) = 0
    insertionPointZeroZeroZero(2) = 0
    ConstantInsertionPointZeroZeroZero = insertionPointZeroZeroZero
End Function

Public Function ConstantInsertionPointOneOneOne() As Variant
    Dim insertionPointOneOneOne(0 To 2) As Double
    insertionPointOneOneOne(0) = 1
    insertionPointOneOneOne(1) = 1
    insertionPointOneOneOne(2) = 1
    ConstantInsertionPointOneOneOne = insertionPointOneOneOne
End Function









' Section Files








'FSO, Needs reference to Microsoft Visual Basic for Applications Extensibility 5.3 or greater
Sub FSOGetFileName(fileName As String)
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'Get File Name no Extension
    FSOGetFileName = Left(fileName, InStr(fileName, ".") - 1)
End Sub

Public Function GetFolderName(path As String) As String
    Dim directory As String
    directory = Left(path, InStrRev(path, "\") - 1)
    GetFolderName = directory
End Function

Public Function GetNewDwgFileInTempFolder() As String
    Dim fileName As String
    fileName = GetNewFileInTempFolderWithExtension(".dwg")
    GetNewDwgFileInTempFolder = fileName
End Function

Public Function GetNewFileInTempFolderWithExtension(newExtension As String) As String
    Dim fileName As String
    fileName = GetNewFileInTempFolder()
    fileName = Replace(fileName, ".tmp", newExtension)
    GetNewFileInTempFolderWithExtension = fileName
End Function

Public Function GetNewFileInTempFolder() As String
    Dim fileName As String
    fileName = GetTempFileName()
    Dim tempFolder As String
    tempFolder = GetTempFolder()
    Dim newFile As String
    newFile = tempFolder + "\" + fileName
    GetNewFileInTempFolder = newFile
End Function

Public Function GetTempFileName() As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim tempFileName As String
    tempFileName = FSO.GetTempName()
    GetTempFileName = tempFileName
End Function

Public Function GetTempFolder() As String
    Static tempFolderPath As String

    If Len(tempFolderPath) = 0 Then
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        Const tempFolderFlag = 2
        tempFolderPath = fs.GetSpecialFolder(tempFolderFlag)
    End If
    GetTempFolder = tempFolderPath
End Function

