Attribute VB_Name = "TestModule1"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Set SDI to 1, because Rubber duck has issues opening new .dwgs
    ' when Acad is not in SDI mode 1
    ThisDrawing.SetVariable "SDI", 1
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestGetActiveLayerName()
    Dim layerName As String
    Dim layer As AcadLayer
    Set layer = ThisDrawing.ActiveLayer
    layerName = layer.Name
    Assert.isTrue True
End Sub

'@TestMethod
Public Sub TestInsertDrawingAsBlockAndDelete_ShouldPass()
    Dim appAcad As AcadApplication
    Set appAcad = GetAcad
    
    
    Dim blockName As String
    blockName = "testBlock"
    Dim blockObject As AcadBlock
    Set blockObject = ThisDrawing.Blocks.Add(Utilities.ConstantInsertionPointZeroZeroZero(), blockName)
    
    blockObject.AddAttribute 5#, acAttributeModeNormal, "promptString", Utilities.ConstantInsertionPointZeroZeroZero(), "tag", "value"
    blockObject.AddLine Utilities.ConstantInsertionPointZeroZeroZero(), Utilities.ConstantInsertionPointOneOneOne()
    
    
    Dim blockReference As AcadBlockReference
    Set blockReference = ThisDrawing.ModelSpace.InsertBlock(ConstantInsertionPointOneOneOne(), blockName, 1#, 1#, 1#, 0#)
    
    
    Dim blockDrawingPath As String
    blockDrawingPath = SaveCurrentDwgToTempFolder(ThisDrawing)
    
    Dim newDoc As AcadDocument
    ' use this method to open new document is SDI = 1
    Set newDoc = ThisDrawing.New("acad.dwt")
    
    InsertDrawingThenDelete newDoc, blockDrawingPath
    Set blockReference = newDoc.ModelSpace.InsertBlock(ConstantInsertionPointOneOneOne(), blockName, 1#, 1#, 1#, 0#)
    
    Assert.isTrue True
End Sub


Public Function IsAutoCADQuiescent(appAcad As AcadApplication) As Integer
    On Error GoTo ErrorHandler
    Dim State As AcadState
    IsAutoCADQuiescent = -1
    Set State = appAcad.GetAcadState
    DoEvents
    While (State.IsQuiescent <> True)
        DoEvents
        Set State = appAcad.GetAcadState
    Wend
    Set State = Nothing
    IsAutoCADQuiescent = 0
    Exit Function
ErrorHandler:
    MsgBox "Error: " & Err.Number & " :: " & Err.Description
End Function

