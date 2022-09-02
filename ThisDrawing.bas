VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Sub Hello()
    Dim appAcad As AcadApplication
    Set appAcad = GetObject(, "AutoCAD.Application")

    Dim newDoc As AcadDocument
    Set newDoc = ThisDrawing.Application.Documents.Add("acad.dwt")
    'IsAutoCADQuiescent appAcad
    Dim test As String
    test = ""
End Sub
