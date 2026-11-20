Attribute VB_Name = "ColorCitations"
' Colours every in-text citation in the document body blue:
'   - Word's own citations (CITATION fields)
'   - ReadCube / Papers SmartCite citations (content controls titled SmartCite Citation)
' The bibliography is not changed. Run it again after the citations are refreshed,
' because Word or SmartCite may redraw them in black.

Option Explicit

Private Const CITE_COLOR As Long = 12611584   ' RGB(0, 112, 192)

Public Sub ColorCitationsBlue()
    Dim f As Field, cc As ContentControl
    For Each f In ActiveDocument.Fields
        If f.Type = wdFieldCitation Then f.Result.Font.Color = CITE_COLOR
    Next
    For Each cc In ActiveDocument.ContentControls
        If cc.Title = "SmartCite Citation" Then cc.Range.Font.Color = CITE_COLOR
    Next
End Sub
