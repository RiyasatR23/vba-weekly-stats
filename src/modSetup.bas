Attribute VB_Name = "modSetup"
'==============================================================================
' Run AddReportButton once after importing the modules. It puts a
' "Generate weekly report" button on the Control sheet.
'==============================================================================
Option Explicit

Public Sub AddReportButton()
    Dim sh As Worksheet, shp As Shape
    Set sh = ThisWorkbook.Worksheets("Control")

    On Error Resume Next
    sh.Shapes("btnGenerate").Delete
    On Error GoTo 0

    Set shp = sh.Shapes.AddShape(msoShapeRoundedRectangle, sh.Range("B12").Left, sh.Range("B12").Top, 220, 40)
    With shp
        .Name = "btnGenerate"
        .OnAction = "GenerateWeeklyReport"
        .Fill.ForeColor.RGB = RGB(75, 43, 214)
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "Generate weekly report"
            .Font.Bold = msoTrue
            .Font.Size = 13
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    MsgBox "Button added to the Control sheet.", vbInformation
End Sub
