Attribute VB_Name = "Module1"
Sub ExportIllustrationAsPDF()

    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim folderPath As String

    Set pres = ActivePresentation
    Set sld = ActiveWindow.View.Slide

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a shape or group first.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)

    If pres.Path = "" Then
        folderPath = Environ("USERPROFILE") & "\Desktop\"
    Else
        folderPath = pres.Path & "\"
    End If

    Dim pdfPath As String
    pdfPath = folderPath & shp.Name & ".pdf"

    ' --- Create new presentation FIRST ---
    Dim newPres As Presentation
    Set newPres = Presentations.Add(WithWindow:=msoTrue)
    newPres.PageSetup.SlideWidth = shp.Width
    newPres.PageSetup.SlideHeight = shp.Height

    Dim newSld As Slide
    Set newSld = newPres.Slides.Add(1, ppLayoutBlank)

    ' --- Go back to original, copy shape AFTER new pres is ready ---
    pres.Windows(1).Activate
    shp.Select
    shp.Copy

    ' --- Switch to new pres and paste ---
    newPres.Windows(1).Activate
    newPres.Windows(1).View.GotoSlide 1

    Dim pasted As ShapeRange
    Set pasted = newSld.Shapes.Paste
    pasted.Left = 0
    pasted.Top = 0

    ' --- Export as PDF ---
    newPres.ExportAsFixedFormat2 pdfPath, ppFixedFormatTypePDF

    newPres.Saved = True
    newPres.Close

    MsgBox "Saved to:" & vbCrLf & pdfPath, vbInformation

End Sub
