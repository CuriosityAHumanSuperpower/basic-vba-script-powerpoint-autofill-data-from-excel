VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub PPTcreation()
'----------------------------------------------
'DEFINITIONS
'----------------------------------------------
'Creation PPT
    Dim PPT As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
'Open PPT and create new presentation
    Set PPT = New PowerPoint.Application
    Set PPTPres = PPT.Presentations.Add
    PPT.Visible = True
    PPT.ActiveWindow.ViewType = ppViewSlide
'Loop for
    Dim i As Integer
    Dim Nlig As Integer
    Dim j As Integer
    Dim Ncol As Integer
'Number of slides and items
    Ncol = Range("'EXECUTION'!C2").Value + 2
    Nlig = Range("'EXECUTION'!C3").Value + 2

'Variable
    Dim contenuCell As String
    Dim L As Integer 'width
    Dim H As Integer 'height
    Dim Police As String
    Dim textSize As Integer
    Dim X As Integer 'position from the left side of the slide
    Dim Y As Integer 'position from de top side of the slide
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim fileAddress As String
'----------------------------------------------
'LOOP
'----------------------------------------------
For i = 3 To Nlig
'Add new slide
    Set PPTSlide = PPT.ActivePresentation.Slides.Add(PPT.ActivePresentation.Slides.Count + 1, ppLayoutBlank)
'    PPT.ActiveWindow.View.GotoSlide PPT.ActivePresentation.Slides.Count
    Set PPTSlide = PPT.ActivePresentation.Slides(PPT.ActivePresentation.Slides.Count)

'Add information to slide i
    For j = 3 To Ncol
        'Gathering informations
        contenuCell = Worksheets(1).Cells(i, j).Value
        X = Worksheets(2).Cells(4, j).Value * 28
        Y = Worksheets(2).Cells(5, j).Value * 28
        L = Worksheets(2).Cells(6, j).Value * 28
        H = Worksheets(2).Cells(7, j).Value * 28
        'Test if it is adresse to add an image
        If Left(contenuCell, 2) = "//" Then
        'Image
            fileAddress = Mid(contenuCell, 3, Len(contenuCell))
            'MsgBox (fileAddress)
            PPTSlide.Shapes.AddPicture Filename:=fileAddress, LinkToFile:=msoTrue, SaveWithDocument:=msoTrue, Left:=X, Top:=Y, Width:=L, Height:=H
        'Text
        Else
            Police = Worksheets(2).Cells(9, j).Value
            Red = Worksheets(2).Cells(10, j).Value
            Green = Worksheets(2).Cells(11, j).Value
            Blue = Worksheets(2).Cells(12, j).Value
            textSize = Worksheets(2).Cells(13, j).Value
            With PPTSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=X, Top:=Y, Width:=L, Height:=H).TextFrame.TextRange
                .Text = contenuCell
                .Font.Size = textSize
                .Font.Name = Police
                .Font.Color = RGB(Red, Green, Blue)
                If Worksheets(2).Cells(14, j).Value = 1 Then
                .Font.Bold = True
                End If
                If Worksheets(2).Cells(15, j).Value = 1 Then
                .Font.Italic = True
                End If
                If Worksheets(2).Cells(16, j).Value = 1 Then
                .Font.Underline = True
                End If
            End With
        End If
    Next j
Next i
End Sub



