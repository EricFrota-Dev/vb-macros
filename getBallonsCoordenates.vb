Sub ExportMangaToJson()

    Dim doc As Document
    Set doc = ActiveDocument

    Dim jsonOutput As String
    jsonOutput = "["

    Dim p As Page
    Dim skippedPages As Integer
    skippedPages = 0

    For Each p In doc.Pages
    ' INVERTE a ordem dos shapes
    Dim tempShapes As New Collection
    Set tempShapes = Nothing
    Set tempShapes = New Collection
    Dim i As Long
    For i = p.Shapes.Count To 1 Step -1
        Dim s As Shape
        Set s = p.Shapes(i)

        If s.Type = cdrEllipseShape Or s.Type = cdrRectangleShape Then
            tempShapes.Add s
        End If
    Next i

    ' Se não houver balões, pula página
    If tempShapes.Count = 0 Then
        skippedPages = skippedPages + 1
        GoTo ContinueLoop
    End If

    ' Usa o número real da página no CorelDRAW (p.Index)
    Dim pageJson As String
    pageJson = "{""page"":" & p.Index & ",""balloons"":["

    For i = 1 To tempShapes.Count
        Dim b As Shape
        Set b = tempShapes(i)

        Dim x As Double, y As Double, w As Double, h As Double
        x = b.PositionX / p.SizeWidth
        y = ((p.SizeHeight + (b.PositionY - b.SizeHeight)) - p.SizeHeight) / p.SizeHeight
        w = b.SizeWidth / p.SizeWidth
        h = b.SizeHeight / p.SizeHeight

        pageJson = pageJson & _
            "{""type"":""" & IIf(b.Type = cdrEllipseShape, "ellipse", "rectangle") & """," & _
            """x"": " & ToInvariant(x) & ", " & _
            """y"": " & ToInvariant(y) & ", " & _
            """width"": " & ToInvariant(w) & ", " & _
            """height"": " & ToInvariant(h) & "},"
    Next i

    ' Remove última vírgula
    If tempShapes.Count > 0 Then
        pageJson = Left(pageJson, Len(pageJson) - 1)
    End If

    pageJson = pageJson & "]},"
    jsonOutput = jsonOutput & pageJson

ContinueLoop:
Next p


    ' Remove última vírgula final
    If Right(jsonOutput, 1) = "," Then
        jsonOutput = Left(jsonOutput, Len(jsonOutput) - 1)
    End If

    jsonOutput = jsonOutput & "]"

    ' Salvar e abrir JSON
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\manga_export.json"

    Dim fnum As Integer
    fnum = FreeFile
    Open tempPath For Output As #fnum
    Print #fnum, jsonOutput
    Close #fnum

    Shell "notepad.exe """ & tempPath & """", vbNormalFocus

    MsgBox "Exportação concluída!" & vbCrLf & _
           "Páginas puladas (sem balões): " & skippedPages, vbInformation
End Sub

Function ToInvariant(value As Double) As String
    ToInvariant = Replace(Format(value, "0.000"), ",", ".")
End Function
