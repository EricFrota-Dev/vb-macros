Sub InserirTextoEmTodasAsPaginas()
    Dim filePath As String
    Dim fileContent() As String
    Dim fNum As Integer
    Dim lineText As String
    Dim i As Long, totalLines As Long
    Dim index As Long
    Dim pg As Page
    Dim s As Shape

    ' Selecionar o arquivo .txt
    filePath = CorelScriptTools.GetFileBox("Arquivos de Texto|*.txt", "Selecione o arquivo de texto")
    If filePath = "" Then
        MsgBox "Nenhum arquivo selecionado.", vbExclamation
        Exit Sub
    End If

    ' Ler o conteúdo com suporte a acentuação
    fNum = FreeFile
    Open filePath For Input As #fNum
    lineText = Input$(LOF(fNum), fNum)
    Close #fNum

    ' Quebra em linhas
    fileContent = Split(Replace(lineText, vbCrLf, vbLf), vbLf)
    totalLines = UBound(fileContent)
    index = 0

    ' Para cada página, coletar e processar os campos de texto em ordem invertida
    For Each pg In ActiveDocument.Pages
        Dim textFields As New Collection

        ' Coletar todos os ParagraphText da página atual
        For Each s In pg.shapes
            If s.Type = cdrTextShape Then
                If s.Text.Type = cdrParagraphText Then
                    textFields.Add s
                End If
            End If
        Next s

        ' Preencher os textos em ordem reversa dentro da página
        For i = textFields.count To 1 Step -1
            If index > totalLines Then Exit Sub

            Set s = textFields(i)
            s.Text.Story = fileContent(index)
            s.Text.Frame.VerticalAlignment = cdrAlignVerticalCenter

            index = index + 1
        Next i
    Next pg

    MsgBox "Texto inserido com sucesso em todas as páginas!", vbInformation
End Sub
