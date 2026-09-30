Attribute VB_Name = "extractToWord"
' ==========================================================
' vbaWord: Word Form Batch Summarizer (Output to Word)
' Copy source ranges so question and answer formatting is preserved.
' ==========================================================

Sub SummarizeToNewWordDoc()
    Dim wdApp As Object, wdDoc As Object, targetDoc As Object
    Dim sectionDoc As Object, questionRange As Object, answerRange As Object
    Dim fd As FileDialog, fileItem As Variant, ff As Object, dict As Object
    Dim key As Variant, rawFileName As String, displayName As String
    Dim qText As String, i As Long, prevEnd As Long, hasQuestionText As Boolean
    Dim pType As Long, isUnprotected As Boolean, docPwd As String

    docPwd = modConfig.GetDocPassword()

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "请选择旧式窗体问卷汇总至 Word"
        .Filters.Clear: .Filters.Add "Word Documents", "*.doc; *.docx; *.docm", 1

        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                rawFileName = Dir(fileItem)
                displayName = GetMappedName(rawFileName)

                Set wdDoc = wdApp.Documents.Open(Filename:=fileItem, ReadOnly:=False, Visible:=False)

                pType = wdDoc.ProtectionType: isUnprotected = False
                If pType <> -1 Then
                    On Error Resume Next
                    Err.Clear
                    wdDoc.Unprotect Password:=docPwd
                    If Err.Number = 0 Then isUnprotected = True
                    On Error GoTo 0
                End If

                prevEnd = 0
                For i = 1 To wdDoc.FormFields.Count
                    Set ff = wdDoc.FormFields(i)
                    Set questionRange = wdDoc.Range(prevEnd, ff.Range.Start)
                    qText = questionRange.Text
                    qText = Replace(qText, vbCr, ""): qText = Replace(qText, vbLf, "")
                    qText = Replace(qText, ":", ""): qText = Replace(qText, "：", "")
                    qText = Trim(qText)

                    hasQuestionText = (qText <> "")
                    If Not hasQuestionText Then qText = ff.Name

                    If Not dict.Exists(qText) Then
                        Set sectionDoc = wdApp.Documents.Add(Visible:=False)
                        dict.Add qText, sectionDoc
                        If hasQuestionText Then
                            AppendFormattedRange sectionDoc, questionRange
                            If Right$(questionRange.Text, 1) <> vbCr Then AppendPlainText sectionDoc, vbCr
                        Else
                            AppendPlainText sectionDoc, qText & vbCr
                        End If
                    Else
                        Set sectionDoc = dict(qText)
                    End If

                    AppendPlainText sectionDoc, "【" & displayName & "】: "
                    If ff.Range.Fields.Count > 0 Then
                        Set answerRange = ff.Range.Fields(1).Result
                        AppendFormattedRange sectionDoc, answerRange
                    Else
                        AppendPlainText sectionDoc, ff.Result
                    End If
                    AppendPlainText sectionDoc, "; " & vbCr
                    prevEnd = ff.Range.End
                Next i

                If pType <> -1 And isUnprotected Then
                    On Error Resume Next
                    wdDoc.Protect Type:=pType, NoReset:=True, Password:=docPwd
                    On Error GoTo 0
                End If

                wdDoc.Close SaveChanges:=False
                Set wdDoc = Nothing
            Next fileItem

            Set targetDoc = wdApp.Documents.Add
            For Each key In dict.Keys
                Set sectionDoc = dict(key)
                AppendFormattedRange targetDoc, sectionDoc.Content
                sectionDoc.Close SaveChanges:=False
            Next key
            wdApp.Visible = True
            MsgBox "汇总完成！", vbInformation
        End If
    End With
End Sub

Private Sub AppendFormattedRange(ByVal destinationDoc As Object, ByVal sourceRange As Object)
    Dim insertion As Object
    Set insertion = destinationDoc.Range(destinationDoc.Content.End - 1, destinationDoc.Content.End - 1)
    insertion.FormattedText = sourceRange.FormattedText
End Sub

Private Sub AppendPlainText(ByVal destinationDoc As Object, ByVal value As String)
    Dim insertion As Object
    Set insertion = destinationDoc.Range(destinationDoc.Content.End - 1, destinationDoc.Content.End - 1)
    insertion.Text = value
End Sub

Private Function GetMappedName(originalName As String) As String
    Dim mapWs As Worksheet: Dim i As Long: GetMappedName = originalName

    On Error Resume Next: Set mapWs = ThisWorkbook.Sheets("mapping"): On Error GoTo 0
    If Not mapWs Is Nothing Then
        For i = 2 To mapWs.Cells(mapWs.Rows.Count, "A").End(xlUp).Row
            If InStr(1, originalName, mapWs.Cells(i, 1).Value, vbTextCompare) > 0 Then
                GetMappedName = mapWs.Cells(i, 2).Value: Exit Function
            End If
        Next i
    End If
End Function
