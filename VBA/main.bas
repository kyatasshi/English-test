Attribute VB_Name = "main"
Option Explicit

Const WS_CAPTURE_QUESTIONS As String = "テスト問題取込"
Const WS_EXPORT_QUESTIONS As String = "Sheet2"
Const NUMBER As String = "番号"
Const TOKEN As String = "単語"
Const TRANSLATION As String = "訳"
Const RANDOM As String = "乱数"

Sub CleanCells()
    
    Sheets(WS_CAPTURE_QUESTIONS).Range("A:D").clear
    
    msgbox "テスト問題を削除しました"
    
End Sub

Sub ImportExaminationQuestions()

    Dim wsImport As Worksheet
    Set wsImport = Worksheets(WS_CAPTURE_QUESTIONS)
    
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\target_word_list.csv"
    
    Dim queryTable As queryTable
    Set queryTable = wsImport.QueryTables.Add(Connection:="TEXT;" & filePath, _
                                                Destination:=wsImport.Range("A1"))
                                          
    With queryTable
        .TextFilePlatform = 932
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .RefreshStyle = xlOverwriteCells
        .Refresh
        .Delete
    End With
    
    msgbox "テスト問題を取り込みました"

End Sub

Sub ExtractQuestions()

    Sheets(WS_CAPTURE_QUESTIONS).Select
    
    Dim beginNum As Integer, endNum As Integer
    Dim endRow As Integer
    Dim msg As String
    
    beginNum = Cells(8, "E")
    endNum = Cells(8, "F")
    endRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    If IsNull(beginNum) Or IsNull(endNum) Then
        msg = "取得したい問題の番号を入力してください。"
    ElseIf IsNumeric(beginNum) = False Or IsNumeric(endNum) = False Then
        msg = "数値を入力してください。"
    ElseIf beginNum = 0 Or endNum = 0 Then
        msg = "1以上の整数を入力してください。"
    ElseIf endNum > endRow Then
        msg = "問題は" & endRow - 1 & "問までしかありません。"
    Else
        Dim i As Integer, j As Integer
        
        '対象設問群に乱数を設定
        Randomize
        For i = beginNum To endNum
            Range(Cells(i + 1, "D"), Cells(i + 1, "D")) = Rnd()
        Next i
        
        Dim targetCollection As New Collection
        
        With targetCollection
            .Add Range(Cells(beginNum + 1, "A"), Cells(endNum + 1, "A")), NUMBER
            .Add Range(Cells(beginNum + 1, "B"), Cells(endNum + 1, "B")), TOKEN
            .Add Range(Cells(beginNum + 1, "C"), Cells(endNum + 1, "C")), TRANSLATION
            .Add Range(Cells(beginNum + 1, "D"), Cells(endNum + 1, "D")), RANDOM
        End With
        
        Sheets(WS_EXPORT_QUESTIONS).Select
        
        Dim columns As Variant
        columns = Array(NUMBER, TOKEN, TRANSLATION, RANDOM)
        
        For i = LBound(columns) To UBound(columns)
            For j = 1 To (endNum - beginNum) + 1
                Cells(j, i + 1) = targetCollection.Item(columns(i))(j)
            Next j
        Next i
        
        msg = "使用する問題を作成しました。"
        
        'Dim targetRange As Range
        'Dim targetData As Variant
        'Set targetRange = Range(Cells(beginNum + 1, "A"), Cells(endNum + 1, "C"))
        'targetData = targetRange
        
        'For i = LBound(targetData, 1) To UBound(targetData, 1)
        '    For j = LBound(targetData, 2) To UBound(targetData, 2)
        '        Cells(i, j) = targetData(i, j)
        '    Next j
        'Next i
    End If
    
    msgbox msg

End Sub

