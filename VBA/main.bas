Attribute VB_Name = "main"
Option Explicit

Const WS_CAPTURE_QUESTIONS As String = "テスト問題取込"
Const WS_TEMPORARY As String = "作業シート"
Const WS_QUESTION_AND_ANSWER As String = "解答"
Const WS_QUESTION As String = "問題"
Const NUMBER As String = "番号"
Const TOKEN As String = "単語"
Const TRANSLATION As String = "訳"
Const RANDOM As String = "乱数"

Sub CleanCells()
    
    Sheets(WS_CAPTURE_QUESTIONS).Range("A:D").clear
    
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

End Sub

Sub SelectQuestionsToUse()

    Sheets(WS_CAPTURE_QUESTIONS).Select
    
    Dim beginNum As Integer, endNum As Integer
    Dim maxRow As Integer
    
    beginNum = Cells(8, "F")
    endNum = Cells(8, "G")
    maxRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    If CheckValue(beginNum, endNum, maxRow) Then
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
        
        Sheets(WS_TEMPORARY).Select
        ActiveSheet.Cells.clear
        
        Dim questionColumns As Variant
        questionColumns = Array(NUMBER, TOKEN, TRANSLATION, RANDOM)
        
        Dim questionRange As Long
        questionRange = (endNum - beginNum) + 1

        For i = LBound(questionColumns) To UBound(questionColumns)
            For j = 1 To questionRange
                Cells(j, i + 1) = targetCollection.Item(questionColumns(i))(j)
            Next j
        Next i
        
        Range(Cells(1, "A"), Cells(questionRange, "D")).Sort _
            Key1:=Range("D1"), _
            order1:=xlAscending, _
            Header:=xlNo, _
            ordercustom:=1, _
            MatchCase:=False, _
            Orientation:=xlTopToBottom, _
            SortMethod:=xlPinYin
            
        Range("D1:D" & CStr(questionRange)).clear
        Sheets(WS_CAPTURE_QUESTIONS).Range("D1:D" & CStr(maxRow)).clear
    Else
        End
    End If

End Sub

Sub ExtractExamQuestion()

    Sheets(WS_TEMPORARY).Select

    Dim questionNum As Long
    ' 入力された値を取得するが、とりあえず100だけを想定
    questionNum = Sheets(WS_CAPTURE_QUESTIONS).Cells(9, "F")
    
    Dim examQuestion As Variant
    ReDim examQuestion(1 To questionNum, 1 To 3)
    
    Dim i As Integer, j As Integer
    For i = LBound(examQuestion, 1) To UBound(examQuestion, 1)
        For j = LBound(examQuestion, 2) To UBound(examQuestion, 2)
            examQuestion(i, j) = Cells(i, j)
        Next j
    Next i
    
    ActiveSheet.Cells.clear
    
    ' 50問ずつに分ける
    'Dim questionCollection As New Collection
    Dim halfQuestionNum As Long
    halfQuestionNum = questionNum / 2
    
    For i = LBound(examQuestion, 1) To halfQuestionNum
        For j = LBound(examQuestion, 2) To UBound(examQuestion, 2)
            Cells(i, j) = examQuestion(i, j)
        Next j
    Next i
    
    For i = halfQuestionNum + 1 To UBound(examQuestion, 1)
        For j = LBound(examQuestion, 2) To UBound(examQuestion, 2)
            Cells(i - 50, j + 3) = examQuestion(i, j)
        Next j
    Next i

End Sub

Sub CreateQuestionAndAnswer()
    
    Sheets(WS_QUESTION_AND_ANSWER).Select
    ActiveSheet.Cells.clear
    
    Sheets(WS_TEMPORARY).Cells.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues, operation:=xlNone
    Application.CutCopyMode = False
    
    Cells(1, 1).EntireRow.Insert
    
    Dim questionColumns As Variant
    questionColumns = Array(NUMBER, TOKEN, TRANSLATION)
    
    Dim i As Integer, j As Integer
    For i = LBound(questionColumns) To UBound(questionColumns)
        Cells(1, i + 1) = questionColumns(i)
        Cells(1, i + 4) = questionColumns(i)
    Next i
    
    '以下シート整形
    Dim maxRow As Long, maxCol As Long
    maxRow = Cells(Rows.Count, 1).End(xlUp).Row
    maxCol = Cells(1, Columns.Count).End(xlToLeft).column
    
    Range(Cells(1, 1), Cells(maxRow, maxCol)).Borders.LineStyle = xlContinuous
    '現時点ではAutoFitを使用しているが、列幅は固定値に変更予定
    Range(Cells(1, 1), Cells(maxRow, maxCol)).EntireColumn.AutoFit
    
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1).EntireColumn.Insert
    
    Columns("D").ColumnWidth = 40
    Columns("G").ColumnWidth = 40
    
    Sheets(WS_TEMPORARY).Cells.clear
    
End Sub

Sub CreateExamQuestion()

    Sheets(WS_QUESTION_AND_ANSWER).Select
    Sheets(WS_QUESTION).Cells.clear

    Cells.Copy
    Sheets(WS_QUESTION).Range("A1").PasteSpecial Paste:=xlPasteAll, operation:=xlNone
    Application.CutCopyMode = False

    Sheets(WS_QUESTION).Select

    Dim maxRow As Long
    maxRow = Cells(Rows.Count, "B").End(xlUp).Row

    '訳部分を削除
    Union(Range(Cells(3, "D"), Cells(maxRow, "D")), Range(Cells(3, "G"), Cells(maxRow, "G"))).ClearContents

End Sub

Function CheckValue(ByVal beginNum As Long, ByVal endNum As Long, ByVal maxRow As Long) As Boolean

    Dim msg As String
    msg = ""

    If IsNull(beginNum) Or IsNull(endNum) Then
        msg = "取得したい問題の番号を入力してください。"
    ElseIf IsNumeric(beginNum) = False Or IsNumeric(endNum) = False Then
        msg = "数値を入力してください。"
    ElseIf beginNum = 0 Or endNum = 0 Then
        msg = "1以上の整数を入力してください。"
    ElseIf endNum > maxRow Then
        msg = "問題は" & maxRow - 1 & "問までしかありません。"
    End If
    
    If msg = "" Then
        CheckValue = True
    Else
        msgbox msg
        CheckValue = False
    End If

End Function

' 100問より多くても対応するか処理は後回し
'Function hoge(ByRef examQuestions As Variant, ByVal questionNumber As Long) As Collection
'
'Dim resultCollection As New Collection
'Dim tmpArray As Variant
'
'' questionNumber / 2 が50で割り切れた場合の処理
'' questionNumber / 2 が50で割り切れた場合の処理
'
'
'End Function



