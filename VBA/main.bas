Attribute VB_Name = "main"
Option Explicit

Const WS_CAPTURE_QUESTIONS As String = "�e�X�g���捞"
Const WS_TEMPORARY As String = "��ƃV�[�g"
Const NUMBER As String = "�ԍ�"
Const TOKEN As String = "�P��"
Const TRANSLATION As String = "��"
Const RANDOM As String = "����"

Sub CleanCells()
    
    Sheets(WS_CAPTURE_QUESTIONS).Range("A:D").clear
    
    msgbox "�e�X�g�����폜���܂���"
    
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
    
    msgbox "�e�X�g���S�̂���荞�݂܂���"

End Sub

Sub SelectQuestionsToUse()

    Sheets(WS_CAPTURE_QUESTIONS).Select
    
    Dim beginNum As Integer, endNum As Integer
    Dim endRow As Integer
    Dim msg As String
    
    beginNum = Cells(8, "E")
    endNum = Cells(8, "F")
    endRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    If IsNull(beginNum) Or IsNull(endNum) Then
        msg = "�擾���������̔ԍ�����͂��Ă��������B"
    ElseIf IsNumeric(beginNum) = False Or IsNumeric(endNum) = False Then
        msg = "���l����͂��Ă��������B"
    ElseIf beginNum = 0 Or endNum = 0 Then
        msg = "1�ȏ�̐�������͂��Ă��������B"
    ElseIf endNum > endRow Then
        msg = "����" & endRow - 1 & "��܂ł�������܂���B"
    Else
        Dim i As Integer, j As Integer
                        
        '�Ώېݖ�Q�ɗ�����ݒ�
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
        
        Dim columns As Variant
        columns = Array(NUMBER, TOKEN, TRANSLATION, RANDOM)
        
        Dim questionNum As Long
        questionNum = (endNum - beginNum) + 1

        For i = LBound(columns) To UBound(columns)
            For j = 1 To questionNum
                Cells(j, i + 1) = targetCollection.Item(columns(i))(j)
            Next j
        Next i
        
        Range(Cells(1, "A"), Cells(questionNum, "D")).Sort _
            Key1:=Range("D1"), _
            order1:=xlAscending, _
            Header:=xlNo, _
            ordercustom:=1, _
            MatchCase:=False, _
            Orientation:=xlTopToBottom, _
            SortMethod:=xlPinYin
            
        Range("D1:D" & CStr(questionNum)).clear
        Sheets(WS_CAPTURE_QUESTIONS).Range("D1:D" & CStr(endRow)).clear
        
        msg = "�g�p������͈̔͂𒊏o���܂����B"
    End If
    
    msgbox msg

End Sub
