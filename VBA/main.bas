Attribute VB_Name = "Module1"
Option Explicit

Sub clean_cells()
    
    Sheets("テスト問題取込").Cells.clear
    
    msgbox "テスト問題を削除しました"
    
End Sub

Sub import_examination_questions()

    Dim wsImport As Worksheet
    Set wsImport = Worksheets("テスト問題取込")
    
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
