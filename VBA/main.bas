Attribute VB_Name = "Module1"
Option Explicit

Sub clean_cells()
    
    Sheets("�e�X�g���捞").Cells.clear
    
    msgbox "�e�X�g�����폜���܂���"
    
End Sub

Sub import_examination_questions()

    Dim wsImport As Worksheet
    Set wsImport = Worksheets("�e�X�g���捞")
    
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
    
    msgbox "�e�X�g������荞�݂܂���"

End Sub
