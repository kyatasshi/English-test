Attribute VB_Name = "button"
Option Explicit

Sub テスト問題取込()

    main.CleanCells
    
    main.ImportExaminationQuestions

End Sub

Sub テスト問題作成()

    main.SelectQuestionsToUse
    
    main.ExtractExamQuestion
    
    main.CreateQuestionAndAnswer
    
    main.CreateExamQuestion
    
End Sub
