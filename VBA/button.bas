Attribute VB_Name = "button"
Option Explicit

Sub �e�X�g���捞()

    main.CleanCells
    
    main.ImportExaminationQuestions

End Sub

Sub �e�X�g���쐬()

    main.SelectQuestionsToUse
    
    main.ExtractExamQuestion
    
    main.CreateQuestionAndAnswer
    
    main.CreateExamQuestion
    
End Sub
