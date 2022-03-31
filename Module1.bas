Attribute VB_Name = "Module1"
Sub main()

    Dim masterFilePath As String
    Dim examDirPath As String
    
    Dim macroFile As String
    Dim masterFile As String
    Dim examFile As String
    
    Dim selectionList() As Variant
    
    masterFilePath = Cells(5, 2).Value
    examDirPath = Cells(6, 2).Value
    
    macroFile = ActiveWorkbook.Name
    
'Master file check
    Call fileCheck(masterFilePath)
    
'Exam file check
    buf = Dir(examDirPath & "*.xlsx")
    Do While buf <> ""
        Call fileCheck(examDirPath + buf)
        buf = Dir()
    Loop

'Master file open
    Workbooks.Open masterFilePath
    masterFile = Right(masterFilePath, InStrRev(masterFilePath, "\"))
    
'Exam file open execute loop
    examFile = Dir(examDirPath & "*.xlsx")
    Do While examFile <> ""
        Workbooks.Open examDirPath & examFile
        
        
        'Exam file latest recent
        Workbooks(examFile).Activate
        Worksheets("aaa").Activate
        Cells(Rows.Count - 1, 2).End(xlUp).Select
        Range(Selection, Selection.End(xlToRight)).Select
        If (Selection.Count > 10) Then
            MsgBox "âΩÇ©Ç®Ç©ÇµÇ¢"
            End
        End If
        selectionList = Selection
        
        'Master file same recent Check
        Workbooks(masterFile).Activate
        Worksheets("aaa").Activate
        
        Cells(Rows.Count - 1, 2).End(xlUp).Select
        
        latestUpdated = True
        Do While True
            matchRow = True
            i = 0
            For Each word In selectionList
                If (word = Selection.Offset(0, i).Value) Then
                    i = i + 1
                Else
                    latestUpdated = False
                    matchRow = False
                    Exit For
                End If
            Next
            If (matchRow) Then
                Exit Do
            End If
            Selection.Offset(-1, 0).Select
            If (Selection.Row < 3) Then
                Exit Do
            End If
        Loop
        If (latestUpdated) Then
            MsgBox examFile + "ÇÕóöóÇ™ä˘Ç…ç≈êVÇ≈Ç∑"
            Workbooks(examFile).Close SaveChanges:=False
            Exit Do
        End If

        'Recent copy paste
        Range(Selection, Selection.End(xlToRight).End(xlDown)).Select
        Selection.Copy
        
        Workbooks(examFile).Activate
        Worksheets("aaa").Activate
        Cells(Rows.Count - 1, 2).End(xlUp).Select
        Selection.PasteSpecial
        
                
        'Exam file latest Exam
        Workbooks(examFile).Activate
        Worksheets("bbb").Activate
        
        Call filterOFF
        
        Cells(Rows.Count - 1, 1).End(xlUp).Select
        Range(Selection, Selection.Offset(0, 2)).Select
        selectionList = Selection
        
        'Master file same exam Check
        Workbooks(masterFile).Activate
        Worksheets("bbb").Activate
        
        Cells(Rows.Count - 1, 1).End(xlUp).Select
        
        latestUpdated = True
        Do While True
            matchRow = True
            i = 0
            For Each word In selectionList
                If (word = Selection.Offset(0, i).Value) Then
                    i = i + 1
                Else
                    latestUpdated = False
                    matchRow = False
                    Exit For
                End If
            Next
            If (matchRow) Then
                Exit Do
            End If
            Selection.Offset(-1, 0).Select
            If (Selection.Row < 3) Then
                Exit Do
            End If
        Loop
        If (latestUpdated) Then
            MsgBox examFile + "ÇÕééå±ï\Ç™ä˘Ç…ç≈êVÇ≈Ç∑"
            Workbooks(examFile).Close SaveChanges:=False
            Exit Do
        End If
        
        'Recent copy paste
        Range(Selection, Selection.Offset(0, 2).End(xlDown)).Select
        Selection.Copy
        
        Workbooks(examFile).Activate
        Worksheets("bbb").Activate
        Cells(Rows.Count - 1, 1).End(xlUp).Select
        Selection.PasteSpecial
                
        
        Call filterOn("D3", "ÅZ")
        Workbooks(examFile).Close SaveChanges:=True
        
        examFile = Dir()
    Loop
    Workbooks(masterFile).Close SaveChanges:=False

End Sub


Sub fileCheck(filePath)

    fileName = Right(masterFilePath, InStrRev(masterFilePath, "\"))
    
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = fileName Then
            MsgBox fileName + "Ç™ä˘Ç…äJÇ¢ÇƒÇ¢Ç‹Ç∑"
            End
        End If
    Next wb
    
    If Not Dir(filePath) <> "" Then
        MsgBox "Not found :" + filePath, vbExclamation
        End
    End If

End Sub


Sub filterOFF()
Attribute filterOFF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' filterOFF Macro
'

'
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    

End Sub

Sub filterOn(x, y)
'
' filterOn Macro
'

'
    If (x = "") Then
        x = "A1"
    End If
    If (y = "") Then
        y = "ÅZ"
    End If
    
    f = Range(x).Column

    ActiveSheet.Range(x).AutoFilter Field:=f, Criteria1:=y

End Sub
