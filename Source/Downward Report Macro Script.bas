Attribute VB_Name = "Module3"
Sub CustomFormReportingStepsReport()
Attribute CustomFormReportingStepsReport.VB_Description = "This macro prompts the user for an employee ID or EMail and the number of reporting steps to display an creates a report showing all employees who report to that employee within the designated number of steps."
Attribute CustomFormReportingStepsReport.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' CustomFormReportingStepsReport Macro
' This macro prompts the user for an employee ID or EMail and the number of reporting steps to display an creates a report showing all employees who report to that employee within the designated number of steps.
'
' Keyboard Shortcut: Ctrl+d
'
    'Select the sheet that has the Headcount Report Info. Note that this should be named 'sheet1' for the script to work
    Sheets("Sheet1").Select
    Rows("1:1").Select
    
    InputForm.Show
    
End Sub

Function CustomConvertEmail(input_text)
        'Function to convert email input to Employee ID
        
        'error handling
        On Error GoTo InvalidEmail
        
        IDColumn = findCol("Empl ID")
        Emp_ID_Path = "$" & IDColumn & "$1:$" & IDColumn & Row_Length()
        
        Set Search = ActiveSheet.Range(Emp_ID_Path)
        Set searchlast = Search.Cells(Search.Cells.Count)
        Dim FindRow As Range
        
        EmailColumn = findCol("Email")
        Set FindRow = ActiveSheet.Range(EmailColumn & ":" & EmailColumn).Find(What:=input_text, LookIn:=xlValues)
        Dim rowOfEmail As Integer
        rowOfEmail = FindRow.Row
        search_text = ActiveSheet.Range(IDColumn & rowOfEmail).Value
        CustomConvertEmail = search_text
        Exit Function
        
InvalidEmail:
        Dim msg As String
        msg = input_text & " not found. Please enter a new search criteria."
        MsgBox (msg)
        InputForm.Show
       
        
End Function

Sub findReports(search_text, stepLimit)

    'Find the employee that relates to the entered search criteria
    Column = findCol("Empl ID")
    Emp_ID_Path = "$" & Column & "$1:$" & Column & Row_Length()
    Set Search = ActiveSheet.Range(Emp_ID_Path)
    Set searchlast = Search.Cells(Search.Cells.Count)
    Set foundValue = ActiveSheet.Range(Emp_ID_Path).Find(search_text, searchlast, xlValues, LookAt:=xlWhole)
    
    EmpID = search_text
    
    If foundValue Is Nothing Then
        MsgBox (search_text & " was unable to be found. Please ensure that you entered the correct email or Employee ID. Remember to only include the digits in the employee ID.")
        'Call ReportingStepsReport 'this was used to restart Module 2 during testing
        Call CustomFormReportingStepsReport
    Else
        'Create top level collection to hold all subsequent levels of reports. each individual in tree will have their own collection
        Dim List As Collection
        Set List = dirReportLookup(EmpID)
              
       'Create Reporting Sheet
       reportName = createReport(EmpID)
       
       'Assign the subordinate count to a variable.
       subCount = subordinateCount(EmpID, stepLimit, 0, 0, reportName) - 1
       
       'Open the Report sheet to view results
       Sheets(reportName).Select
       
       'Autofit column width for ease of reading
        ActiveSheet.Columns("A:z").ColumnWidth = 40
        ActiveSheet.Rows("1:" & Row_Length()).AutoFit
        ActiveSheet.Columns("A:z").AutoFit
        
        'ActiveSheet.Range("A1:z" & Row_Length()).EntireRow.AutoFit
        
       'Run Reporting Hierarchy Report based on Tier
       MsgBox ("There are " & subCount & " Employees who report up to " & EmpID & " within " & stepLimit & " reporting levels.")
              
    End If
End Sub

Function dirReportLookup(EmpID) As Collection
'Return a Collection containing all direct reports to the employee passed as a parameter

    Dim rFND As Range
    Dim sFirstAddress
    
    EmpIDColumn = findCol("Empl ID")
    
    Dim empList As New Collection
    
    'Set search range to just be the 'Supv ID' column
    SupvColumn = findCol("Supv ID")
    Emp_ID_Path = "$" & SupvColumn & "$1:$" & SupvColumn & Row_Length()
    Set SearchRange = Sheets("sheet1").Range(Emp_ID_Path)
    
    'loop through and find all employees who report to the searched employee
    With SearchRange
        Set rFND = .Cells.Find(What:=EmpID, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not rFND Is Nothing Then
            sFirstAddress = rFND.Address
            Do
                Row = rFND.Row
                Dim Emp As cSupervisor
                Set Emp = New cSupervisor
                Emp.ID = ActiveSheet.Range(EmpIDColumn & Row)
                empList.Add (Emp.ID)
                Set rFND = .FindNext(rFND)
            Loop While Not rFND Is Nothing And rFND.Address <> sFirstAddress
        End If
    End With
    
    Set dirReportLookup = empList
End Function


Function subordinateCount(EmpID, FinalStep, CurrentStep, Count, reportName) As Integer
    'Returns the number of employees that report up to the employee (EmpID) indicated within the specified number of sub-supervisors (FinalStep), including the searched
    'Employee (so subtract 1 from the result for a true accounting of subordinates. The 'Count' and 'CurrentStep' parameters should be 0 when being called
    'for the first time. These variables are used for recursive purposes.
    'For example, if Manager A has two reports (Supv B and Supv C) and Supv C has 2 reports (Emp D and Emp E) then there would be 2 individuals within two 'steps' of Manager A (B,C)
    ' and there would be 4 individuals within two 'steps' of Manager A (B,C,D,E).
    
    Dim List As Collection
    Set List = dirReportLookup(EmpID)
    Count = Count + 1
    
    'MsgBox ("reportName: " & reportName & " Count: " & Count & " Steps: " & Steps & " EmpID: " & EmpID)
    Sheets(reportName).Cells(Count, CurrentStep + 1).Value = emp_Name(EmpID) & Chr(10) & emp_Title(EmpID) & Chr(10) & EmpID
    
    'MsgBox (CurrentStep & " v. " & FinalStep)
    If (CurrentStep < FinalStep) Then
        For Each DR In List
            Count = subordinateCount(DR, FinalStep, CurrentStep + 1, Count, reportName)
        Next DR
               
        subordinateCount = Count
        
    Else
        subordinateCount = Count
        
    End If
        
End Function


Function createReport(EmpID)
    'Creates a new sheet for the search results to be displayed on
    
    'Set new sheet's name as a variable
    Dim NewSheetName As String
    NewSheetName = (EmpID & "_Subordinate_Report")
    
    'Delete any existing worksheets with the same name as the one about to be created
    If (SheetExists(NewSheetName)) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(EmpID & "_Subordinate_Report").Delete
        Application.DisplayAlerts = True
    End If
    
    'Create new worksheet
    Dim newsheet
    Set newsheet = Sheets.Add(After:=Sheets(Worksheets.Count), Count:=1, Type:=xlWorksheet)
    
    'Rename new Worksheet
    newsheet.Name = NewSheetName
    createReport = NewSheetName
   
    'Reselects the HR Headcount report sheet for calculation to occur
    Sheets("sheet1").Select
   
End Function








