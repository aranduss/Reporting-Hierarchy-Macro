Attribute VB_Name = "Module1"

Sub Manager_Hierarchy_Report()
Attribute Manager_Hierarchy_Report.VB_Description = "Creates a new tab which identifies all managers in the reporting structure starting from the employee you enter."
Attribute Manager_Hierarchy_Report.VB_ProcData.VB_Invoke_Func = "m\n14"
'
' Manager_Hierarchy_Report Macro
' Creates a new tab which identifies all managers in the reporting structure starting from the employee you enter.
'
' Keyboard Shortcut: Ctrl+m
'
    If checkForHeadcount = True Then

        'Select the sheet that has the Headcount Report Info. Note that this should be named 'sheet1' for the script to work
        Sheets("Sheet1").Select
        Rows("1:1").Select
        
        'Delete all rows until the Title bars are on Row 1. This is acomplished by looking for the row containing the 'Empl ID' field
        'Call formatReport
    
        'Creat Msg Box
        input_text = Application.InputBox(prompt:="Please enter the Employee ID or Email Address", Type:=2)
            
        If (input_text = "False") Then
            Exit Sub
        Else
            Dim search_text As String
        
            'If input is email, convert to ID for lookup
            If InStr(1, input_text, "@") > 0 Then
                search_text = ConvertEmail(input_text)
            Else
                search_text = input_text
            End If
            
            'If search_text = "" then the "Cancel" button was clicked in the 'Incorrect Input' Input Box so the Macro is discontinued
            If (search_text <> "") Then
                Call findEmployee(search_text)
            End If
        End If
    Else
        MsgBox ("Please import the most current Headcount report with the Employee ID, Name, Title, and Supervisor ID fields as a new tab in this workbook and name it 'sheet1'.")
        Exit Sub
    End If
    
End Sub

Function checkForHeadcount()
'Check to see if sheet1 exists
    If (SheetExists("sheet1")) Then
        checkForHeadcount = True
    Else
        checkForHeadcount = False
    End If
End Function

Sub formatReport()
    'Delete all rows until the Title bars are on Row 1. This is acomplished by looking for the row containing the 'Empl ID' field
    Do While (ActiveSheet.Range(findCol("Empl ID") & "1").Value <> "Empl ID")
        Rows(1).EntireRow.Delete
    Loop
   
End Sub

Sub findEmployee(search_text)
    'Find the employee that relates to the entered search criteria
    Column = findCol("Empl ID")
    Emp_ID_Path = "$" & Column & "$1:$" & Column & Row_Length()
    Set Search = ActiveSheet.Range(Emp_ID_Path)
    Set searchlast = Search.Cells(Search.Cells.Count)
    
    'Assign found cell to a variable
    Set foundValue = ActiveSheet.Range(Emp_ID_Path).Find(search_text, searchlast, xlValues, LookAt:=xlWhole)
    
    If foundValue Is Nothing Then
        MsgBox (search_text & " was unable to be found. Please ensure that you entered the correct email or Employee ID. Remember to only include the digits in the employee ID.")
        Call Manager_Hierarchy_Report
    Else
          'Find Supervisor Info
        SupID = emp_Supervisor_ID(foundValue)
        supName = emp_Name(SupID)
               
        'Create a reporting hierarchy in supList collection starting from entered employee and ending with the CEO
        Dim supList As Collection
        Set supList = New Collection
        reachedCEO = "False"
        Counter = 0
        Do While reachedCEO = "False" And Counter < 50
        
            Dim Emp As cSupervisor
            Set Emp = New cSupervisor
            Emp.ID = foundValue
            Emp.Name = emp_Name(Emp.ID)
            Emp.Title = emp_Title(Emp.ID)
            Emp.Supervisor_ID = emp_Supervisor_ID(Emp.ID)
            supList.Add Emp
            foundValue = Emp.Supervisor_ID
                    
            lastEmpID = Emp.ID
            nextEmpID = Emp.Supervisor_ID
            
            Counter = Counter + 1
            
            'The CEO in this example has no Supervisor ID so this is the trigger to stop searching
            If (Emp.Supervisor_ID = "") Then
                reachedCEO = True
            End If
        Loop
            
        'Create Report sheet
        Call create_Report(supList, supList(1).ID)
    End If
    
End Sub

Function ConvertEmail(input_text)
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
        ConvertEmail = search_text
        Exit Function
        
InvalidEmail:
        Dim msg As String
        msg = input_text & " not found. Please ensure that you entered the correct email or Employee ID. Remember to only include the digits in the employee ID."
        
        input_text = Application.InputBox(prompt:=msg, Type:=2)
                
        If (input_text = "False") Then
            Exit Function
        Else
            If InStr(1, input_text, "@") > 0 Then
                search_text = ConvertEmail(input_text)
            Else
                search_text = input_text
            End If
        
            ConvertEmail = search_text
                    
        End If
        
End Function

Function Row_Length() As Integer
    'This function finds the last row in use. This is used mostly for the Range criteria when working with Worksheet commands (i.e. ActiveSheet.Range("$A$1:$A" & Row_Length()).
    Row_Length = Sheets("sheet1").Range("A" & Rows.Count).End(xlUp).Row
End Function


Function findCol(CellText)
    'function to find the column with a header containing the passed info
    ThisWorkbook.Sheets("Sheet1").Select
    Set findCol = ActiveSheet.Range("A1:Z" & Row_Length()).Find(What:=CellText, LookIn:=xlValues, LookAt:=xlWhole)
    columnLetter = Split(Cells(1, findCol.Column).Address, "$")(1)
    findCol = columnLetter
End Function

Function emp_Name(emp_Id)
    
    'Create dynamic Range for whatever column the "Empl ID' field is in.
    Emp_ID_Path = "$" & findCol("Empl ID") & "$1:$" & findCol("Empl ID") & Row_Length()

    'Find Employee ID field for the employee searched for so that the row value of the cell can be used to find their Name from the 'Name' column
     ThisWorkbook.Sheets("sheet1").Select
    Rows("1:1").Select
    Selection.AutoFilter
    Set Search = ActiveSheet.Range(Emp_ID_Path)
    Set searchlast = Search.Cells(Search.Cells.Count)
    Set foundValue = ActiveSheet.Range(Emp_ID_Path).Find(emp_Id, searchlast, xlValues)
    
    'Find the column that contains the 'Name' Field
    Column = findCol("Name")
    
    'Use the Cell containing the employee's ID to find record's row and the identified name column to find the employee's name
    emp_nameCell = Column & foundValue.Row()
    
    'Assign the employee's name to a variable
    emp_Name = Sheets("sheet1").Range(emp_nameCell).Value
End Function

Function emp_Supervisor_ID(emp_Id)
    
    'Create dynamic Range for whatever column the "Empl ID' field is in.
    Emp_ID_Path = "$" & findCol("Empl ID") & "$1:$" & findCol("Empl ID") & Row_Length()

    'Find Employee ID field for the employee searched for so that the row value of the cell can be used to find their Supervisor's ID from the 'Name' column
     ThisWorkbook.Sheets("sheet1").Select
    Rows("1:1").Select
    Selection.AutoFilter
    Set Search = Sheets("sheet1").Range(Emp_ID_Path)
    Set searchlast = Search.Cells(Search.Cells.Count)
    Set foundValue = Sheets("sheet1").Range(Emp_ID_Path).Find(emp_Id, searchlast, xlValues)
    
    'Find the column that contains the 'Suvpervisor ID' Field
    Column = findCol("Supv ID")
    
    'Use the Cell containing the employee's ID to find record's row and the identified 'Supervisor ID' column to find the employee's Supervisor ID
    emp_nameCell = Column & foundValue.Row()
    emp_Supervisor_ID = Sheets("sheet1").Range(emp_nameCell).Value
End Function

Function emp_Title(emp_Id)

    'Create dynamic Range for whatever column the "Empl ID' field is in.
    Emp_ID_Path = "$" & findCol("Empl ID") & "$1:$" & findCol("Empl ID") & Row_Length()

    'Find Employee ID field for the employee searched for so that the row value of the cell can be used to find their Title from the 'Title' column
     ThisWorkbook.Sheets("Sheet1").Select
    Rows("1:1").Select
    Selection.AutoFilter
    Set Search = Sheets("sheet1").Range(Emp_ID_Path)
    Set searchlast = Search.Cells(Search.Cells.Count)
    Set foundValue = Sheets("sheet1").Range(Emp_ID_Path).Find(emp_Id, searchlast, xlValues)
    
    'Find the column that contains the 'Job Title' Field
    Column = findCol("Job Title")
    
    'Use the Cell containing the employee's ID to find record's row and the identified 'Title' column to find the employee's Title
    emp_nameCell = Column & foundValue.Row()
    emp_Title = Sheets("sheet1").Range(emp_nameCell).Value
End Function

Function SheetExists(shtName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function

Sub create_Report(supList As Collection, ID As String)
    'Set new sheet's name as a variable
    Dim NewSheetName As String
    NewSheetName = (ID & "_Report")

    'Delete any existing worksheets with the same name as the one about to be created
    If (SheetExists(NewSheetName)) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(ID & "_Report").Delete
        Application.DisplayAlerts = True
    End If

    'Create new worksheet
    Dim newsheet
    Set newsheet = Sheets.Add(After:=Sheets(Worksheets.Count), Count:=1, Type:=xlWorksheet)
    
    'Rename new Worksheet
    newsheet.Name = NewSheetName
    newsheet.Select
    
    'Set format of report columns to Text to avoid the 0's in ID's being cut off
    Range("A1:D" & supList.Count + 1).NumberFormat = "@"
        
    'Set Title Row Values in report sheet
    ActiveSheet.Range("A1").Value = "Employee ID"
    ActiveSheet.Range("B1").Value = "Name"
    ActiveSheet.Range("C1").Value = "Title"
    ActiveSheet.Range("D1").Value = "Supervisor ID"
            
    'Print all supervisors stored in supList in order
    x = 1
    Do While x < supList.Count + 1
        Set Emp = supList(x)
        ActiveSheet.Range("A" & x + 1).Value = Emp.ID()
        ActiveSheet.Range("B" & x + 1).Value = Emp.Name()
        ActiveSheet.Range("C" & x + 1).Value = Emp.Title()
        ActiveSheet.Range("D" & x + 1).Value = Emp.Supervisor_ID()
        
        x = x + 1
    Loop
    
    'Autofit column width for ease of reading
     ThisWorkbook.ActiveSheet.Range("A1:D" & supList.Count + 1).EntireColumn.AutoFit
    
    'Create a summary box indicating how many reporting supervisors exist in reporting structure
    SummaryBox = MsgBox(supList(1).Name & " has " & supList.Count & " employees in their reporting structure", vbOKOnly, "Summary")
        
End Sub



