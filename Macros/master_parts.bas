Attribute VB_Name = "Module1"
'Basic Macro to create Master Parts list from a unit parts list
'Creator: Shane Burkhart

'Left are the check for whether or not there is a hand
'If there is a hand then Left and Right vars will be set appropriately
'However if there is no hand Left will be left as 0 to skip duplicates

'Constants
Public Const MAX_ROWS As Long = 65536
Public Const NOT_PROJECT_PROMPT As String = "You are not on a project page."
Public Const NOT_PROJECT_TITLE As String = "Not A Valid Page"

Public Const PROJECT_NAME_CELL As String = "B1"
Public Const PROJECT_PART_NUM_COLUMN = "B"
Public Const PROJECT_HAND_COLUMN = "D"
Public Const PROJECT_BUILDING_COLUMN = "F"
Public Const PROJECT_MEASURE_COLUMN = "I"
Public Const PROJECT_UNIT_COLUMN = "G"
Public Const PROJECT_MULTIPLYER_COLUMN = "H" 'Num of takeoff per unit
Public Const PROJECT_DATA_BEGIN = 6

Public Const MASTER_DATA_BEGIN As Integer = 5
Public Const MASTER_PROJECT_COLUMN As String = "A"
Public Const MASTER_PART_NUM_COLUMN As String = "C"
Public Const MASTER_HAND_COLUMN As String = "E"
Public Const MASTER_QUANTITY_COLUMN As String = "G"
Public Const MASTER_BUILDING_COLUMN As String = "J"
Public Const MASTER_FLOOR_COLUMN As String = "K"
Public Const MASTER_DIVISION_COLUMN As String = "B"
Public Const MASTER_MEASURE_COLUMN As String = "H"
Public Const MASTER_SHEET_NAME As String = "Master Parts List"

Public Const VALID_SHEET_NAME As String = "Validation Source Lists"
Public Const VALID_DATA_BEGIN As Integer = 5
Public Const VALID_PROJECT_COLUMN As String = "A"
Public Const VALID_DIVISION_COLUMN As String = "B"

Public Const UNIT_SORT_BEGIN As String = "6"
Public Const UNIT_BASEMENT_STD_COLUMN = "L"
Public Const UNIT_BASEMENT_REV_COLUMN = "M"
Public Const UNIT_FIRST_STD_COLUMN = "O"
Public Const UNIT_FIRST_REV_COLUMN = "P"
Public Const UNIT_SECOND_STD_COLUMN = "R"
Public Const UNIT_SECOND_REV_COLUMN = "S"
Public Const UNIT_THIRD_STD_COLUMN = "U"
Public Const UNIT_THIRD_REV_COLUMN = "V"
Public Const UNIT_FOURTH_STD_COLUMN = "X"
Public Const UNIT_FOURTH_REV_COLUMN = "Y"

    
Sub CreateMasterPartsList()
    'Variables
    Dim projectName As String
    Dim invalidMessage As Integer
    
    'Check for valid page
    projectName = ActiveSheet.range(PROJECT_NAME_CELL).Value
    If Not IsValidJob(projectName) Then
        invalidMessage = MsgBox(NOT_PROJECT_PROMPT, vbOKOnly, NOT_PROJECT_TITLE)
        Exit Sub
    End If
    
    'Sort Unit Parts List - Bldg, Part#, Hand
    ActiveSheet.range(UNIT_SORT_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=ActiveSheet.Columns(PROJECT_BUILDING_COLUMN), _
        Key2:=ActiveSheet.Columns(PROJECT_PART_NUM_COLUMN), _
        Key3:=ActiveSheet.Columns(PROJECT_HAND_COLUMN)

    'Sort Master By Job
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN)
    
    'Delete Job From Master
    DeleteJobFromMaster (projectName)
    
    'Sort Master By Job
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN)
    
    'Add Data to Master
    TransferDataFromUnitToMaster (projectName) 'Transfer Raw data.
    ConsolidateDuplicatesOnMaster 'Find the entries that need combining
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort _
    Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN), Order1:=xlDescending 'Sort to Remove space order
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN) 'Sort put back in order
    InsertDivisions
    'CheckUnitOfMeasure
    
    'Sort Unit by Unit then part number
    ActiveSheet.range(PROJECT_DATA_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=ActiveSheet.Columns(PROJECT_UNIT_COLUMN), _
        Key2:=ActiveSheet.Columns(PROJECT_PART_NUM_COLUMN)
End Sub

Function InsertDivisions()
    Dim masterRow As Integer: masterRow = MASTER_DATA_BEGIN
    Dim partNum As String: Dim j As Integer
    While Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = ""
        partNum = Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PART_NUM_COLUMN).Value
        Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_DIVISION_COLUMN).Value = GetDivision(partNum) 'Set division
        masterRow = masterRow + 1
    Wend
End Function

Function ConsolidateDuplicatesOnMaster()
    Dim masterRow As Integer: masterRow = MASTER_DATA_BEGIN
    Dim qty As Integer: Dim j As Integer
    While Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = ""
        qty = 0
        j = GetEndOfSameBelowMaster(masterRow)
        For i = masterRow To j Step 1
            qty = qty + Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_QUANTITY_COLUMN).Value 'Add up qty
        Next
        Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_QUANTITY_COLUMN).Value = qty 'Set first to correct quantity
        For i = masterRow + 1 To j Step 1
            Sheets(MASTER_SHEET_NAME).Cells(i, "A").EntireRow.Value = ""
        Next
        masterRow = j + 1
    Wend
End Function

Function TransferDataFromUnitToMaster(projectName As String)
    Dim bR As Integer: Dim bL As Integer: Dim f1R As Integer:
    Dim f1L As Integer: Dim f2R As Integer: Dim f2L As Integer:
    Dim f3R As Integer: Dim f3L As Integer: Dim f4R As Integer:
    Dim f4L As Integer: Dim j As Integer: Dim hand As String:
    Dim building As String: Dim partNum As String: Dim measure As String
    Dim i As Integer: Dim multi As Integer
    Dim masterRow As Integer: Dim projectRow As Integer
    masterRow = GetNextEmptyRow(MASTER_PROJECT_COLUMN, MASTER_DATA_BEGIN, MASTER_SHEET_NAME)
    projectRow = PROJECT_DATA_BEGIN
    While Not Cells(projectRow, PROJECT_PART_NUM_COLUMN).Value = ""
        bR = 0: bL = 0: f1R = 0: f1L = 0: f2R = 0: f2L = 0: f3R = 0: f3L = 0: f4R = 0: f4L = 0
        j = GetEndOfSameBelowUnit(projectRow)
        hand = GetHandUnit(j): building = GetBuildingUnit(j): partNum = GetPartNumUnit(j): measure = GetMeasureUnit(j)
        For i = projectRow To j Step 1
            multi = ActiveSheet.Cells(i, PROJECT_MULTIPLYER_COLUMN).Value
            bR = bR + GetBasementRight(hand, i) * multi
            bL = bL + GetBasementLeft(hand, i) * multi
            f1R = f1R + GetFirstRight(hand, i) * multi
            f1L = f1L + GetFirstLeft(hand, i) * multi
            f2R = f2R + GetSecondRight(hand, i) * multi
            f2L = f2L + GetSecondLeft(hand, i) * multi
            f3R = f3R + GetThirdRight(hand, i) * multi
            f3L = f3L + GetThirdLeft(hand, i) * multi
            f4R = f4R + GetFourthRight(hand, i) * multi
            f4L = f4L + GetFourthLeft(hand, i) * multi
        Next
        'Write Data
        'Basement
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, bR, building, "B", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, bL, building, "B", measure)
        'First
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f1R, building, "1", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f1L, building, "1", measure)
        'Second
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f2R, building, "2", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f2L, building, "2", measure)
        'Third
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f3R, building, "3", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f3L, building, "3", measure)
        'Fourth
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f4R, building, "4", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, hand, f4L, building, "4", measure)
        
        projectRow = j + 1
    Wend
End Function


Function WriteRowMaster(row As Integer, project As String, partNum As String, hand As String, qty As Integer, bldg As String, floor As String, measure As String)
    If qty = 0 Then
        WriteRowMaster = 0
        Exit Function
    End If
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_PROJECT_COLUMN).Value = project
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_PART_NUM_COLUMN).Value = partNum
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_HAND_COLUMN).Value = hand
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_QUANTITY_COLUMN).Value = qty
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_BUILDING_COLUMN).Value = bldg
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_FLOOR_COLUMN).Value = floor
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_MEASURE_COLUMN).Value = measure
    WriteRowMaster = 1
End Function

Function GetFourthLeft(hand As String, row As Integer)
    If hand = "R" Then
        GetFourthLeft = GetNum(row, UNIT_FOURTH_REV_COLUMN)
    ElseIf hand = "L" Then
        GetFourthLeft = GetNum(row, UNIT_FOURTH_STD_COLUMN)
    Else
        GetFourthLeft = 0
    End If
End Function

Function GetFourthRight(hand As String, row As Integer)
    If hand = "R" Then
        GetFourthRight = GetNum(row, UNIT_FOURTH_STD_COLUMN)
    ElseIf hand = "L" Then
        GetFourthRight = GetNum(row, UNIT_FOURTH_REV_COLUMN)
    Else
        GetFourthRight = GetNum(row, UNIT_FOURTH_STD_COLUMN) + GetNum(row, UNIT_FOURTH_REV_COLUMN)
    End If
End Function

Function GetThirdLeft(hand As String, row As Integer)
    If hand = "R" Then
        GetThirdLeft = GetNum(row, UNIT_THIRD_REV_COLUMN)
    ElseIf hand = "L" Then
        GetThirdLeft = GetNum(row, UNIT_THIRD_STD_COLUMN)
    Else
        GetThirdLeft = 0
    End If
End Function

Function GetThirdRight(hand As String, row As Integer)
    If hand = "R" Then
        GetThirdRight = GetNum(row, UNIT_THIRD_STD_COLUMN)
    ElseIf hand = "L" Then
        GetThirdRight = GetNum(row, UNIT_THIRD_REV_COLUMN)
    Else
        GetThirdRight = GetNum(row, UNIT_THIRD_STD_COLUMN) + GetNum(row, UNIT_THIRD_REV_COLUMN)
    End If
End Function

Function GetSecondLeft(hand As String, row As Integer)
    If hand = "R" Then
        GetSecondLeft = GetNum(row, UNIT_SECOND_REV_COLUMN)
    ElseIf hand = "L" Then
        GetSecondLeft = GetNum(row, UNIT_SECOND_STD_COLUMN)
    Else
        GetSecondLeft = 0
    End If
End Function

Function GetSecondRight(hand As String, row As Integer)
    If hand = "R" Then
        GetSecondRight = GetNum(row, UNIT_SECOND_STD_COLUMN)
    ElseIf hand = "L" Then
        GetSecondRight = GetNum(row, UNIT_SECOND_REV_COLUMN)
    Else
        GetSecondRight = GetNum(row, UNIT_SECOND_STD_COLUMN) + GetNum(row, UNIT_SECOND_REV_COLUMN)
    End If
End Function

Function GetFirstLeft(hand As String, row As Integer)
    If hand = "R" Then
        GetFirstLeft = GetNum(row, UNIT_FIRST_REV_COLUMN)
    ElseIf hand = "L" Then
        GetFirstLeft = GetNum(row, UNIT_FIRST_STD_COLUMN)
    Else
        GetFirstLeft = 0
    End If
End Function

Function GetFirstRight(hand As String, row As Integer)
    If hand = "R" Then
        GetFirstRight = GetNum(row, UNIT_FIRST_STD_COLUMN)
    ElseIf hand = "L" Then
        GetFirstRight = GetNum(row, UNIT_FIRST_REV_COLUMN)
    Else
        GetFirstRight = GetNum(row, UNIT_FIRST_STD_COLUMN) + GetNum(row, UNIT_FIRST_REV_COLUMN)
    End If
End Function

Function GetBasementLeft(hand As String, row As Integer)
    If hand = "R" Then
        GetBasementLeft = GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    ElseIf hand = "L" Then
        GetBasementLeft = GetNum(row, UNIT_BASEMENT_STD_COLUMN)
    Else
        GetBasementLeft = 0
    End If
End Function

Function GetBasementRight(hand As String, row As Integer)
    If hand = "R" Then
        GetBasementRight = GetNum(row, UNIT_BASEMENT_STD_COLUMN)
    ElseIf hand = "L" Then
        GetBasementRight = GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    Else
        GetBasementRight = GetNum(row, UNIT_BASEMENT_STD_COLUMN) + GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    End If
End Function

Function GetNum(row As Integer, col As String)
    If Cells(row, col).Value = "" Then
        GetNum = 0
    Else
        GetNum = Cells(row, col).Value
    End If
End Function
Function GetHandUnit(row As Integer)
    If ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value = "L" Or ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value = "R" Then
        GetHandUnit = ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value
    Else
        GetHandUnit = ""
    End If
End Function

Function GetPartNumUnit(row As Integer)
    GetPartNumUnit = ActiveSheet.Cells(row, PROJECT_PART_NUM_COLUMN).Value
End Function

Function GetBuildingUnit(row As Integer)
    GetBuildingUnit = ActiveSheet.Cells(row, PROJECT_BUILDING_COLUMN).Value
End Function

Function GetMeasureUnit(row As Integer)
    GetMeasureUnit = ActiveSheet.Cells(row, PROJECT_MEASURE_COLUMN).Value
End Function

Function GetNextEmptyRow(col As String, row As Integer, sheet As String)
    While Not Sheets(sheet).range(col & row).Value = ""
        row = row + 1
    Wend
    GetNextEmptyRow = row
End Function

Function DeleteJobFromMaster(projectName As String)
    For i = MAX_ROWS To MASTER_DATA_BEGIN Step -1
        If Sheets(MASTER_SHEET_NAME).Cells(i, "A").Value = projectName Then
            Sheets(MASTER_SHEET_NAME).Cells(i, "A").EntireRow.Delete
        End If
    Next
End Function

Function IsValidJob(projectName As String)
    If projectName = "" Then
        IsValidJob = False
        Exit Function
    End If
    For Each c In Sheets(VALID_SHEET_NAME).range(GetValidRangeDown(VALID_DATA_BEGIN, VALID_PROJECT_COLUMN, VALID_SHEET_NAME)).Cells
        If projectName = c.Value Then
            IsValidJob = True
            Exit Function
        End If
    Next
    IsValidJob = False
End Function

Function GetDivision(partNum As String)
    For Each c In Sheets(VALID_SHEET_NAME).range(GetValidRangeDown(VALID_DATA_BEGIN, VALID_DIVISION_COLUMN, VALID_SHEET_NAME)).Cells
        If Left(partNum, 2) = Left(c.Value, 2) Then
            GetDivision = c.Value
            Exit Function
        End If
    Next
    GetDivision = "No Division"
End Function

Function GetEndOfSameBelowUnit(rowNum As Integer)
    Dim i As Integer: Dim partN As String
    Dim hand As String: Dim bldg As String
    partN = ActiveSheet.Cells(rowNum, PROJECT_PART_NUM_COLUMN).Value
    hand = ActiveSheet.Cells(rowNum, PROJECT_HAND_COLUMN).Value
    bldg = ActiveSheet.Cells(rowNum, PROJECT_BUILDING_COLUMN).Value
    i = rowNum
    While ActiveSheet.Cells(i, PROJECT_PART_NUM_COLUMN).Value = partN And _
        ActiveSheet.Cells(i, PROJECT_HAND_COLUMN).Value = hand And ActiveSheet.Cells(i, PROJECT_BUILDING_COLUMN).Value = bldg
        i = i + 1
    Wend
    GetEndOfSameBelowUnit = (i - 1)
End Function

Function GetEndOfSameBelowMaster(rowNum As Integer)
    Dim i As Integer: Dim partN As String: Dim floor As String
    Dim hand As String: Dim bldg As String: Dim projectName As String
    partN = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_PART_NUM_COLUMN).Value
    hand = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_HAND_COLUMN).Value
    bldg = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_BUILDING_COLUMN).Value
    projectName = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_PROJECT_COLUMN).Value
    floor = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_FLOOR_COLUMN).Value
    i = rowNum
    While Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PART_NUM_COLUMN).Value = partN And _
        Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_HAND_COLUMN).Value = hand And Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_BUILDING_COLUMN).Value = bldg And _
        Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PROJECT_COLUMN).Value = projectName And Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_FLOOR_COLUMN).Value = floor
        i = i + 1
    Wend
    GetEndOfSameBelowMaster = (i - 1)
End Function

Function GetValidRangeDown(row As Integer, col As String, sheet As String)
    Dim start As String
    start = col & row
    While Not Sheets(sheet).range(col & row).Value = ""
        row = row + 1
    Wend
    GetValidRangeDown = start & ":" & col & (row - 1)
End Function


