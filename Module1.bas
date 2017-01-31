Attribute VB_Name = "Module1"
Option Explicit

Public Sub createWeekData()
    Dim x As Integer
    'Add data to the Week Number combobox.
    If (StartUpForm.lstWeekNumber.ListCount < 1) Then
        For x = 1 To 52
            StartUpForm.lstWeekNumber.AddItem ("Week " & x)
        Next x
    End If
End Sub

Public Sub createData()
    Dim xlOffset As Integer
    Dim woOffset As Integer
    Dim weekNumber As String
    Dim remainder As Integer
    Dim numberOfBoxes As Integer
    Dim firstSerialNumber As Integer
    Dim lastSerialNumber As Integer
    Dim serialStart As Integer
    Dim x As Integer
    
    Dim productCode As String
    Dim pumpsOrdered As Integer
    Dim pumpsPerBox As Integer
    Dim worksOrder As String
    Dim worksOrderNumber As String
    
    Dim thisYear As Variant
    Dim result As Variant
    
    Application.DisplayAlerts = False
    
    ' Turn off updating to speed up the process.
    Application.ScreenUpdating = False
    
    'Clear all cells
    Worksheets("LabelData").Cells.Clear
    
    'The number of characters at the start of the Works Order Number that are not numeric plus one.
    woOffset = 3
    
    'Start after the header line in the spreadsheet.
    xlOffset = 1
    
    With StartUpForm
        If (.txtProductCode.Text > "") Then
            productCode = UCase(.txtProductCode.Text)
            'Debug.Print "Product Code " & productCode & vbCrLf
        Else
            result = MsgBox("Please enter a product code.", vbInformation, "Product Code")
        End If
    
        If (.txtWorksOrder.Text > "") Then
            worksOrder = UCase(.txtWorksOrder.Text)
            'Debug.Print "Works Order " & worksOrder & vbCrLf
            Debug.Print "Works Order " & result
        Else
            result = MsgBox("Please enter a Works Order number.", vbInformation, "Product Code")
        End If
    
        weekNumber = Int(Mid(.lstWeekNumber.Value, 6, 2))
        'Debug.Print weekNumber
        'Debug.Print "The week number is " & test
        'Debug.Print "The week number is " & weekNumber
    
        pumpsOrdered = Int(.numberOfPumps.Value)
        If (pumpsOrdered = 0) Then
            result = MsgBox("Please enter a the number of pumps in the order.", vbInformation, "Number of Pumps")
        End If
        'Debug.Print "Pumps Ordered = " & Int(pumpsOrdered)
    
        pumpsPerBox = Int(.numberOfPumpsPerBox.Value)
        If (pumpsPerBox = 0) Then
            result = MsgBox("Please enter the number of pumps a box will hold.", vbInformation, "Box Capacity")
        End If
        'Debug.Print "Pumps per box = " & pumpsPerBox
        
                
        ' Andy needs to be able to change where the serial numbers start.
        serialStart = Int(.txtSerialStart.Value)
                
        If (serialStart > 0) Then
            firstSerialNumber = serialStart
            lastSerialNumber = Int(pumpsPerBox) + serialStart - 1
            Debug.Print "Serial start = " & serialStart
        Else
            firstSerialNumber = 1
            lastSerialNumber = Int(pumpsPerBox)
            Debug.Print "Serial start = " & serialStart
        End If
        
    End With
    
    worksOrderNumber = Mid(worksOrder, woOffset)
    'Debug.Print "Works Order Number = " & worksOrderNumber & vbCrLf
    
    thisYear = format(Date, "YY")
    'Debug.Print "This year is " & thisYear & vbCrLf
    
    If pumpsOrdered > 0 And pumpsPerBox > 0 Then
        numberOfBoxes = Int(pumpsOrdered / pumpsPerBox)
        remainder = pumpsOrdered Mod pumpsPerBox
        'Debug.Print "Number of boxes is " & numberOfBoxes & " the remainder is " & remainder & "."
    End If

    If remainder > 0 Then
        numberOfBoxes = numberOfBoxes + 1
    End If

    'Write the headers to the worksheet.
    With Worksheets("LabelData")
        Range("A1").Value = "Product Code"
        Range("B1").Value = "Works Order No."
        Range("C1").Value = "First Serial Number in the Box"
        Range("D1").Value = "Last Serial Number in the Box"
        Range("E1").Value = "Number of Pumps in the Box"
        Range("F1").Value = "Box X of Y"
    End With

   


    For x = 1 To numberOfBoxes Step 1
    ' Information required on the labels: Product Code PRODUCT; Works Order Number WO12345;
    ' Serial Number From: 16130051 12345 to: 16130100 12345. This format is for RD1, APC and HiP.
    ' The format for SSCOR pumps is WO12345-0001 i.e. works order no. hyphen and pump number.
    ' Number of Pumps in the Box  50
    ' Box 2 of 30
    Cells(x + xlOffset, 1).Value = productCode
    Cells(x + xlOffset, 2).Value = worksOrder
    Cells(x + xlOffset, 6).Value = "Box " & x & " of " & numberOfBoxes
        
        If ((x = numberOfBoxes) And (remainder > 0)) Then
            lastSerialNumber = Int(lastSerialNumber - pumpsPerBox + remainder)
            With Worksheets("LabelData")
                Cells(x + xlOffset, 3).Value = format(currentYear, "00") & format(weekNumber, "00") & format(firstSerialNumber, "0000") & " " & worksOrderNumber
                Cells(x + xlOffset, 4).Value = format(currentYear, "00") & format(weekNumber, "00") & format(lastSerialNumber, "0000") & " " & worksOrderNumber
                Cells(x + xlOffset, 5).Value = remainder
                Cells(x + xlOffset, 6).Value = "Box " & x & " of " & numberOfBoxes
            End With
        Else
            With Worksheets("LabelData")
                Cells(x + xlOffset, 3).Value = format(currentYear, "00") & format(weekNumber, "00") & format(firstSerialNumber, "0000") & " " & worksOrderNumber
                Cells(x + xlOffset, 4).Value = format(currentYear, "00") & format(weekNumber, "00") & format(lastSerialNumber, "0000") & " " & worksOrderNumber
                Cells(x + xlOffset, 5).Value = pumpsPerBox
            End With
            firstSerialNumber = firstSerialNumber + pumpsPerBox
            lastSerialNumber = Int(lastSerialNumber + pumpsPerBox)
        End If
    Next x
   
    Worksheets("LabelData").Cells(1, 1).Activate
    
    'Save the data in the existing document.
    ThisWorkbook.Save

    ' Turn updating back on to display the data.
    Application.ScreenUpdating = True
    
    ' Turn updating back on to display the data.
    'Application.DisplayAlerts = True
End Sub

Public Function currentWeek()
    'List box starts at 0 so is off by one.
    currentWeek = format(Date, "ww")
End Function

Public Function currentYear()
    'List box starts at 0 so is off by one.
    currentYear = format(Date, "YY")
End Function

Public Function initialiseStartUpForm()

    Application.DisplayAlerts = False
    
    'Clear all cells
    Worksheets("LabelData").Cells.Clear
    
    createWeekData
    
    With StartUpForm
        ' Set the defaults on the form.
        .lstWeekNumber.Selected(currentWeek()) = True
        .txtProductCode.Text = ""
        .txtWorksOrder.Text = ""
        .numberOfPumps.Value = 0
        .numberOfPumpsPerBox = 0
        .txtSerialStart.Text = "0"
        .txtProductCode.SetFocus
    End With


    Application.DisplayAlerts = True

End Function

Public Function selectedWeek() As Integer
    Dim i As Integer
    
    Dim iCount As Integer
    iCount = StartUpForm.lstWeekNumber.ListCount - 1
    Debug.Print iCount
    
    For i = 0 To i = iCount
        If (StartUpForm.lstWeekNumber.Selected(i) = True) Then
            selectedWeek = i
            Debug.Print "selected week is " & selectedWeek & " and i is " & i
        End If
    Next i
        
    'selectedWeek = currentWeek()
    
End Function
