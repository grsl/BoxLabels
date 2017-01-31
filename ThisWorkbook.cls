VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Workbook_Open()
    
    'Clear all cells
    Worksheets("LabelData").Cells.Clear

    'Create the data for the Week drop down list
    Call createWeekData
    
    Worksheets("LabelData").Cells(1, 1).Activate
    
    StartUpForm.Show
End Sub

