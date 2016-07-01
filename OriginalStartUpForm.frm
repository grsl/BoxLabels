VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartUpForm 
   Caption         =   "Create Packing Labels"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11070
   OleObjectBlob   =   "OriginalStartUpForm.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "StartUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    StartUpForm.Hide
End Sub

Private Sub btnSave_Click()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

Private Sub btnClear_Click()
    'Clear all cells
    Worksheets("LabelData").Cells.Clear
    ' Set the text boxes to zero.
        
    Dim cntrl As MSForms.Control
        For Each cntrl In StartUpForm.Controls
            If TypeOf cntrl Is TextBox Then
                cntrl.tabkeybehaviour = False
            End If
        Next cntrl
    
    With StartUpForm
        ' Set the default item in the three list boxes to the first one.
        .lstWorksOrderNumber.ListIndex = 0
        .lstWorksOrderNumber.Selected(0) = True
        
        ' Initialise the text boxes to zero.
        .numberOfPumps.Value = 0
        .numberOfPumpsPerBox = 0
        
    End With
End Sub

Private Sub btnPrintLabels_Click()
'    Dim wordApp As Word.Application
'    Dim wordDoc As Word.Document
'    Dim dataSource As MailMergeDataSource
    
'    Set wordApp = CreateObject("Word.Application")
'    Set wordDoc = wordApp.Documents.Open(ActiveWorkbook.Path & "\subAssemblyLabels.docm")
    
'    wordDoc.Activate
    'With wordDoc.MailMerge
    '    .MainDocumentType = wdMailingLabels
        '.OpenDataSource Name:=ActiveDocument.Path & "\Labels.docm", _
        '                ReadOnly:=True, _
        '                Connection:="LabelData"
    'End With
                                                                
    'Application.Wait (Now + TimeValue("0:00:10"))
    
    'wordDoc.Close   ' Close the document
    'wordApp.Quit    ' Close the Word application
    
'    Set wordDoc = Nothing
'    Set wordApp = Nothing
End Sub

Private Sub CreateLabelData_Click()
    Call createData
End Sub

Private Sub lstWorksOrderNumber_Click()
    Dim Selected As Boolean
    Selected = False
    
    For x = 0 To lstWorksOrderNumber.ListCount - 1
        If lstWorksOrderNumber.Selected(x) = True Then
            Selected = True
            bWorksOrder = True
        End If
    Next x
    
    If Not Selected Then
        StartUpForm.lstWorksOrderNumber.SetFocus
        MsgBox "Please select a works order number."
    End If
End Sub

Private Sub numberOfPumps_KeyPress(ByVal key As MSForms.ReturnInteger)
    If key < vbKey0 Or key > vbKey9 Then
        key = 0 ' this prevents the non-numeric data from showing up in the TextBox
        MsgBox "You can only enter numbers"
    End If
End Sub

Private Sub numberOfPumpsPerBox_KeyPress(ByVal key As MSForms.ReturnInteger)
    If key < vbKey0 Or key > vbKey9 Then
        key = 0 ' this prevents the non-numeric data from showing up in the TextBox
        MsgBox "You can only enter numbers"
    End If
End Sub

Private Sub UserForm_Activate()
    Dim arr, tmp, arFields As Variant
    Dim x As Integer
    Dim str As String
    Dim dbConnection As ADODB.Connection
    Dim dbRecordSet As ADODB.recordSet
    Dim cntrl As MSForms.Control
    
    For Each cntrl In StartUpForm.Controls
        If TypeOf cntrl Is TextBox Then
            cntrl.tabkeybehaviour = False
        End If
    Next cntrl
   
    Dim bWorksOrder As Boolean
    Dim bNumberOfPumps As Boolean
    Dim bNumberOfBoxes As Boolean
    
    'Clear all cells
    Worksheets("LabelData").Cells.Clear

    'Connect to the database.
    Set dbConnection = New ADODB.Connection
    dbConnection.ConnectionString = "driver={SQL Server};server=CAP-APPS64;uid=sa;pwd=CharlesA1;database=CAP-Test"
    'dbConnection.ConnectionString = "driver={SQL Server};server=CAP-APPS64;uid=sa;pwd=CharlesA1;database=CAP-Live"
    dbConnection.Open
    
    x = 0
    
    Set dbRecordSet = dbConnection.Execute("Select * from vw_subAssemblyLabels")
    ' Find out if the connection is valid.
    If dbConnection.State = adStateOpen Then
        
        dbRecordSet.MoveFirst
        
        While Not dbRecordSet.EOF
            arr = dbRecordSet.GetRows(1)
            If Not dbRecordSet.EOF Then
                dbRecordSet.MoveNext
            End If
            
            With StartUpForm.lstWorksOrderNumber
                .ColumnCount = 3
                .ColumnWidths = "55;55;400"
                .AddItem
                .List(x, 0) = arr(0, 0)
                .List(x, 1) = arr(1, 0)
                .List(x, 2) = arr(2, 0)
            End With
           
            x = x + 1
        Wend
        
    Else
        MsgBox "Can't connect to the database." & vbCr & vbCr & "Please contact your system administrator."
    End If
    
    dbRecordSet.Close
    StartUpForm.lstWorksOrderNumber.ListIndex = 0
    StartUpForm.lstWorksOrderNumber.Selected(0) = True
    
    ' Initialise the text boxes to zero.
    StartUpForm.numberOfPumps.Value = 0
    StartUpForm.numberOfPumpsPerBox = 0
    
End Sub
