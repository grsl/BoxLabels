VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartUpForm 
   Caption         =   "Create Packing Labels"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7125
   OleObjectBlob   =   "StartUpForm.frx":0000
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
    Application.DisplayAlerts = False
    StartUpForm.Hide
    Application.DisplayAlerts = True
End Sub

Private Sub btnClear_Click()
    initialiseStartUpForm
End Sub

Private Sub CreateLabelData_Click()
    Application.DisplayAlerts = False
    createData
    ThisWorkbook.Save
    Application.DisplayAlerts = True
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

Private Sub txtProductCode_Change()

End Sub

Private Sub txtProductCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim result As Variant
    
    If (Me.txtProductCode.Text > "") Then
        productCode = UCase(Me.txtProductCode.Text)
        Cancel = False
        'Debug.Print "Product Code " & productCode & vbCrLf
    Else
        result = MsgBox("Please enter a product code.", vbInformation, "Product Code")
        Cancel = True
    End If
End Sub

Private Sub txtWorksOrder_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim result As Variant
    
    If (Me.txtWorksOrder.Text > "") Then
        worksOrder = UCase(Me.txtWorksOrder.Text)
        Cancel = False
        'Debug.Print "Works Order is " & worksOrder & vbCrLf
    Else
        result = MsgBox("Please enter a product code.", vbInformation, "Product Code")
        Cancel = True
    End If
End Sub

Private Sub UserForm_Activate()
    initialiseStartUpForm
End Sub

Private Sub UserForm_Terminate()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

Private Sub txtSerialStart_KeyPress(ByVal key As MSForms.ReturnInteger)
    If key < vbKey0 Or key > vbKey9 Then
        key = 0 ' this prevents the non-numeric data from showing up in the TextBox
        MsgBox "You can only enter numbers"
    End If
End Sub

