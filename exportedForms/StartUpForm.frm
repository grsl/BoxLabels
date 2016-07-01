VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartUpForm 
   Caption         =   "Create Packing Labels"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   -14805
   ClientWidth     =   9855
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
    StartUpForm.Hide
End Sub

Private Sub cmbProductCode_Enter()
    StartUpForm.lblFeedback = vbCr & "Please select the correct Product Code from the drop down list."
End Sub

Private Sub cmbWeekNumber_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
End Sub

Private Sub cmbWorksOrderNumber_Enter()
    StartUpForm.lblFeedback = vbCr & "Please select the correct Works Order Number from the drop down list."
End Sub
Private Sub cmbWeekNumber_Enter()
    StartUpForm.lblFeedback = vbCr & "Please select the correct Week Number from the drop down list."
End Sub

Private Sub lblWeekNumber_Click()

End Sub

Private Sub numberOfPumps_Enter()
    StartUpForm.lblFeedback = vbCr & "Please enter the number of pumps that have been ordered." & vbCr & vbCr & "Only whole numbers are acceptable."
End Sub

Private Sub numberOfPumps_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(numberOfPumps.Value) Then
        Cancel = False
    Else
        Cancel = True
        MsgBox "Please enter a number.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub numberOfPumps_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        Cancel = True
        MsgBox vbCr & "Only numbers 0 to 9 can be emtered in this field." & vbCr & vbCr & "Please enter a number.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub numberOfPumpsPerBox_Enter()
    StartUpForm.lblFeedback = "Please enter the maximum number of pumps that a box will hold." & vbCr & vbCr & "Only whole numbers are acceptable."
End Sub

Private Sub numberOfPumpsPerBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = False
    StartUpForm.btnCreateLabelData.Enabled = True
    'MsgBox " Leaving Numbber of Pumps", vbOKOnly
End Sub

Private Sub numberOfPumpsPerBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        Cancel = True
        MsgBox vbCr & "Only numbers 0 to 9 can be emtered in this field." & vbCr & vbCr & "Please enter a number.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub productDetails_Click()

End Sub

Private Sub title_Click()

End Sub

Private Sub txbProductCodeSuffix_Enter()
    StartUpForm.lblFeedback = "There is no error checking on this, so whatever you enter will be appended to the product code. If you want a space between the product code and the suffix you will have to add it."
End Sub

Private Sub CreateLabelData_Click()
    MsgBox "Data Created"
End Sub

Private Sub UserForm_Activate()
    'StartUpForm.cmbProductCode.SetFocus
    StartUpForm.btnCreateLabelData.Enabled = False
End Sub

Private Sub UserForm_Initialize()
'    StartUpForm.cmbProductCode.SetFocus
'    Me.cmbWeekNumber.List = Worksheets("labeldata").Range("AB1:AB52").Value
'    Me.cmbWorksOrderNumber.List = Worksheets("labeldata").Range("AC1:AC70").Value
'    Me.cmbProductCode.List = Worksheets("labeldata").Range("AD1:AD2761").Value
End Sub

Private Sub UserForm_Layout()
    cmbProductCode.TabStop = True
    cmbProductCode.TabIndex = 0
    cmbWorksOrderNumber.TabStop = True
    cmbWorksOrderNumber.TabIndex = 1
    cmbWeekNumber.TabStop = True
    cmbWeekNumber.TabIndex = 2
    numberOfPumps.TabStop = True
    numberOfPumps.TabIndex = 3
    numberOfPumpsPerBox.TabStop = True
    numberOfPumpsPerBox.TabIndex = 4
    txbProductCodeSuffix.TabStop = True
    txbProductCodeSuffix.TabIndex = 5
    txbSerialNumberSuffix.TabStop = True
    txbSerialNumberSuffix.TabIndex = 6
    chkSscor.TabStop = True
    chkSscor.TabIndex = 7
    btnCancel.TabStop = True
    btnCancel.TabIndex = 8
    btnCreateLabelData.TabStop = True
    btnCreateLabelData.TabIndex = 9
End Sub
