Attribute VB_Name = "modMyWizard"
Option Explicit

Public Const gintStepMax = "5"  'Don't forget to add new frames to your Form, if you need more Steps
Public gintStep As Integer      'We are at this Step
Public gintStepCheck As Integer 'To start the timer-Functions

Public gSettingsName(0 To 2) As String


Public Function StartUpWizard()
    Dim i As Integer
    
    ' Do some Stuff in here, fill the wizard with Data or some information.
    
    gSettingsName(0) = "Auto-Installation"
    gSettingsName(1) = "Userdefined Installation"
    gSettingsName(2) = "Installation for Advanced Users"
    
    For i = 0 To UBound(gSettingsName)
        frmWizard.cmbInstallationType.AddItem gSettingsName(i)
    Next
    frmWizard.cmbInstallationType.ListIndex = 0
    
    
    
End Function

