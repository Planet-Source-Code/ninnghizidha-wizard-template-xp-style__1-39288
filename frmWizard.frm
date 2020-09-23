VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Installation Wizard"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Timer tmrStep 
      Interval        =   1
      Left            =   720
      Top             =   4920
   End
   Begin MyWizard.Xp_ProgressBar barSteps 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4200
      Width           =   5055
      _ExtentX        =   13785
      _ExtentY        =   450
      ProgressLook    =   1
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.PictureBox PictureBox 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   0
      Picture         =   "frmWizard.frx":0000
      ScaleHeight     =   4515
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   8124.031
      TabIndex        =   0
      Top             =   0
      Width           =   7920
      Begin VB.Frame frmStep 
         Caption         =   "Step 5"
         Height          =   4335
         Index           =   5
         Left            =   -2040
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtSummary 
            BackColor       =   &H8000000F&
            Height          =   2055
            Left            =   600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   23
            Text            =   "frmWizard.frx":47E7
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "And finally .."
            Height          =   615
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame frmStep 
         Caption         =   "Step 4"
         Height          =   4095
         Index           =   4
         Left            =   -2280
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton cmdBoring 
            Caption         =   "BORING!"
            Height          =   735
            Left            =   720
            TabIndex        =   28
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   ".. some more Information, followed by even more controls .."
            Height          =   495
            Left            =   360
            TabIndex        =   25
            Top             =   960
            Width           =   4335
         End
      End
      Begin VB.Frame frmStep 
         Caption         =   "Step 3"
         Height          =   4215
         Index           =   3
         Left            =   -2400
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton cmdFolderSQLPlus 
            Caption         =   "..."
            Height          =   255
            Left            =   4200
            TabIndex        =   18
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox txtSQLPlus 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Text            =   "C:\WINNT\SYSTEM32"
            Top             =   3120
            Width           =   3255
         End
         Begin VB.CheckBox chkSQLPlus 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "Or mess around with this Folder?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   16
            Top             =   2760
            Width           =   3255
         End
         Begin VB.CommandButton cmdFolderSQL 
            Caption         =   "..."
            Height          =   255
            Left            =   4200
            TabIndex        =   15
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtSQLPath 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   600
            TabIndex        =   14
            Text            =   "C:\WINNT"
            Top             =   2160
            Width           =   3255
         End
         Begin VB.CheckBox chkSQL 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Would you like to do something in this folder?"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label lblStep3_Text0 
            BackStyle       =   0  'Transparent
            Caption         =   "Sorry, i forgot them in the last Step. Here they are:"
            Height          =   1095
            Left            =   720
            TabIndex        =   26
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame frmStep 
         Caption         =   "Step 2"
         Height          =   4335
         Index           =   2
         Left            =   -2400
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "You love controls? I love controls too .."
            Height          =   2655
            Left            =   360
            TabIndex        =   27
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.Frame frmStep 
         Caption         =   "Step 1"
         Height          =   4215
         Index           =   1
         Left            =   -2400
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label lblStep1_Text2 
            BackStyle       =   0  'Transparent
            Height          =   1575
            Left            =   600
            TabIndex        =   22
            Top             =   2520
            Width           =   4335
         End
         Begin VB.Label lblStep1_Text1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hello there. Now we are in the Wizard.."
            Height          =   615
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame frmStep 
         Caption         =   "Step 0"
         Height          =   4335
         Index           =   0
         Left            =   -2400
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   5055
         Begin VB.ComboBox cmbInstallationType 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown-Liste
            TabIndex        =   12
            Top             =   3000
            Width           =   3255
         End
         Begin VB.Label lblStep0_Text0 
            BackStyle       =   0  'Transparent
            Caption         =   "Please choose the Installation."
            Height          =   1455
            Left            =   240
            TabIndex        =   20
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label lblHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the Wizard"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   4575
         End
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbInstallationType_Click()
    'Call LoadDataToWizard
    ' Hide or Show some Data, add more maxSteps, Enable some textboxes ..
    ' .. what you like. Like this:
    lblStep1_Text1 = "Hello there. Now you are inside the Wizard." & vbCrLf & vbCrLf & "You've chosen the " & gSettingsName(frmWizard.cmbInstallationType.ListIndex)
End Sub

Private Sub cmdBack_Click()
    gintStep = gintStep - 1
End Sub


Private Sub cmdBoring_Click()
    lblStep3_Text0.Caption = "Hello again .. :D "
    gintStep = 3
End Sub

Private Sub cmdCancle_Click()
    frmWizard.Hide
    Unload frmWizard
    End
 
End Sub


Private Sub cmdHelp_Click()
    MsgBox "Need Help?"
End Sub

Private Sub cmdNext_Click()
    
    Select Case gintStep
        ' do some validation, save something of get Data from File ..
        Case 0: Debug.Print "After Step 0"
        Case 1: Debug.Print "After Step 1"
        Case 2: Debug.Print "After Step 2"
        Case 3: Debug.Print "After Step 3"
        Case 4: Debug.Print "After Step 4"
        Case 5: Debug.Print "After Step 5 and end function"
        
    End Select
    
    If gintStep = gintStepMax Then
        
        ' call WorkWithTheDataAndDoSomething()
        ' Here goes the Code of function, that processes the collected Information.
        
        frmWizard.Hide
        Unload frmWizard
        
    End If
    
    gintStep = gintStep + 1
    
End Sub



Private Sub Form_Load()

    gintStepCheck = gintStepMax
    gintStep = 0 'Start-Step
    frmWizard.barSteps.Max = CInt(gintStepMax)
    Call StartUpWizard
    
    ' Write the Labels here, 'cause i don't know who to CrLf in the Property-Sheet
    lblHeader.Caption = "Welcome to the" & vbCrLf & " PLANET SOURCE CODE" & vbCrLf & "Installation Wizard"
    lblStep0_Text0.Caption = "This wizard will help you to install the newest release of the PLANET SOURCE CODE on your machine." & vbCrLf & vbCrLf & _
                            "Please choose your Installation-Type:"
End Sub






Private Sub tmrStep_Timer()
    
    If gintStep <> gintStepCheck And gintStep <= gintStepMax Then
        
        frmStep.Item(gintStep).Visible = True
        frmStep.Item (gintStepChec)
        
        With frmStep.Item(gintStepCheck)
            .Visible = False
            .Width = "1"
            .Height = 1
            .Left = "0"
            .Top = 0
            .BorderStyle = 0
            .BackColor = frmWizard.PictureBox.BackColor
        End With
        
        With frmStep.Item(gintStep)
            .Visible = True
            .Width = "5224,806"
            .Height = 4215
            .Left = "2728,682"
            .Top = 0
            .BorderStyle = 0
            .BackColor = frmWizard.PictureBox.BackColor
        End With
        

        gintStepCheck = gintStep
        
        If gintStep = gintStepMax Then
            frmWizard.cmdNext.Caption = "Finish"
            'frmWizard.cmdNext.Enabled = False
        Else
            frmWizard.cmdNext.Caption = "Next >"
            frmWizard.cmdNext.Enabled = True
        End If
    
        If gintStep = 0 Then
            frmWizard.cmdBack.Visible = False
            frmWizard.cmdBack.Enabled = False
        Else
            frmWizard.cmdBack.Visible = True
            frmWizard.cmdBack.Enabled = True
        End If
            
        
        
        frmWizard.barSteps.Value = gintStep
       frmWizard.Caption = "Installation Wizard (Step " & gintStep + 1 & " of " & gintStepMax + 1 & ")"
   
    End If
    


End Sub
