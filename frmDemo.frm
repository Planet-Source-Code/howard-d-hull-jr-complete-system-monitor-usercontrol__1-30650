VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{2BD531E5-B5CD-4EE1-997F-1D96891863EA}#2.0#0"; "dtSystemMonitor.ocx"
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Monitor"
   ClientHeight    =   7305
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SystemMonitor.dtSystemMonitor dtSystemMonitor1 
      Height          =   1065
      Left            =   150
      TabIndex        =   59
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1879
      VirtualMemoryDisplayed=   0   'False
      DiagramStyle    =   1
      HDDriveLetter   =   "C:\"
   End
   Begin VB.CheckBox chkTopMost 
      Caption         =   "Always on Top."
      Height          =   195
      Left            =   4050
      TabIndex        =   58
      Top             =   6870
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4380
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Hard Drive Monitor"
      Height          =   2025
      Left            =   150
      TabIndex        =   42
      Top             =   3810
      Width           =   3705
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Percent Available :"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   53
         Top             =   1650
         Width           =   1785
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   12
         Left            =   2100
         TabIndex        =   52
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   11
         Left            =   2100
         TabIndex        =   51
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Used Space :"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   50
         Top             =   1395
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Available Free Space :"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   49
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   10
         Left            =   2100
         TabIndex        =   48
         Top             =   1110
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   9
         Left            =   2100
         TabIndex        =   47
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total HD Size :"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   46
         Top             =   630
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Free Space :"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   45
         Top             =   885
         Width           =   1785
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   8
         Left            =   2100
         TabIndex        =   44
         Top             =   855
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   1155
      Left            =   180
      TabIndex        =   37
      Top             =   5910
      Width           =   3705
      Begin VB.ComboBox cmbDiagramStyle 
         Height          =   315
         ItemData        =   "frmDemo.frx":0000
         Left            =   1410
         List            =   "frmDemo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   300
         Width           =   2085
      End
      Begin VB.ComboBox cmbUpdateInterval 
         Height          =   315
         ItemData        =   "frmDemo.frx":0025
         Left            =   1380
         List            =   "frmDemo.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagram Style"
         Height          =   255
         Left            =   150
         TabIndex        =   57
         Top             =   330
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Update Speed"
         Height          =   255
         Left            =   150
         TabIndex        =   56
         Top             =   750
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Properties"
      Height          =   1725
      Left            =   4050
      TabIndex        =   28
      Top             =   4980
      Width           =   3585
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   225
         Left            =   1500
         TabIndex        =   34
         Top             =   1410
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.ComboBox cmbAppearance 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox cmbBorderStyle 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   690
         Width           =   2085
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   1500
         TabIndex        =   29
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Appearance"
         Height          =   165
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Border Style"
         Height          =   165
         Index           =   1
         Left            =   150
         TabIndex        =   32
         Top             =   765
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "System Monitor"
      Height          =   2565
      Left            =   150
      TabIndex        =   13
      Top             =   1170
      Width           =   3705
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   7
         Left            =   2100
         TabIndex        =   39
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Memory Load :"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   38
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Virtual Memory :"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   27
         Top             =   2250
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Virtual Memory :"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   2010
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Paging File :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   1650
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Paging File :"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   24
         Top             =   1410
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Physical Memory :"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Physical Memory :"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblSysMonitorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "CPU Load :"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   6
         Left            =   2100
         TabIndex        =   20
         Top             =   2220
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   5
         Left            =   2100
         TabIndex        =   19
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   4
         Left            =   2100
         TabIndex        =   18
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   3
         Left            =   2100
         TabIndex        =   17
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   16
         Top             =   1050
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   1
         Left            =   2100
         TabIndex        =   15
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label lblSysMonitor 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   0
         Left            =   2100
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Monitor Display Options"
      Height          =   3705
      Left            =   4050
      TabIndex        =   0
      Top             =   1170
      Width           =   3585
      Begin VB.CheckBox chkOption 
         Caption         =   "Hard Drive Free Space"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   41
         Top             =   2190
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   6
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   40
         Top             =   2100
         Width           =   375
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   5
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   36
         Top             =   1686
         Width           =   375
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Virtual Memory"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   35
         Top             =   1776
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   4
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   3210
         Width           =   375
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Scrolling"
         Height          =   195
         Index           =   4
         Left            =   450
         TabIndex        =   10
         Top             =   2910
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   3
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Background Grid"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "PageFile"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   1364
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "MEMORY Load"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   952
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "CPU Load"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   540
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   2
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   1274
         Width           =   375
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   1
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   862
         Width           =   375
      End
      Begin VB.PictureBox picControlColor 
         Height          =   375
         Index           =   0
         Left            =   2790
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   1
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Background Color"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   3270
         Width           =   2385
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Click Box to change Color"
         Height          =   225
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "&File"
      Begin VB.Menu mnuFILE_EXIT 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'// Win32 API for TopMost Form
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'// Private CONSTs for TopMost Form
Private Const HWND_TOPMOST = -1&
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1&
Private Const SWP_NOMOVE = &H2&
Private Const SWP_NOACTIVATE = &H10&
Private Const SWP_SHOWWINDOW = &H40&

Private Sub chkEnabled_Click()
    dtSystemMonitor1.Enabled = (chkEnabled.Value = vbChecked)
End Sub

Private Sub chkOption_Click(Index As Integer)
    
    Select Case Index
        Case 0  '// CPU Load
            dtSystemMonitor1.CPUMonitorDisplayed = (chkOption(Index).Value = vbChecked)
        
        Case 1  '// MEMORY Load
            dtSystemMonitor1.MemoryLoadDisplayed = (chkOption(Index).Value = vbChecked)
        
        Case 2  '// PageFile
            dtSystemMonitor1.PageFileDisplayed = (chkOption(Index).Value = vbChecked)
    
        Case 3  '// BG Grid
            dtSystemMonitor1.DisplayGrid = (chkOption(Index).Value = vbChecked)
            '// Enable/Disable the Scrolling Checkbox
            '// based on if the BG Grid is displayed
            If (chkOption(Index).Value = vbChecked) Then
                chkOption(4).Enabled = True
            Else
                chkOption(4).Enabled = False
            End If
        
        Case 4  '// Scrolling BG
            dtSystemMonitor1.ScrollGrid = (chkOption(Index).Value = vbChecked)
        
        Case 5  '// Virtual Memory
            dtSystemMonitor1.VirtualMemoryDisplayed = (chkOption(Index).Value = vbChecked)
        
        Case 6  '// HD
            dtSystemMonitor1.HDMonitorDisplayed = (chkOption(Index).Value = vbChecked)
        
    End Select
    
End Sub

Private Sub chkTopMost_Click()
    SetTopMostForm Me, (chkTopMost.Value = vbChecked)
End Sub

Private Sub chkVisible_Click()
    dtSystemMonitor1.Visible = (chkVisible.Value = vbChecked)
End Sub

Private Sub cmbAppearance_Click()
    dtSystemMonitor1.Appearance = cmbAppearance.ListIndex
End Sub

Private Sub cmbBorderStyle_Click()
    dtSystemMonitor1.BorderStyle = cmbBorderStyle.ListIndex
End Sub

Private Sub cmbDiagramStyle_Click()
    dtSystemMonitor1.DiagramStyle = cmbDiagramStyle.ListIndex
End Sub

Private Sub cmbUpdateInterval_Click()
    dtSystemMonitor1.UpdateInterval = cmbUpdateInterval.ItemData(cmbUpdateInterval.ListIndex)
End Sub

Private Sub Drive1_Change()
    dtSystemMonitor1.HDDriveLetter = Drive1.Drive
End Sub

Private Sub dtSystemMonitor1_Update(ByVal CPULoadPercent As Long, ByVal MemoryLoadPercent As Long, ByVal PhysicalMemoryTotal As Long, ByVal PhysicalMemoryAvailable As Long, ByVal PhysicalMemoryAvailablePercent As Single, ByVal PageFileTotal As Long, ByVal PageFileAvailable As Long, ByVal PageFileAvailablePercent As Single, ByVal VirtualMemoryTotal As Long, ByVal VirtualMemoryAvailable As Long, ByVal VirtualMemoryAvailablePercent As Single, ByVal HDTotalBytes As Currency, ByVal HDTotalFreeBytes As Currency, ByVal HDAvailableFreeBytes As Currency, ByVal HDTotalBytesUsed As Currency, ByVal HDAvailablePercent As Single)

    '// Use the usercontrol
    With dtSystemMonitor1
        lblSysMonitor(0).Caption = Format$(CPULoadPercent, "##0") & " %"
        lblSysMonitor(7).Caption = Format$(MemoryLoadPercent, "##0") & " %"
        lblSysMonitor(1).Caption = .FormatFilesize(PhysicalMemoryTotal)
        lblSysMonitor(2).Caption = .FormatFilesize(PhysicalMemoryAvailable)
        
        lblSysMonitor(3).Caption = .FormatFilesize(PageFileTotal)
        lblSysMonitor(4).Caption = .FormatFilesize(PageFileAvailable)
        
        lblSysMonitor(5).Caption = .FormatFilesize(VirtualMemoryTotal)
        lblSysMonitor(6).Caption = .FormatFilesize(VirtualMemoryAvailable)
    
        lblSysMonitor(8).Caption = .FormatFilesize(HDTotalFreeBytes)
        lblSysMonitor(9).Caption = .FormatFilesize(HDTotalBytes)
        lblSysMonitor(10).Caption = .FormatFilesize(HDAvailableFreeBytes)
        lblSysMonitor(11).Caption = .FormatFilesize(HDTotalBytesUsed)
        lblSysMonitor(12).Caption = Format$(HDAvailablePercent, "##0.0") & " %"
    End With
    
End Sub

Private Sub Form_Activate()
    SetTopMostForm Me, True
End Sub

Private Sub Form_Load()
    
    '//////////////////////////////////////
    '// Retreive current properties
    '//////////////////////////////////////
    '// Select the checkboxes
    With dtSystemMonitor1
        chkOption(0).Value = IIf(.CPUMonitorDisplayed, vbChecked, vbUnchecked)
        chkOption(1).Value = IIf(.MemoryLoadDisplayed, vbChecked, vbUnchecked)
        chkOption(2).Value = IIf(.PageFileDisplayed, vbChecked, vbUnchecked)
        chkOption(3).Value = IIf(.DisplayGrid, vbChecked, vbUnchecked)
        chkOption(4).Value = IIf(.ScrollGrid, vbChecked, vbUnchecked)
        chkOption(5).Value = IIf(.VirtualMemoryDisplayed, vbChecked, vbUnchecked)
        chkOption(6).Value = IIf(.HDMonitorDisplayed, vbChecked, vbUnchecked)
    End With
        
    '// Change picturebox colors
    picControlColor(0).BackColor = dtSystemMonitor1.CPUMonitorColor
    picControlColor(1).BackColor = dtSystemMonitor1.MemoryLoadColor
    picControlColor(2).BackColor = dtSystemMonitor1.PageFileColor
    picControlColor(3).BackColor = dtSystemMonitor1.DisplayGridColor
    picControlColor(4).BackColor = dtSystemMonitor1.BackColor
    picControlColor(5).BackColor = dtSystemMonitor1.VirtualMemoryColor
    picControlColor(6).BackColor = dtSystemMonitor1.HDMonitorColor
    '//////////////////////////////////////
    
    
    '// Initial form setup
    cmbDiagramStyle.ListIndex = dtSystemMonitor1.DiagramStyle
    
    cmbUpdateInterval.ListIndex = dtSystemMonitor1.DiagramStyle
    
    cmbAppearance.AddItem "Flat"
    cmbAppearance.AddItem "3D"
    cmbAppearance.ListIndex = dtSystemMonitor1.Appearance
    
    cmbBorderStyle.AddItem "None"
    cmbBorderStyle.AddItem "Fixed Single"
    cmbBorderStyle.ListIndex = dtSystemMonitor1.BorderStyle
    
    '// Get the current system levels
    SetupInitialValues
    
End Sub

Private Sub mnuFILE_EXIT_Click()
    Unload Me
End Sub

Private Sub picControlColor_Click(Index As Integer)
On Error GoTo Err_picControlColor_Click

    CommonDialog1.CancelError = True
    CommonDialog1.Color = picControlColor(Index).BackColor
    CommonDialog1.Flags = cdlCCFullOpen Or cdlCCRGBInit
    CommonDialog1.ShowColor

    '// Change Picturebox to new color
    picControlColor(Index).BackColor = CommonDialog1.Color
    
    '// Update appropriate Property with new color
    Select Case Index
        Case 0: dtSystemMonitor1.CPUMonitorColor = picControlColor(Index).BackColor
        Case 1: dtSystemMonitor1.MemoryLoadColor = picControlColor(Index).BackColor
        Case 2: dtSystemMonitor1.PageFileColor = picControlColor(Index).BackColor
        Case 3: dtSystemMonitor1.DisplayGridColor = picControlColor(Index).BackColor
        Case 4: dtSystemMonitor1.BackColor = picControlColor(Index).BackColor
        Case 5: dtSystemMonitor1.VirtualMemoryColor = picControlColor(Index).BackColor
        Case 6: dtSystemMonitor1.HDMonitorColor = picControlColor(Index).BackColor
    End Select
    Exit Sub
    
Err_picControlColor_Click:
    If Err.Number = 32755 Then
        '// CancelError - Do nothing
    Else
        '// Other Error - Display Error
        MsgBox "Error occured! Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly Or vbExclamation, "Error"
    End If
    Err.Clear
    
End Sub


Private Sub SetTopMostForm(frmForm As Form, bTopMost As Boolean)
    If bTopMost Then
        'set this form always on top
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        'set this form always on top
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub


Private Sub SetupInitialValues()
Dim CPULoadPercent                  As Long
Dim MemoryLoadPercent               As Long
Dim PhysicalMemoryTotal             As Long
Dim PhysicalMemoryAvailable         As Long
Dim PhysicalMemoryAvailablePercent  As Single
Dim PageFileTotal                   As Long
Dim PageFileAvailable               As Long
Dim PageFileAvailablePercent        As Single
Dim VirtualMemoryTotal              As Long
Dim VirtualMemoryAvailable          As Long
Dim VirtualMemoryAvailablePercent   As Single
Dim HDTotalBytes                    As Currency
Dim HDTotalFreeBytes                As Currency
Dim HDAvailableFreeBytes            As Currency
Dim HDTotalBytesUsed                As Currency
Dim HDAvailablePercent              As Single

    '// Use the usercontrol
    With dtSystemMonitor1
        Call .GetCurrentSystemLevels(CPULoadPercent, MemoryLoadPercent, _
                    PhysicalMemoryTotal, PhysicalMemoryAvailable, PhysicalMemoryAvailablePercent, _
                    PageFileTotal, PageFileAvailable, PageFileAvailablePercent, _
                    VirtualMemoryTotal, VirtualMemoryAvailable, VirtualMemoryAvailablePercent, _
                    HDTotalBytes, HDTotalFreeBytes, HDAvailableFreeBytes, HDTotalBytesUsed, HDAvailablePercent)
                    
        lblSysMonitor(0).Caption = Format$(CPULoadPercent, "##0") & " %"
        lblSysMonitor(7).Caption = Format$(MemoryLoadPercent, "##0") & " %"
        lblSysMonitor(1).Caption = .FormatFilesize(PhysicalMemoryTotal)
        lblSysMonitor(2).Caption = .FormatFilesize(PhysicalMemoryAvailable)
        
        lblSysMonitor(3).Caption = .FormatFilesize(PageFileTotal)
        lblSysMonitor(4).Caption = .FormatFilesize(PageFileAvailable)
        
        lblSysMonitor(5).Caption = .FormatFilesize(VirtualMemoryTotal)
        lblSysMonitor(6).Caption = .FormatFilesize(VirtualMemoryAvailable)
    
        lblSysMonitor(8).Caption = .FormatFilesize(HDTotalFreeBytes)
        lblSysMonitor(9).Caption = .FormatFilesize(HDTotalBytes)
        lblSysMonitor(10).Caption = .FormatFilesize(HDAvailableFreeBytes)
        lblSysMonitor(11).Caption = .FormatFilesize(HDTotalBytesUsed)
        lblSysMonitor(12).Caption = Format$(HDAvailablePercent, "##0.0") & " %"
    End With
    
End Sub

