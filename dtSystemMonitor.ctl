VERSION 5.00
Begin VB.UserControl dtSystemMonitor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   1140
   End
End
Attribute VB_Name = "dtSystemMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//---------------------------------------------------------------------------
'// Border
Public Enum eBorderStyles
    [None]
    [Fixed Single]      '// Default
End Enum
'//---------------------------------------------------------------------------

'//---------------------------------------------------------------------------
'// Apperances
Public Enum eApperances
    [Flat]
    [3D]                '// Default
End Enum
'//---------------------------------------------------------------------------

'//---------------------------------------------------------------------------
'// How the data should be displayed
Public Enum DiagramStyles
    TYPE_LINE
    TYPE_POINT
    TYPE_BAR
End Enum
'//---------------------------------------------------------------------------

'Default Property Values:
Const m_def_HDDriveLetter = ""
Const m_def_HDMonitorDisplayed = True
Const m_def_HDMonitorColor = &HFF80FF
Const m_def_DiagramStyle = DiagramStyles.TYPE_LINE
Const m_def_ScrollGrid = True
Const m_def_UpdateInterval = 1000
Const m_def_DisplayGrid = True
Const m_def_DisplayGridColor = &H8000&
Const m_def_CPUMonitorColor = &HFF00&
Const m_def_CPUMonitorDisplayed = True
Const m_def_VirtualMemoryColor = &HFF0000
Const m_def_VirtualMemoryDisplayed = True
Const m_def_MemoryLoadColor = &HFFFF&
Const m_def_MemoryLoadDisplayed = True
Const m_def_PageFileColor = &HFF&
Const m_def_PageFileDisplayed = True
Const m_def_Appearance = eApperances.[3D]
Const m_def_BorderStyle = eBorderStyles.[Fixed Single]
Const m_def_BackColor = &H0

'Property Variables:
Dim m_HDDriveLetter             As String
Dim m_DiagramStyle              As DiagramStyles
Dim m_ScrollGrid                As Boolean
Dim m_UpdateInterval            As Long
Dim m_DisplayGridColor          As OLE_COLOR
Dim m_DisplayGrid               As Boolean
Dim m_CPUMonitorColor           As OLE_COLOR
Dim m_CPUMonitorDisplayed       As Boolean
Dim m_HDMonitorColor            As OLE_COLOR
Dim m_HDMonitorDisplayed        As Boolean
Dim m_PageFileColor             As OLE_COLOR
Dim m_PageFileDisplayed         As Boolean
Dim m_MemoryLoadColor           As OLE_COLOR
Dim m_MemoryLoadDisplayed       As Boolean
Dim m_VirtualMemoryColor        As OLE_COLOR
Dim m_VirtualMemoryDisplayed    As Boolean
Dim m_Appearance                As eApperances
Dim m_BorderStyle               As eBorderStyles
Dim m_BackColor                 As OLE_COLOR
Dim m_Enabled                   As Boolean

'// Private Usercontrol Variables
Private m_GridOffset            As Long
Private m_IsWinNT               As Boolean
Private lngStartPosition        As Long     'Needed to not to display first zero values when starting a new diagram

'// Dimension Array of Info
Dim SystemMonitorArray()        As SystemMonitorStruct

'Event Declarations:
Event Update(ByVal CPULoadPercent As Long, ByVal MemoryLoadPercent As Long, _
             ByVal PhysicalMemoryTotal As Long, ByVal PhysicalMemoryAvailable As Long, ByVal PhysicalMemoryAvailablePercent As Single, _
             ByVal PageFileTotal As Long, ByVal PageFileAvailable As Long, ByVal PageFileAvailablePercent As Single, _
             ByVal VirtualMemoryTotal As Long, ByVal VirtualMemoryAvailable As Long, ByVal VirtualMemoryAvailablePercent As Single, _
             ByVal HDTotalBytes As Currency, ByVal HDTotalFreeBytes As Currency, _
             ByVal HDAvailableFreeBytes As Currency, ByVal HDTotalBytesUsed As Currency, _
             ByVal HDAvailablePercent As Single)
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    RedrawControl
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As eApperances
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eApperances)
    UserControl.Appearance() = New_Appearance
    RedrawControl
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As eBorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As eBorderStyles)
    UserControl.BorderStyle() = New_BorderStyle
    RedrawControl
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CPUMonitorDisplayed() As Boolean
Attribute CPUMonitorDisplayed.VB_Description = "Set/Return a value that determines whether the CPU Level is displayed"
    CPUMonitorDisplayed = m_CPUMonitorDisplayed
End Property

Public Property Let CPUMonitorDisplayed(ByVal New_CPUMonitorDisplayed As Boolean)
    m_CPUMonitorDisplayed = New_CPUMonitorDisplayed
    RedrawControl
    PropertyChanged "CPUMonitorDisplayed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get PageFileDisplayed() As Boolean
Attribute PageFileDisplayed.VB_Description = "Set/Return a value that determines whether the Available PageFile Level is displayed"
    PageFileDisplayed = m_PageFileDisplayed
End Property

Public Property Let PageFileDisplayed(ByVal New_PageFileDisplayed As Boolean)
    m_PageFileDisplayed = New_PageFileDisplayed
    RedrawControl
    PropertyChanged "PageFileDisplayed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get VirtualMemoryDisplayed() As Boolean
Attribute VirtualMemoryDisplayed.VB_Description = "Set/Return a value that determines whether the Virtual Memory level is displayed"
    VirtualMemoryDisplayed = m_VirtualMemoryDisplayed
End Property

Public Property Let VirtualMemoryDisplayed(ByVal New_VirtualMemoryDisplayed As Boolean)
    m_VirtualMemoryDisplayed = New_VirtualMemoryDisplayed
    RedrawControl
    PropertyChanged "VirtualMemoryDisplayed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CPUMonitorColor() As OLE_COLOR
Attribute CPUMonitorColor.VB_Description = "Return/Set the color to use to display the CPU Level"
    CPUMonitorColor = m_CPUMonitorColor
End Property

Public Property Let CPUMonitorColor(ByVal New_CPUMonitorColor As OLE_COLOR)
    m_CPUMonitorColor = New_CPUMonitorColor
    RedrawControl
    PropertyChanged "CPUMonitorColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MemoryLoadDisplayed() As Boolean
Attribute MemoryLoadDisplayed.VB_Description = "Set/Return a value that determines whether the Available Memory is displayed"
    MemoryLoadDisplayed = m_MemoryLoadDisplayed
End Property

Public Property Let MemoryLoadDisplayed(ByVal New_MemoryLoadDisplayed As Boolean)
    m_MemoryLoadDisplayed = New_MemoryLoadDisplayed
    RedrawControl
    PropertyChanged "MemoryLoadDisplayed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HDMonitorDisplayed() As Boolean
    HDMonitorDisplayed = m_HDMonitorDisplayed
End Property

Public Property Let HDMonitorDisplayed(ByVal New_HDMonitorDisplayed As Boolean)
    m_HDMonitorDisplayed = New_HDMonitorDisplayed
    RedrawControl
    PropertyChanged "HDMonitorDisplayed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get HDDriveLetter() As String
    HDDriveLetter = m_HDDriveLetter
End Property

Public Property Let HDDriveLetter(ByVal New_HDDriveLetter As String)
    '// Quick and dirty validation
    If New_HDDriveLetter = "" Then Exit Property
    If Len(New_HDDriveLetter) > 1 Then _
        New_HDDriveLetter = UCase(Left(New_HDDriveLetter, 1)) & ":\"
    m_HDDriveLetter = New_HDDriveLetter
    PropertyChanged "HDDriveLetter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HDMonitorColor() As OLE_COLOR
    HDMonitorColor = m_HDMonitorColor
End Property

Public Property Let HDMonitorColor(ByVal New_HDMonitorColor As OLE_COLOR)
    m_HDMonitorColor = New_HDMonitorColor
    PropertyChanged "HDMonitorColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MemoryLoadColor() As OLE_COLOR
Attribute MemoryLoadColor.VB_Description = "Return/Set the color to use to display the level of Available Memory "
    MemoryLoadColor = m_MemoryLoadColor
End Property

Public Property Let MemoryLoadColor(ByVal New_MemoryLoadColor As OLE_COLOR)
    m_MemoryLoadColor = New_MemoryLoadColor
    RedrawControl
    PropertyChanged "MemoryColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get PageFileColor() As OLE_COLOR
Attribute PageFileColor.VB_Description = "Return/Set the color to use to display the level of Available PageFile"
    PageFileColor = m_PageFileColor
End Property

Public Property Let PageFileColor(ByVal New_PageFileColor As OLE_COLOR)
    m_PageFileColor = New_PageFileColor
    RedrawControl
    PropertyChanged "PageFileColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get VirtualMemoryColor() As OLE_COLOR
Attribute VirtualMemoryColor.VB_Description = "Return/Set the color to use to display the level of Available Virtual Memory"
    VirtualMemoryColor = m_VirtualMemoryColor
End Property

Public Property Let VirtualMemoryColor(ByVal New_VirtualMemoryColor As OLE_COLOR)
    m_VirtualMemoryColor = New_VirtualMemoryColor
    RedrawControl
    PropertyChanged "VirtualMemoryColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    RedrawControl
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get DisplayGrid() As Boolean
    DisplayGrid = m_DisplayGrid
End Property

Public Property Let DisplayGrid(ByVal New_DisplayGrid As Boolean)
    m_DisplayGrid = New_DisplayGrid
    RedrawControl
    PropertyChanged "DisplayGrid"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisplayGridColor() As OLE_COLOR
    DisplayGridColor = m_DisplayGridColor
End Property

Public Property Let DisplayGridColor(ByVal New_DisplayGridColor As OLE_COLOR)
    m_DisplayGridColor = New_DisplayGridColor
    RedrawControl
    PropertyChanged "DisplayGridColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1000
Public Property Get UpdateInterval() As Long
    UpdateInterval = m_UpdateInterval
End Property

Public Property Let UpdateInterval(ByVal New_UpdateInterval As Long)
    m_UpdateInterval = New_UpdateInterval
    '// Only change the Timer Interval if we are Not in IDE
    If Ambient.UserMode Then Timer1.Interval = m_UpdateInterval
    PropertyChanged "UpdateInterval"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ScrollGrid() As Boolean
    ScrollGrid = m_ScrollGrid
End Property

Public Property Let ScrollGrid(ByVal New_ScrollGrid As Boolean)
    m_ScrollGrid = New_ScrollGrid
    RedrawControl
    PropertyChanged "ScrollGrid"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Reset() As Boolean
    m_GridOffset = 0
    RedrawControl
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get DiagramStyle() As DiagramStyles
    DiagramStyle = m_DiagramStyle
End Property

Public Property Let DiagramStyle(ByVal New_DiagramStyle As DiagramStyles)
    m_DiagramStyle = New_DiagramStyle
    RedrawControl
    PropertyChanged "DiagramStyle"
End Property



'// Yeh I know... Timers... AAAGGGHHHH....
'// I normally use the ccrpHiRes Timer for ActiveX controls but didn't
'// want to force others to DL & install it just to use this control.
Private Sub Timer1_Timer()
Dim l                           As Long

    '// Exit Sub if the control is Disabled
    If Not UserControl.Enabled Then Exit Sub
    
    '// Exit if we are in IDE
    If Not Ambient.UserMode Then Exit Sub
    
    '// Update m_GridOffset to move the BG if we are scrolling
    m_GridOffset = m_GridOffset - 1
    
    '// Move all values from array one position lower
    '// A faster way would be to use the CopyMem function.
    For l = 1 To UserControl.ScaleWidth - 1
        SystemMonitorArray(l - 1) = SystemMonitorArray(l)
    Next
        
    '// Increment StartPosition
    If lngStartPosition >= 1 Then lngStartPosition = lngStartPosition - 1
    
    '// Get latest values
    Call MEMORYQuery(SystemMonitorArray(l - 1))
    Call HDQuery(m_HDDriveLetter, SystemMonitorArray(l - 1))
    SystemMonitorArray(l - 1).CPULoadPercent = CPUQuery(m_IsWinNT)
        
    '// Draw
    RedrawControl
    
    '// Trigger Event
    With SystemMonitorArray(l - 1)
        RaiseEvent Update(.CPULoadPercent, .MemoryLoadPercent, _
                          .PhysicalMemoryTotal, .PhysicalMemoryAvailable, .PhysicalMemoryAvailablePercent, _
                          .PageFileTotal, .PageFileAvailable, .PageFileAvailablePercent, _
                          .VirtualMemoryTotal, .VirtualMemoryAvailable, .VirtualMemoryAvailablePercent, _
                          .HDTotalBytes, .HDTotalFreeBytes, .HDAvailableFreeBytes, _
                          .HDTotalBytesUsed, .HDAvailablePercent)
    End With
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CPUMonitorDisplayed = m_def_CPUMonitorDisplayed
    m_PageFileDisplayed = m_def_PageFileDisplayed
    m_VirtualMemoryDisplayed = m_def_VirtualMemoryDisplayed
    m_CPUMonitorColor = m_def_CPUMonitorColor
    m_MemoryLoadDisplayed = m_def_MemoryLoadDisplayed
    m_MemoryLoadColor = m_def_MemoryLoadColor
    m_PageFileColor = m_def_PageFileColor
    m_VirtualMemoryColor = m_def_VirtualMemoryColor
    m_DisplayGrid = m_def_DisplayGrid
    m_DisplayGridColor = m_def_DisplayGridColor
    m_UpdateInterval = m_def_UpdateInterval
    m_ScrollGrid = m_def_ScrollGrid
    m_DiagramStyle = m_def_DiagramStyle
    m_HDMonitorDisplayed = m_def_HDMonitorDisplayed
    m_HDMonitorColor = m_def_HDMonitorColor
    m_HDDriveLetter = Left(DetermineWindowsFolder, 3)
    UserControl.Appearance = m_def_Appearance
    UserControl.BorderStyle = m_def_BorderStyle
    UserControl.BackColor = m_def_BackColor
    
    RedrawControl
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CPUMonitorDisplayed = PropBag.ReadProperty("CPUMonitorDisplayed", m_def_CPUMonitorDisplayed)
    m_PageFileDisplayed = PropBag.ReadProperty("PageFileDisplayed", m_def_PageFileDisplayed)
    m_VirtualMemoryDisplayed = PropBag.ReadProperty("VirtualMemoryDisplayed", m_def_VirtualMemoryDisplayed)
    m_CPUMonitorColor = PropBag.ReadProperty("CPUMonitorColor", m_def_CPUMonitorColor)
    m_MemoryLoadDisplayed = PropBag.ReadProperty("MemoryLoadDisplayed", m_def_MemoryLoadDisplayed)
    m_MemoryLoadColor = PropBag.ReadProperty("MemoryColor", m_def_MemoryLoadColor)
    m_PageFileColor = PropBag.ReadProperty("PageFileColor", m_def_PageFileColor)
    m_VirtualMemoryColor = PropBag.ReadProperty("VirtualMemoryColor", m_def_VirtualMemoryColor)
    m_DisplayGrid = PropBag.ReadProperty("DisplayGrid", m_def_DisplayGrid)
    m_DisplayGridColor = PropBag.ReadProperty("DisplayGridColor", m_def_DisplayGridColor)
    m_UpdateInterval = PropBag.ReadProperty("UpdateInterval", m_def_UpdateInterval)
    m_ScrollGrid = PropBag.ReadProperty("ScrollGrid", m_def_ScrollGrid)
    m_DiagramStyle = PropBag.ReadProperty("DiagramStyle", m_def_DiagramStyle)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_HDMonitorDisplayed = PropBag.ReadProperty("HDMonitorDisplayed", m_def_HDMonitorDisplayed)
    m_HDMonitorColor = PropBag.ReadProperty("HDMonitorColor", m_def_HDMonitorColor)
    m_HDDriveLetter = PropBag.ReadProperty("HDDriveLetter", m_def_HDDriveLetter)
    
    ConfigureControl
End Sub

Private Sub UserControl_Resize()
    RedrawControl
End Sub

Private Sub UserControl_Terminate()
    '// Shut down the CPU Querying
    CPUTerminate m_IsWinNT
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("CPUMonitorDisplayed", m_CPUMonitorDisplayed, m_def_CPUMonitorDisplayed)
    Call PropBag.WriteProperty("PageFileDisplayed", m_PageFileDisplayed, m_def_PageFileDisplayed)
    Call PropBag.WriteProperty("VirtualMemoryDisplayed", m_VirtualMemoryDisplayed, m_def_VirtualMemoryDisplayed)
    Call PropBag.WriteProperty("CPUMonitorColor", m_CPUMonitorColor, m_def_CPUMonitorColor)
    Call PropBag.WriteProperty("MemoryLoadDisplayed", m_MemoryLoadDisplayed, m_def_MemoryLoadDisplayed)
    Call PropBag.WriteProperty("MemoryColor", m_MemoryLoadColor, m_def_MemoryLoadColor)
    Call PropBag.WriteProperty("PageFileColor", m_PageFileColor, m_def_PageFileColor)
    Call PropBag.WriteProperty("VirtualMemoryColor", m_VirtualMemoryColor, m_def_VirtualMemoryColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("DisplayGrid", m_DisplayGrid, m_def_DisplayGrid)
    Call PropBag.WriteProperty("DisplayGridColor", m_DisplayGridColor, m_def_DisplayGridColor)
    Call PropBag.WriteProperty("UpdateInterval", m_UpdateInterval, m_def_UpdateInterval)
    Call PropBag.WriteProperty("ScrollGrid", m_ScrollGrid, m_def_ScrollGrid)
    Call PropBag.WriteProperty("DiagramStyle", m_DiagramStyle, m_def_DiagramStyle)
    Call PropBag.WriteProperty("HDMonitorDisplayed", m_HDMonitorDisplayed, m_def_HDMonitorDisplayed)
    Call PropBag.WriteProperty("HDMonitorColor", m_HDMonitorColor, m_def_HDMonitorColor)
    Call PropBag.WriteProperty("HDDriveLetter", m_HDDriveLetter, m_def_HDDriveLetter)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub


Private Sub RedrawControl()
    
    UserControl.Cls
    UserControl.BackColor = m_BackColor
    
    '// Display Grid
    If m_DisplayGrid Then DrawGrid
    
    '// We only want to draw the system monitor data
    '// when running
    If Ambient.UserMode Then
        '// Graph the enabled functions
        If m_CPUMonitorDisplayed Then GraphCPULoad
        If m_MemoryLoadDisplayed Then GraphMemoryLoad
        If m_PageFileDisplayed Then GraphPageFile
        If m_VirtualMemoryDisplayed Then GraphVirtualMemory
        If m_HDMonitorDisplayed Then GraphHardDrive
    End If
    UserControl.Refresh
    
End Sub


Private Sub DrawGrid()
Dim X       As Long
Dim Y       As Long
Const VertSplits    As Long = 10
Const HorzSplits    As Long = 5
Const const_tolerance = 0.0001 'Used to fix last line tolerance problem in some cases

    '// Draw the Grid depending on if the BG is scrolling
    If m_ScrollGrid Then
        '// Draw Vertical Lines OffSet
        For X = m_GridOffset To UserControl.ScaleWidth - 1 Step ((UserControl.ScaleWidth - 1) / (VertSplits + 1))
            UserControl.Line (X, 0)-(X, UserControl.ScaleHeight), m_DisplayGridColor
        Next
    Else
        '// Draw Vertical Lines
        For X = 0 To UserControl.ScaleWidth - 1 Step ((UserControl.ScaleWidth - 1) / (VertSplits + 1))
            UserControl.Line (X, 0)-(X, UserControl.ScaleHeight), m_DisplayGridColor
        Next
    End If
    
    '// Draw Horizontal Lines
    For Y = 0 To UserControl.ScaleHeight - 1 Step ((UserControl.ScaleHeight - 1) / (HorzSplits + 1))
        UserControl.Line (0, Y)-(UserControl.ScaleWidth, Y), m_DisplayGridColor
    Next

    '// Box Around Edge
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_DisplayGridColor, B
    
    '// Reset m_GridOffset, when first line is not visible anymore
    If m_GridOffset <= -Int((UserControl.ScaleWidth - 1 / (HorzSplits + 1))) Then
        m_GridOffset = 0
    End If
    
End Sub


'// Draw the Array Of CPU Load values
Private Sub GraphCPULoad()
Dim X           As Long
Dim Y           As Long
Dim y2          As Long

    'Draw line diagram only if theres 2 or more values defined
    If lngStartPosition <= UserControl.ScaleWidth - 1 Then
        
        Select Case m_DiagramStyle

            Case TYPE_LINE
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).CPULoadPercent / 100) * UserControl.ScaleHeight)
                    y2 = UserControl.ScaleHeight - ((SystemMonitorArray(X + 1).CPULoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X + 1, y2), m_CPUMonitorColor
                Next

            Case TYPE_POINT
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).CPULoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.PSet (X + 1, Y), m_CPUMonitorColor
                Next

            Case TYPE_BAR
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).CPULoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X, UserControl.ScaleHeight), m_CPUMonitorColor
                Next
            
        End Select
    End If
    
End Sub


'// Draw the Array Of CPU Load values
Private Sub GraphMemoryLoad()
Dim X           As Long
Dim Y           As Long
Dim y2          As Long
    
    'Draw line diagram only if theres 2 or more values defined
    If lngStartPosition <= UserControl.ScaleWidth - 1 Then
        
        Select Case m_DiagramStyle

            Case TYPE_LINE
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).MemoryLoadPercent / 100) * UserControl.ScaleHeight)
                    y2 = UserControl.ScaleHeight - ((SystemMonitorArray(X + 1).MemoryLoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X + 1, y2), m_MemoryLoadColor
                Next

            Case TYPE_POINT
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).MemoryLoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.PSet (X + 1, Y), m_MemoryLoadColor
                Next

            Case TYPE_BAR
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).MemoryLoadPercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X, UserControl.ScaleHeight), m_MemoryLoadColor
                Next
            
        End Select
    End If
    
End Sub


'// Draw Page File Data
Private Sub GraphPageFile()
Dim X           As Long
Dim Y           As Long
Dim y2          As Long

    'Draw line diagram only if theres 2 or more values defined
    If lngStartPosition <= UserControl.ScaleWidth - 1 Then
        
        Select Case m_DiagramStyle

            Case TYPE_LINE
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).PageFileAvailablePercent / 100) * UserControl.ScaleHeight)
                    y2 = UserControl.ScaleHeight - ((SystemMonitorArray(X + 1).PageFileAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X + 1, y2), m_PageFileColor
                Next

            Case TYPE_POINT
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).PageFileAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.PSet (X + 1, Y), m_PageFileColor
                Next

            Case TYPE_BAR
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).PageFileAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X, UserControl.ScaleHeight), m_PageFileColor
                Next
            
        End Select
    End If
    
End Sub


'// Draw the Array Of Virtual Memory values
Private Sub GraphVirtualMemory()
Dim X           As Long
Dim Y           As Long
Dim y2          As Long

    'Draw line diagram only if theres 2 or more values defined
    If lngStartPosition <= UserControl.ScaleWidth - 1 Then
        
        Select Case m_DiagramStyle

            Case TYPE_LINE
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).VirtualMemoryAvailablePercent / 100) * UserControl.ScaleHeight)
                    y2 = UserControl.ScaleHeight - ((SystemMonitorArray(X + 1).VirtualMemoryAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X + 1, y2), m_VirtualMemoryColor
                Next

            Case TYPE_POINT
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).VirtualMemoryAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.PSet (X + 1, Y), m_VirtualMemoryColor
                Next

            Case TYPE_BAR
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).VirtualMemoryAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X, UserControl.ScaleHeight), m_VirtualMemoryColor
                Next
            
        End Select
    End If

End Sub


'// Draw the Array Of HD bytes free
Private Sub GraphHardDrive()
Dim X           As Long
Dim Y           As Long
Dim y2          As Long
    
    'Draw line diagram only if theres 2 or more values defined
    If lngStartPosition <= UserControl.ScaleWidth - 1 Then
        
        Select Case m_DiagramStyle

            Case TYPE_LINE
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).HDAvailablePercent / 100) * UserControl.ScaleHeight)
                    y2 = UserControl.ScaleHeight - ((SystemMonitorArray(X + 1).HDAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X + 1, y2), m_def_HDMonitorColor
                Next

            Case TYPE_POINT
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).HDAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.PSet (X + 1, Y), m_def_HDMonitorColor
                Next

            Case TYPE_BAR
                For X = lngStartPosition + 1 To UserControl.ScaleWidth - 2
                    Y = UserControl.ScaleHeight - ((SystemMonitorArray(X).HDAvailablePercent / 100) * UserControl.ScaleHeight)
                    UserControl.Line (X, Y)-(X, UserControl.ScaleHeight), m_def_HDMonitorColor
                Next
            
        End Select
    End If
    


End Sub


'// Routine to setup inital variables and set the Timer
Private Sub ConfigureControl()
Dim lngElements     As Long

    '// Configure the Timer if we are NOT in Design mode / IDE
    If Ambient.UserMode Then
        Timer1.Interval = m_UpdateInterval
        Timer1.Enabled = True
    End If
    
    '// Dimension Array of System Monitor Info
    lngElements = UserControl.ScaleX(UserControl.Width, vbTwips, vbPixels)
    ReDim SystemMonitorArray(lngElements) As SystemMonitorStruct
    
    '// Check to see if user is using Windows NT/2K/Me/XP?
    '// We have to use different methods to get
    '// CPU usage based on the OS Version.
    '// MS can never make anything simple :)
    m_IsWinNT = IsWinNTInstalled
    
    '// Initialize the CPU Monitor.
    CPUInitialize m_IsWinNT
    
    '// Setup lngStartPosition to the pixel just before the end of the usercontrol
    lngStartPosition = UserControl.ScaleWidth - 1
    
    '// Set Usercontrol Appearance... etc...
    UserControl.Appearance = m_Appearance
    UserControl.BorderStyle = m_BorderStyle
    UserControl.BackColor = m_BackColor
    UserControl.Enabled = m_Enabled
    
End Sub


'// Convert a number into K, MB or GB formatted string
Public Function FormatFilesize(nValue As Variant) As String
'// We pass the value as variant to work with Single, Double & Currency variables
'// We devide by 1024 to convert bytes to K
'// We devide by 1048576 to convert bytes to MB
'// We Devide by 1073741824 to convert bytes to GB

    If nValue <= 1024 Then
        '// Upto 1K
        FormatFilesize = Format$(nValue, "#,##0 B")
    ElseIf nValue > 1024 And nValue < 1048576 Then
        '// From 1K+1 to 1MB-1
        FormatFilesize = Format$(nValue / 1024, "###,###,##0 K")
    ElseIf nValue > 1048576 And nValue < 1073741824 Then
        '// From 1MB +1 to 1GB -1
        FormatFilesize = Format$(nValue / 1048576, "###,###,##0.00 MB")
    Else
        '// Greater than 1GB
        FormatFilesize = Format$(nValue / 1073741824, "###,###,##0.00 GB")
    End If
    
End Function


'// User callable routine to get the current system levels.
'// If the developer doesn't want to use the Update event.
Public Sub GetCurrentSystemLevels(ByRef CPULoadPercent As Long, ByRef MemoryLoadPercent As Long, _
             ByRef PhysicalMemoryTotal As Long, ByRef PhysicalMemoryAvailable As Long, ByRef PhysicalMemoryAvailablePercent As Single, _
             ByRef PageFileTotal As Long, ByRef PageFileAvailable As Long, ByRef PageFileAvailablePercent As Single, _
             ByRef VirtualMemoryTotal As Long, ByRef VirtualMemoryAvailable As Long, ByRef VirtualMemoryAvailablePercent As Single, _
             ByRef HDTotalBytes As Currency, ByRef HDTotalFreeBytes As Currency, _
             ByRef HDAvailableFreeBytes As Currency, ByRef HDTotalBytesUsed As Currency, _
             ByRef HDAvailablePercent As Single)
             
Dim i       As Long

    '// Move all values from array one position lower
    '// A faster way would be to use the CopyMem function.
    For i = 1 To UserControl.ScaleWidth - 1
        SystemMonitorArray(i - 1) = SystemMonitorArray(i)
    Next
    
    '// Increment StartPosition
    If lngStartPosition >= 1 Then lngStartPosition = lngStartPosition - 1
    
    '// Get latest values
    Call MEMORYQuery(SystemMonitorArray(i - 1))
    Call HDQuery(m_HDDriveLetter, SystemMonitorArray(i - 1))
    SystemMonitorArray(i - 1).CPULoadPercent = CPUQuery(m_IsWinNT)
    
    '// Trigger Event
    With SystemMonitorArray(i - 1)
        CPULoadPercent = .CPULoadPercent
        MemoryLoadPercent = .MemoryLoadPercent
        
        PhysicalMemoryTotal = .PhysicalMemoryTotal
        PhysicalMemoryAvailable = .PhysicalMemoryAvailable
        PhysicalMemoryAvailablePercent = .PhysicalMemoryAvailablePercent
        
        PageFileTotal = .PageFileTotal
        PageFileAvailable = .PageFileAvailable
        PageFileAvailablePercent = .PageFileAvailablePercent
        
        VirtualMemoryTotal = .VirtualMemoryTotal
        VirtualMemoryAvailable = .VirtualMemoryAvailable
        VirtualMemoryAvailablePercent = .VirtualMemoryAvailablePercent
        
        HDTotalBytes = .HDTotalBytes
        HDTotalFreeBytes = .HDTotalFreeBytes
        HDAvailableFreeBytes = .HDAvailableFreeBytes
        HDAvailablePercent = .HDAvailablePercent
        HDTotalBytesUsed = .HDTotalBytesUsed
        
    End With

End Sub

