Attribute VB_Name = "m00_Excel64_WindowManagement"
'===============================================================
' Module: frmWindowManagement (100% Autonomous - No UserForm Code Needed)
' Purpose: Zero-touch window management - just call ShowManagedForm()
' Updated: Fully independent, automatic cleanup, no form code required
'===============================================================
Option Explicit

'=====================================================================
' Required API declarations (64-bit safe) - add only if not already present
'=====================================================================
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0   ' Primary display width
Private Const SM_CYSCREEN As Long = 1   ' Primary display height

' ---------------- Windows API Declarations (64-bit safe) ----------------
#If VBA7 Then
    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    
    #If Win64 Then
        Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
            (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
            (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
#Else
    ' 32-bit declarations
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Public Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    
    Public Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetForegroundWindow Lib "user32" () As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#End If

' Constants
#If VBA7 Then
    Public Const HWND_TOPMOST As LongPtr = -1
    Public Const HWND_NOTOPMOST As LongPtr = -2
#Else
    Public Const HWND_TOPMOST As Long = -1
    Public Const HWND_NOTOPMOST As Long = -2
#End If

Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_NOACTIVATE As Long = &H10
Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_STYLE As Long = -16
Public Const WS_EX_TOPMOST As Long = &H8
Public Const WS_EX_TOOLWINDOW As Long = &H80
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_SYSMENU As Long = &H80000
Public Const SW_SHOW As Long = 5
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOWNOACTIVATE As Long = 4
Public Const SW_MINIMIZE As Long = 6

' ---------------- Private State Management ----------------
Private Type FormSession
    formObject As Object
    FormHwnd As LongPtr
    formCaption As String
    TopmostEnabled As Boolean
    MinimizedExcel As Boolean
    TimerScheduledTime As Date
    WatcherScheduledTime As Date
    CleanupDone As Boolean
End Type

Private gSession As FormSession

' Excel state tracking
Private Type ExcelState
    hWnd As LongPtr
    MinimizedByUs As Boolean
    ShrunkByUs As Boolean
    RectSaved As Boolean
    SavedLeft As Long
    SavedTop As Long
    SavedWidth As Long
    SavedHeight As Long
End Type

Private gExcelState As ExcelState

' Global dictionary (optional - keep if needed for other features)
Public gInputListStore As Object

' ---------------- PUBLIC INTERFACE: Flexible parameters (supports Boolean, Long, Variant) ----------------

Public Sub ShowManagedForm(formObject As Object, _
    Optional formCaption As String = "", _
    Optional minimizeExcel As Variant = True, _
    Optional enableTopmost As Variant = True)
    
    On Error GoTo ErrorHandler
    
    ' Convert Variant parameters to Boolean (supports 1/0, True/False, Excel values)
    Dim bMinimizeExcel As Boolean
    Dim bEnableTopmost As Boolean
    
    bMinimizeExcel = ConvertToBool(minimizeExcel)
    bEnableTopmost = ConvertToBool(enableTopmost)
    
    ' Initialize dictionary if needed
    If gInputListStore Is Nothing Then
        Set gInputListStore = CreateObject("Scripting.Dictionary")
    End If
    
    ' Cleanup any previous session
    If Not gSession.CleanupDone Then
        PerformCleanup
    End If
    
    ' Reset session
    With gSession
        Set .formObject = formObject
        .FormHwnd = 0
        .TopmostEnabled = bEnableTopmost
        .MinimizedExcel = bMinimizeExcel
        .CleanupDone = False
        
        ' Set caption
        If formCaption <> "" Then
            formObject.Caption = formCaption
            .formCaption = formCaption
        Else
            .formCaption = formObject.Caption
        End If
    End With
    
    ' Show form modeless
    formObject.Show vbModeless
    Sleep 100
    DoEvents
    
    ' Get window handle and apply management
    gSession.FormHwnd = GetFormHandle(gSession.formCaption)
    If gSession.FormHwnd = 0 Then
        MsgBox "Warning: Could not get form window handle. Management features disabled.", vbExclamation
        Exit Sub
    End If
    
    ' Apply window styles
    ApplyWindowStyles gSession.FormHwnd
    PositionFormLeftThird gSession.FormHwnd
    
    ' Handle Excel window
    If bMinimizeExcel Then
        MinimizeExcelWindow
    End If
    
    ' Handle topmost
    If bEnableTopmost Then
        MakeWindowTopmost gSession.FormHwnd
        StartTopmostTimer
    Else
        ' Explicitly make NOT topmost
        ClearWindowTopmostFlag gSession.FormHwnd
    End If
    
    ' Start autonomous watcher (checks if form still exists)
    StartFormWatcher
    
    Exit Sub

ErrorHandler:
    MsgBox "Error showing managed form: " & err.description, vbCritical
    PerformCleanup
End Sub

' ---------------- HELPER: Convert various formats to Boolean ----------------
' Supports: True/False, 1/0, -1/0, "True"/"False", "1"/"0", Excel cell values
Private Function ConvertToBool(ByVal Value As Variant) As Boolean
    On Error Resume Next
    
    Select Case VarType(Value)
        Case vbBoolean
            ' Already Boolean
            ConvertToBool = Value
            
        Case vbInteger, vbLong, vbByte, vbSingle, vbDouble, vbCurrency
            ' Numeric: any non-zero = True
            ConvertToBool = (Value <> 0)
            
        Case vbString
            ' String: "1", "true", "yes", "on" = True (case insensitive)
            Dim s As String
            s = LCase(Trim(CStr(Value)))
            ConvertToBool = (s = "1" Or s = "true" Or s = "yes" Or s = "on" Or s = "-1")
            
        Case vbEmpty, vbNull
            ' Empty/Null = False
            ConvertToBool = False
            
        Case Else
            ' Try to convert to numeric
            Dim n As Long
            n = CLng(Value)
            ConvertToBool = (n <> 0)
    End Select
    
    ' If error occurred, default to False
    If err.number <> 0 Then
        err.Clear
        ConvertToBool = False
    End If
End Function

' ---------------- BACKWARD COMPATIBILITY ----------------


Public Sub ShowFormWithWindowManagement(formObject As Object, _
    Optional topmostMode As Long = 1, _
    Optional formCaption As String = "", _
    Optional minimizeExcel As Long = 1)
    
    ' Convert old parameters to new format and call the new function
    ShowManagedForm formObject, formCaption, (minimizeExcel <> 0), (topmostMode <> 0)
End Sub

' ---------------- AUTOMATIC CLEANUP (No UserForm Code Needed) ----------------
Private Sub StartFormWatcher()
    On Error Resume Next
    gSession.WatcherScheduledTime = Now + TimeValue("00:00:01")
    Application.OnTime gSession.WatcherScheduledTime, "CheckFormStillExists"
End Sub

Public Sub CheckFormStillExists()
    On Error Resume Next
    
    ' Exit if already cleaned up
    If gSession.CleanupDone Then Exit Sub
    
    ' Check if form object is still valid
    If gSession.formObject Is Nothing Then
        PerformCleanup
        Exit Sub
    End If
    
    ' Check if window handle is still valid
    If gSession.FormHwnd <> 0 Then
        If IsWindow(gSession.FormHwnd) = 0 Then
            ' Window destroyed - form was closed
            PerformCleanup
            Exit Sub
        End If
    End If
    
    ' Check if form is still in UserForms collection
    Dim uf As Object
    Dim found As Boolean
    found = False
    
    For Each uf In VBA.UserForms
        If uf Is gSession.formObject Then
            found = True
            Exit For
        End If
    Next uf
    
    If Not found Then
        ' Form was unloaded
        PerformCleanup
        Exit Sub
    End If
    
    ' Form still exists - schedule next check
    gSession.WatcherScheduledTime = Now + TimeValue("00:00:01")
    Application.OnTime gSession.WatcherScheduledTime, "CheckFormStillExists"
End Sub

Private Sub StopFormWatcher()
    On Error Resume Next
    If gSession.WatcherScheduledTime <> 0 Then
        Application.OnTime EarliestTime:=gSession.WatcherScheduledTime, _
            Procedure:="CheckFormStillExists", Schedule:=False
        gSession.WatcherScheduledTime = 0
    End If
End Sub

Private Sub PerformCleanup()
    On Error Resume Next
    
    ' Prevent double cleanup
    If gSession.CleanupDone Then Exit Sub
    gSession.CleanupDone = True
    
    ' Stop all timers
    StopFormWatcher
    StopTopmostTimer
    
    ' Remove topmost flag if window still exists
    If gSession.FormHwnd <> 0 Then
        If IsWindow(gSession.FormHwnd) <> 0 Then
            ClearWindowTopmostFlag gSession.FormHwnd
        End If
    End If
    
    ' Restore Excel if we minimized/shrunk it
    If gSession.MinimizedExcel Then
        RestoreExcelWindow
    End If
    
    ' Clear session
    Set gSession.formObject = Nothing
    gSession.FormHwnd = 0
End Sub

' ---------------- Topmost Timer (Only if Enabled) ----------------
Private Sub StartTopmostTimer()
    On Error Resume Next
    If gSession.CleanupDone Then Exit Sub
    
    gSession.TimerScheduledTime = Now + TimeValue("00:00:01")
    Application.OnTime gSession.TimerScheduledTime, "MaintainTopmostState"
End Sub

Public Sub MaintainTopmostState()
    On Error Resume Next
    
    ' Exit if cleaned up or topmost not enabled
    If gSession.CleanupDone Or Not gSession.TopmostEnabled Then Exit Sub
    
    ' Validate window still exists
    If gSession.FormHwnd = 0 Or IsWindow(gSession.FormHwnd) = 0 Then
        StopTopmostTimer
        Exit Sub
    End If
    
    ' Restore if minimized
    If IsIconic(gSession.FormHwnd) <> 0 Then
        ShowWindow gSession.FormHwnd, SW_RESTORE
        BringWindowToTop gSession.FormHwnd
        SetForegroundWindow gSession.FormHwnd
    End If
    
    ' Re-apply topmost if not foreground
    Dim currentForeground As LongPtr
    currentForeground = GetForegroundWindow()
    
    If currentForeground <> gSession.FormHwnd Then
        SetWindowPos gSession.FormHwnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    End If
    
    ' Schedule next check
    gSession.TimerScheduledTime = Now + TimeValue("00:00:01")
    Application.OnTime gSession.TimerScheduledTime, "MaintainTopmostState"
End Sub

Private Sub StopTopmostTimer()
    On Error Resume Next
    If gSession.TimerScheduledTime <> 0 Then
        Application.OnTime EarliestTime:=gSession.TimerScheduledTime, _
            Procedure:="MaintainTopmostState", Schedule:=False
        gSession.TimerScheduledTime = 0
    End If
End Sub

' ---------------- Core Window Management ----------------
Private Function GetFormHandle(Caption As String) As LongPtr
    Dim hWnd As LongPtr
    
    If val(Application.Version) >= 9 Then
        hWnd = FindWindow("ThunderDFrame", Caption)
    Else
        hWnd = FindWindow("ThunderXFrame", Caption)
    End If
    
    If hWnd = 0 Then hWnd = FindWindow(vbNullString, Caption)
    GetFormHandle = hWnd
End Function

Private Sub ApplyWindowStyles(hWnd As LongPtr)
    On Error Resume Next
    Dim lStyle As LongPtr
    
    lStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_APPWINDOW
    lStyle = lStyle And Not WS_EX_TOOLWINDOW
    
    SetWindowLongPtr hWnd, GWL_EXSTYLE, lStyle
End Sub

Private Sub MakeWindowTopmost(hWnd As LongPtr)
    On Error Resume Next
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    ShowWindow hWnd, SW_SHOW
    BringWindowToTop hWnd
    SetForegroundWindow hWnd
End Sub

Private Sub ClearWindowTopmostFlag(ByVal hWnd As LongPtr)
    On Error Resume Next
    If hWnd = 0 Then Exit Sub
    
    Dim exStyle As LongPtr
    exStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
    exStyle = exStyle And Not WS_EX_TOPMOST
    SetWindowLongPtr hWnd, GWL_EXSTYLE, exStyle
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

' ---------------- Excel Window Management ----------------
Private Function GetExcelMainHwnd() As LongPtr
    On Error Resume Next
    Dim h As LongPtr
    
    h = Application.hWnd
    If err.number <> 0 Or h = 0 Then
        err.Clear
        h = FindWindow("XLMAIN", vbNullString)
    End If
    
    If h = 0 Then h = FindWindow(vbNullString, Application.Caption)
    GetExcelMainHwnd = h
End Function

Private Sub MinimizeExcelWindow()
    On Error Resume Next
    Dim h As LongPtr
    
    h = GetExcelMainHwnd()
    If h = 0 Then Exit Sub
    
    gExcelState.hWnd = h
    
    ' Don't minimize if already minimized (not our responsibility)
    If IsIconic(h) <> 0 Then
        gExcelState.MinimizedByUs = False
        Exit Sub
    End If
    
    ' Save rect before minimizing (for optional restore)
    Dim rc As RECT
    If GetWindowRect(h, rc) <> 0 Then
        With gExcelState
            .SavedLeft = rc.Left
            .SavedTop = rc.Top
            .SavedWidth = rc.Right - rc.Left
            .SavedHeight = rc.Bottom - rc.Top
            .RectSaved = True
        End With
    End If
    
    ShowWindow h, SW_MINIMIZE
    gExcelState.MinimizedByUs = True
End Sub

Private Sub RestoreExcelWindow()
    On Error Resume Next
    Dim h As LongPtr
    
    h = IIf(gExcelState.hWnd <> 0, gExcelState.hWnd, GetExcelMainHwnd())
    If h = 0 Then Exit Sub
    
    ' Only restore if we minimized it
    If gExcelState.MinimizedByUs Then
        ShowWindow h, SW_RESTORE
        gExcelState.MinimizedByUs = False
    End If
    
    ' Also restore shrunk window if applicable
    If gExcelState.ShrunkByUs Then
        If gExcelState.RectSaved Then
            MoveWindow h, gExcelState.SavedLeft, gExcelState.SavedTop, _
                gExcelState.SavedWidth, gExcelState.SavedHeight, 1
        End If
        gExcelState.ShrunkByUs = False
        gExcelState.RectSaved = False
    End If
End Sub

' ---------------- Optional: Manual Control Functions ----------------
' These are provided for advanced scenarios but NOT required for normal use

Public Sub ManualCleanup()
    ' Call this ONLY if you need to force cleanup before form closes
    ' Normal usage doesn't need this - cleanup is automatic
    PerformCleanup
End Sub

Public Sub ToggleExcelShrink()
    ' Alternative to minimize: shrink Excel to 1x1 pixel
    On Error Resume Next
    Dim h As LongPtr
    
    h = GetExcelMainHwnd()
    If h = 0 Then Exit Sub
    
    If gExcelState.ShrunkByUs Then
        ' Restore
        If gExcelState.RectSaved Then
            MoveWindow h, gExcelState.SavedLeft, gExcelState.SavedTop, _
                gExcelState.SavedWidth, gExcelState.SavedHeight, 1
            gExcelState.ShrunkByUs = False
            gExcelState.RectSaved = False
        End If
    Else
        ' Shrink
        If IsIconic(h) <> 0 Then Exit Sub ' Don't shrink if minimized
        
        ' Save rect
        Dim rc As RECT
        If GetWindowRect(h, rc) <> 0 Then
            With gExcelState
                .SavedLeft = rc.Left
                .SavedTop = rc.Top
                .SavedWidth = rc.Right - rc.Left
                .SavedHeight = rc.Bottom - rc.Top
                .RectSaved = True
                .hWnd = h
            End With
        End If
        
        MoveWindow h, 0, 0, 1, 1, 1
        gExcelState.ShrunkByUs = True
    End If
End Sub

Public Sub AddWindowButtons(Optional minimizeBox As Boolean = True, Optional maximizeBox As Boolean = False)
    ' Add minimize/maximize buttons to the managed form
    ' Call AFTER ShowManagedForm
    If gSession.FormHwnd = 0 Then Exit Sub
    
    Dim lStyle As LongPtr
    lStyle = GetWindowLongPtr(gSession.FormHwnd, GWL_STYLE)
    
    If minimizeBox Then lStyle = lStyle Or WS_MINIMIZEBOX
    If maximizeBox Then lStyle = lStyle Or WS_MAXIMIZEBOX
    
    lStyle = lStyle Or WS_SYSMENU
    
    SetWindowLongPtr gSession.FormHwnd, GWL_STYLE, lStyle
    DrawMenuBar gSession.FormHwnd
End Sub

'=====================================================================
' Position managed UserForm in the middle of the left 1/3 of the screen
' Fully automatic - no changes required on individual UserForms
'=====================================================================
Private Sub PositionFormLeftThird(ByVal hWnd As LongPtr)
    On Error Resume Next
    
    Const LEFT_SCREEN_RATIO As Double = 1# / 1#          ' Use left 1/3 of primary monitor
    Const HORIZONTAL_MARGIN As Long = 80                ' Distance from left edge (pixels)
    Const CENTER_VERTICALLY As Boolean = True           ' True = vertically centered
    
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim formWidth As Long
    Dim formHeight As Long
    Dim newX As Long
    Dim newY As Long
    
    Dim rcForm As RECT
    
    ' Get primary monitor dimensions
    screenWidth = GetSystemMetrics(SM_CXSCREEN)
    screenHeight = GetSystemMetrics(SM_CYSCREEN)
    
    ' Get current form size
    If GetWindowRect(hWnd, rcForm) = 0 Then Exit Sub
    
    formWidth = rcForm.Right - rcForm.Left
    formHeight = rcForm.Bottom - rcForm.Top
    
    ' Calculate horizontal position within the left 1/3 zone
    Dim zoneWidth As Long
    zoneWidth = CLng(screenWidth * LEFT_SCREEN_RATIO) - (HORIZONTAL_MARGIN * 2)
    
    newX = HORIZONTAL_MARGIN + (zoneWidth - formWidth) \ 2
    If newX < HORIZONTAL_MARGIN Then newX = HORIZONTAL_MARGIN
    
    ' Vertical position
    If CENTER_VERTICALLY Then
        newY = (screenHeight - formHeight) \ 2
    Else
        newY = 100  ' Fixed offset from top if not centered
    End If
    
    ' Safety bounds
    If newX < 0 Then newX = 0
    If newY < 0 Then newY = 0
    
    ' Move and repaint
    MoveWindow hWnd, newX, newY, formWidth, formHeight, 1
    
    ' Ensure visibility
    BringWindowToTop hWnd
End Sub


