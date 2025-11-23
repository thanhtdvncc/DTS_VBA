Attribute VB_Name = "LibDTS_Global"
' Module: LibDTS_Global
Option Explicit

' --- SYSTEM CONSTANTS ---
Public Const DTS_APP_NAME As String = "DTS_CORE_DATA"
Public Const DTS_VERSION As String = "2.0.0"
Public Const DTS_PRECISION As Double = 0.0001
Public Const DTS_CONFIG_FILENAME As String = "settings.json"
Public Const DTS_XDATA_APPNAME As String = "DTS_CORE"

' --- GLOBAL SINGLETONS ---
' Bien toan cuc luu tru cau hinh va logger
Private p_Config As clsDTSConfig
Private p_Logger As Object ' Late binding for Logger

' --- ENUMERATIONS (Type Safety) ---

' Element Types
' Cac loai phan tu
Public Enum DTSElementType
    DTS_ELEM_UNKNOWN = 0
    DTS_ELEM_FRAME = 1      ' Beams, Columns
    DTS_ELEM_AREA = 2       ' Slabs, Walls
    DTS_ELEM_NODE = 3       ' Points
    DTS_ELEM_ANNOTATION = 4 ' Tags, Dims
    DTS_ELEM_REBAR = 5      ' Rebars
End Enum

' Frame Types
' Cac loai thanh
Public Enum DTSFrameType
    DTS_FRM_BEAM = 1
    DTS_FRM_COLUMN = 2
    DTS_FRM_BRACE = 3
    DTS_FRM_PILE = 4
End Enum

' Shape Types (For Section)
' Cac loai hinh dang (Cho tiet dien)
Public Enum DTSShapeType
    DTS_SHP_RECTANGLE = 1
    DTS_SHP_CIRCLE = 2
    DTS_SHP_I_SECTION = 3
    DTS_SHP_T_SECTION = 4
    DTS_SHP_L_SECTION = 5
End Enum

' Rebar Shapes (Standard Codes)
' Ma hinh dang thep (Tieu chuan)
Public Enum DTSRebarShape
    DTS_RBR_00 = 0  ' Unknown
    DTS_RBR_01 = 1  ' Straight
    DTS_RBR_02 = 2  ' Straight with hooks
    DTS_RBR_18 = 18 ' U-Shape
    DTS_RBR_51 = 51 ' Stirrup
End Enum

' --- GLOBAL ACCESSORS (LAZY LOADING) ---

' Get the shared Config instance
' Lay instance cau hinh dung chung
Public Function Config() As clsDTSConfig
    If p_Config Is Nothing Then
        Set p_Config = New clsDTSConfig
        ' Auto-load settings from file on first access
        ' Tu dong tai cai dat tu file khi truy cap lan dau
        p_Config.Reload
    End If
    Set Config = p_Config
End Function

' Get the shared Logger instance
' Lay instance Logger dung chung
Public Function Logger() As Object
    If p_Logger Is Nothing Then
        ' Assume LibDTS_Logger is a Standard Module, but if wrapping in Class:
        ' Gia su dung Module, neu dung Class thi khoi tao o day
        ' Set p_Logger = New clsDTSLogger
    End If
    Set Logger = p_Logger
End Function

' Reset Global State (Call when error occurs or app restarts)
' Dat lai trang thai toan cuc (Goi khi co loi hoac khoi dong lai)
Public Sub ResetGlobals()
    Set p_Config = Nothing
    Set p_Logger = Nothing
End Sub
