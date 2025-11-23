VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSConfig
Option Explicit

' Storage for settings
' Kho chua cai dat
Private p_Settings As Object
Private p_IsDirty As Boolean

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    Set p_Settings = CreateObject("Scripting.Dictionary")
    p_IsDirty = False
End Sub

' --- METHODS ---

' Load settings from disk (via Driver)
' Tai cai dat tu dia (thong qua Driver)
Public Sub Reload()
    Dim loadedSettings As Object
    Set loadedSettings = LibDTS_DriverDB.LoadSettings()
    
    If loadedSettings.count > 0 Then
        Set p_Settings = loadedSettings
    Else
        ' If file not found, load defaults
        ' Neu khong tim thay file, tai mac dinh
        LoadDefaults
    End If
    p_IsDirty = False
End Sub

' Save settings to disk (via Driver)
' Luu cai dat xuong dia (thong qua Driver)
Public Sub Save()
    If p_IsDirty Then
        LibDTS_DriverDB.SaveSettings p_Settings
        p_IsDirty = False
    End If
End Sub

' Set default values
' Thiet lap gia tri mac dinh
Private Sub LoadDefaults()
    ' Layers
    p_Settings("Layer_Beam") = "DTS_BEAM"
    p_Settings("Layer_Column") = "DTS_COLUMN"
    p_Settings("Layer_Wall") = "DTS_WALL"
    p_Settings("Layer_Slab") = "DTS_SLAB"
    p_Settings("Layer_Rebar_Main") = "DTS_REBAR_MAIN"
    p_Settings("Layer_Rebar_Stirrup") = "DTS_REBAR_STIR"
    p_Settings("Layer_Dim") = "DTS_DIM"
    p_Settings("Layer_Tag") = "DTS_TAG"
    
    ' Colors (ACI)
    p_Settings("Color_Beam") = 2    ' Yellow
    p_Settings("Color_Column") = 2  ' Yellow
    p_Settings("Color_Wall") = 2    ' Yellow
    p_Settings("Color_Rebar") = 6   ' Magenta
    
    ' System
    p_Settings("Unit_System") = "Metric"
    p_Settings("Language") = "EN"
    
    p_IsDirty = True
End Sub

' --- ACCESSORS ---

Public Function GetVal(key As String, Optional defaultVal As Variant = "") As Variant
    If p_Settings.exists(key) Then
        GetVal = p_Settings(key)
    Else
        GetVal = defaultVal
    End If
End Function

Public Sub SetVal(key As String, val As Variant)
    If p_Settings.exists(key) Then
        If p_Settings(key) <> val Then
            p_Settings(key) = val
            p_IsDirty = True
        End If
    Else
        p_Settings.Add key, val
        p_IsDirty = True
    End If
End Sub
