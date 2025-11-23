VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSelector 
   Caption         =   "SAP2000 Database Tables (by thanhtdvncc)"
   ClientHeight    =   4949
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   4665
   OleObjectBlob   =   "frmTableSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTableSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' UserForm: frmTableSelector
' Add this as a new UserForm in your VBA project
'===============================================================
' Instructions to create the UserForm:
' 1. In VBA Editor, Insert > UserForm
' 2. Name it: frmTableSelector
' 3. Add controls:
'    - Label: lblTitle (caption: "Select Database Table to Export")
'    - ListBox: lstTables
'    - TextBox: txtSearch (for filtering)
'    - Label: lblSearch (caption: "Search:")
'    - Button: btnExport (caption: "Export to Excel")
'    - Button: btnCancel (caption: "Cancel")
' 4. Paste the code below into the UserForm code module

'--- START OF USERFORM CODE (Paste into frmTableSelector) ---
' UserForm Code for frmTableSelector
Option Explicit

Private m_TableKeys() As String
Private m_TableNames() As String
Private m_ImportTypes() As Long
Private m_NumberTables As Long

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    Me.Caption = "SAP2000 Database Tables (by thanhtdvncc)"
    Me.Width = 450
    Me.height = 400
    
    ' Setup search
    lblSearch.Caption = "Search:"
    lblSearch.Left = 10
    lblSearch.Top = 10
    
    txtSearch.Left = 60
    txtSearch.Top = 10
    txtSearch.Width = 360
    
    ' Setup title
    lblTitle.Caption = "Select Database Table to Export:"
    lblTitle.Left = 10
    lblTitle.Top = 40
    lblTitle.Width = 400
    lblTitle.Font.Bold = True
    
    ' Setup listbox
    lstTables.Left = 10
    lstTables.Top = 65
    lstTables.Width = 410
    lstTables.height = 240
    
    ' Setup buttons
    btnExport.Caption = "Export to Excel"
    btnExport.Left = 180
    btnExport.Top = 315
    btnExport.Width = 120
    btnExport.height = 30
    btnExport.Default = True
    
    btnCancel.Caption = "Cancel"
    btnCancel.Left = 310
    btnCancel.Top = 315
    btnCancel.Width = 110
    btnCancel.height = 30
    btnCancel.Cancel = True
    
    ' Load tables
    Call LoadTables
    
    Exit Sub
ErrHandler:
    MsgBox "Initialization error: " & err.description, vbCritical
End Sub

Private Sub LoadTables()
    On Error GoTo ErrHandler
    
    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Dim ret As Long
    ret = SapModel.DatabaseTables.GetAvailableTables( _
        m_NumberTables, m_TableKeys, m_TableNames, m_ImportTypes)
    
    If ret <> 0 Or m_NumberTables = 0 Then
        MsgBox "No database tables available.", vbInformation
        Unload Me
        Exit Sub
    End If
    
    Call FilterTables("")
    
    Exit Sub
ErrHandler:
    MsgBox "Error loading tables: " & err.description, vbCritical
End Sub

Private Sub FilterTables(searchText As String)
    lstTables.Clear
    
    Dim i As Long
    Dim displayText As String
    searchText = LCase(Trim(searchText))
    
    For i = 0 To m_NumberTables - 1
        displayText = m_TableNames(i)
        
        If searchText = "" Or InStr(1, LCase(displayText), searchText) > 0 Then
            lstTables.AddItem displayText
            lstTables.List(lstTables.ListCount - 1, 1) = CStr(i) ' Store index
        End If
    Next i
    
    If lstTables.ListCount > 0 Then
        lstTables.ListIndex = 0
    End If
End Sub

Private Sub txtSearch_Change()
    Call FilterTables(txtSearch.Value)
End Sub

Private Sub lstTables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnExport_Click
End Sub

Private Sub btnExport_Click()
    If lstTables.ListIndex < 0 Then
        MsgBox "Please select a table.", vbExclamation
        Exit Sub
    End If
    
    Dim idx As Long
    idx = CLng(lstTables.List(lstTables.ListIndex, 1))
    
    Dim selectedKey As String
    selectedKey = m_TableKeys(idx)
    
    Me.Hide
    Call ExportTableToActiveSheet(selectedKey)
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
'--- END OF USERFORM CODE ---



