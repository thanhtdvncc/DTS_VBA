Attribute VB_Name = "Core_Utils"
' ===============================================================
' Module: Core_Utils
' Purpose: Core logic for SAP2000 element filtering and selection
' Notes:
' - Provides public functions for getting element lists by criteria
' - Can be called from UserForm or other modules
' - All logic extracted from frmSelection for reusability
' ===============================================================
Option Explicit

' Gridline map cache (shared)
Private m_GridlineMap As Object

' ===============================================================
' PUBLIC API FUNCTIONS
' ===============================================================

' Function 1: Get element list by criteria
' Parameters:
'   elementType: "Node", "Frame", "Cable", "Tendon", "Area", "Solid", "Link"
'   attributeName: "Name", "GUID", "Group", "Section", "Material", "Constraint", "Support", "Spring", "Design-Type", "AutoSection", "Area-Type", "Property", "Coordinate-X/Y/Z"
'   propertyValues(): Array of property values to match (e.g., Array("FSEC1", "FSEC2"))
'   Optional coordinateRanges: X1, X2, Y1, Y2, Z1, Z2 for coordinate filtering or range coordinate filtering
'   Optional includeGroup: True/False to include group filtering
' Returns: Collection of element names
Public Function DTS_SAP2000_Getlist( _
    ByVal ElementType As String, _
    ByVal attributeName As String, _
    ByRef propertyValues() As Variant, _
    Optional ByVal x1 As Variant = Empty, _
    Optional ByVal x2 As Variant = Empty, _
    Optional ByVal y1 As Variant = Empty, _
    Optional ByVal y2 As Variant = Empty, _
    Optional ByVal z1 As Variant = Empty, _
    Optional ByVal z2 As Variant = Empty, _
    Optional ByVal includeGroup As Boolean = False _
) As Collection

    Dim result As New Collection
    
    ' Validate SAP connection
    If Not EnsureSapModelAvailable() Then
        Set DTS_SAP2000_Getlist = result
        Exit Function
    End If
    
    ' Build property values collection from array
    Dim propCol As New Collection
    Dim i As Long
    On Error Resume Next
    For i = LBound(propertyValues) To UBound(propertyValues)
        If Trim$(CStr(propertyValues(i))) <> "" Then
            propCol.Add CStr(propertyValues(i))
        End If
    Next i
    On Error GoTo 0
    
    If propCol.count = 0 Then
        Set DTS_SAP2000_Getlist = result
        Exit Function
    End If
    
    ' Get elements by union of property values
    Dim unionColl As New Collection
    Dim pv As Variant
    
    For Each pv In propCol
        Dim tmpColl As Collection
        Set tmpColl = GetElementsByPropertyValue_Core(attributeName, CStr(pv), ElementType)
        
        Dim it As Variant
        For Each it In tmpColl
            If Not IsInCollection_Core(unionColl, CStr(it)) Then
                unionColl.Add it
            End If
        Next it
    Next pv
    
    ' Apply coordinate filtering if provided
    If Not IsEmpty(x1) And Not IsEmpty(x2) Then
        Set unionColl = FilterByCoordinateRange_Core(unionColl, ElementType, "X", CDbl(x1), CDbl(x2))
    End If
    If Not IsEmpty(y1) And Not IsEmpty(y2) Then
        Set unionColl = FilterByCoordinateRange_Core(unionColl, ElementType, "Y", CDbl(y1), CDbl(y2))
    End If
    If Not IsEmpty(z1) And Not IsEmpty(z2) Then
        Set unionColl = FilterByCoordinateRange_Core(unionColl, ElementType, "Z", CDbl(z1), CDbl(z2))
    End If
    
    Set DTS_SAP2000_Getlist = unionColl
End Function

' ===============================================================
' DTS_SAP2000_Getlist2node
' ===============================================================
' Purpose:
'   Advanced version of DTS_SAP2000_Getlist with element->node mapping

' Parameters:
'   returnMode (Long):
'       1 = Return elements only (Collection)
'       2 = Return nodes only (Collection)
'       3 = Return mapping Dictionary (Key=ElementName, Value=Collection of NodeNames)

'   elementType: "Node", "Frame", "Cable", "Tendon", "Area", "Solid", "Link"
'   attributeName: "Name", "GUID", "Group", "Section", "Material", "Constraint", "Support", "Spring", "Design-Type", "AutoSection", "Area-Type", "Property", "Coordinate-X/Y/Z"
'   propertyValues(): Array of property values to match (e.g., Array("FSEC1", "FSEC2"))
'   Optional coordinateRanges: X1, X2, Y1, Y2, Z1, Z2 for coordinate filtering or range coordinate filtering

Public Function DTS_SAP2000_Getlist2node( _
    ByVal ElementType As String, _
    ByVal attributeName As String, _
    ByRef propertyValues() As Variant, _
    Optional ByVal x1 As Variant = Empty, _
    Optional ByVal x2 As Variant = Empty, _
    Optional ByVal y1 As Variant = Empty, _
    Optional ByVal y2 As Variant = Empty, _
    Optional ByVal z1 As Variant = Empty, _
    Optional ByVal z2 As Variant = Empty, _
    Optional ByVal returnMode As Long = 1 _
) As Object

    On Error GoTo ErrHandler

    Dim retObj As Object
    Dim elems As Collection
    Dim el As Variant
    Dim elName As String
    Dim elemNodes As Collection
    Dim n As Variant
    Dim nodeName As String

    ' Validate returnMode
    If returnMode < 1 Or returnMode > 3 Then returnMode = 1

    ' Get filtered elements using existing API
    Set elems = DTS_SAP2000_Getlist(ElementType, attributeName, propertyValues, x1, x2, y1, y2, z1, z2)

    ' Prepare empty return based on mode when no results or error
    If elems Is Nothing Or elems.count = 0 Then
        If returnMode = 3 Then
            Set retObj = CreateObject("Scripting.Dictionary")
        Else
            Set retObj = New Collection
        End If
        Set DTS_SAP2000_Getlist2node = retObj
        Exit Function
    End If

    Select Case returnMode
        Case 1
            ' Return elements only (Collection)
            Dim colElems As Collection
            Set colElems = New Collection
            For Each el In elems
                colElems.Add CStr(el)
            Next el
            Set DTS_SAP2000_Getlist2node = colElems
            Exit Function

        Case 2
            ' Return unique nodes only (Collection)
            Dim uniqueNodeTracker As Object
            Set uniqueNodeTracker = CreateObject("Scripting.Dictionary")
            Dim colNodes As Collection
            Set colNodes = New Collection

            For Each el In elems
                elName = CStr(el)
                Set elemNodes = GetNodesFromElement_Core(elName, ElementType)
                If Not elemNodes Is Nothing Then
                    For Each n In elemNodes
                        nodeName = CStr(n)
                        If Not uniqueNodeTracker.exists(nodeName) Then
                            uniqueNodeTracker.Add nodeName, True
                            colNodes.Add nodeName
                        End If
                    Next n
                End If
            Next el

            Set DTS_SAP2000_Getlist2node = colNodes
            Exit Function

        Case 3
            ' Return mapping: Dictionary(ElementName -> Collection of NodeNames)
            Dim mapDict As Object
            Set mapDict = CreateObject("Scripting.Dictionary")

            For Each el In elems
                elName = CStr(el)
                Set elemNodes = GetNodesFromElement_Core(elName, ElementType)

                ' Ensure we always store a Collection object (can be empty)
                If elemNodes Is Nothing Then
                    Dim emptyColl As Collection
                    Set emptyColl = New Collection
                    mapDict.Add elName, emptyColl
                Else
                    mapDict.Add elName, elemNodes
                End If
            Next el

            Set DTS_SAP2000_Getlist2node = mapDict
            Exit Function
    End Select

    ' Fallback
    If returnMode = 3 Then
        Set retObj = CreateObject("Scripting.Dictionary")
    Else
        Set retObj = New Collection
    End If
    Set DTS_SAP2000_Getlist2node = retObj
    Exit Function

ErrHandler:
    LogError "DTS_SAP2000_Getlist2node", err.number, err.description
    If returnMode = 3 Then
        Set retObj = CreateObject("Scripting.Dictionary")
    Else
        Set retObj = New Collection
    End If
    Set DTS_SAP2000_Getlist2node = retObj
    On Error GoTo 0
End Function

' ===============================================================
' MAPPING UTILITIES - Add to end of module
' ===============================================================

' Helper: Get nodes from element mapping (safe accessor)
' Usage: Set nodes = GetNodesFromMapping(mapping, "Frame123")
Public Function GetNodesFromMapping(ByVal mapping As Object, ByVal elemName As String) As Collection
    Dim result As New Collection
    
    On Error Resume Next
    If mapping.exists(elemName) Then
        Set result = mapping(elemName)
    End If
    On Error GoTo 0
    
    Set GetNodesFromMapping = result
End Function

' Helper: Check if element has specific node
' Usage: If ElementHasNode(mapping, "Frame123", "Node45") Then...
Public Function ElementHasNode(ByVal mapping As Object, ByVal elemName As String, ByVal nodeName As String) As Boolean
    ElementHasNode = False
    
    On Error Resume Next
    If mapping.exists(elemName) Then
        Dim nodes As Collection
        Set nodes = mapping(elemName)
        
        Dim n As Variant
        For Each n In nodes
            If CStr(n) = nodeName Then
                ElementHasNode = True
                Exit Function
            End If
        Next n
    End If
    On Error GoTo 0
End Function

' Helper: Get all elements connected to a specific node (reverse lookup)
' Usage: Set frames = GetElementsAtNode(mapping, "Node45")
Public Function GetElementsAtNode(ByVal mapping As Object, ByVal targetNode As String) As Collection
    Dim result As New Collection
    
    On Error Resume Next
    Dim elemName As Variant
    Dim nodes As Collection
    Dim n As Variant
    
    For Each elemName In mapping.keys
        Set nodes = mapping(elemName)
        
        For Each n In nodes
            If CStr(n) = targetNode Then
                result.Add CStr(elemName)
                Exit For
            End If
        Next n
    Next elemName
    
    On Error GoTo 0
    Set GetElementsAtNode = result
End Function

' Helper: Count nodes per element
' Usage: nodeCount = CountNodesInElement(mapping, "Frame123")
Public Function CountNodesInElement(ByVal mapping As Object, ByVal elemName As String) As Long
    CountNodesInElement = 0
    
    On Error Resume Next
    If mapping.exists(elemName) Then
        Dim nodes As Collection
        Set nodes = mapping(elemName)
        CountNodesInElement = nodes.count
    End If
    On Error GoTo 0
End Function

' Helper: Print mapping to Immediate Window (Debug)
' Usage: PrintMapping mapping, "Frames connected to nodes:"
Public Sub PrintMapping(ByVal mapping As Object, Optional ByVal title As String = "Element-Node Mapping")
    If mapping Is Nothing Then
        Debug.Print "[Mapping is Nothing]"
        Exit Sub
    End If
    
    Debug.Print String(60, "=")
    Debug.Print title
    Debug.Print "Total elements: " & mapping.count
    Debug.Print String(60, "-")
    
    Dim elemName As Variant
    Dim nodes As Collection
    Dim n As Variant
    Dim nodeList As String
    
    For Each elemName In mapping.keys
        Set nodes = mapping(elemName)
        
        nodeList = ""
        For Each n In nodes
            If nodeList <> "" Then nodeList = nodeList & ", "
            nodeList = nodeList & CStr(n)
        Next n
        
        Debug.Print CStr(elemName) & " -> [" & nodeList & "]"
    Next elemName
    
    Debug.Print String(60, "=")
End Sub

' Core function to retrieve node connectivity from SAP2000 OAPI
' optimized for different element types
Public Function GetNodesFromElement_Core(ByVal elemName As String, ByVal objType As String) As Collection
    Dim result As New Collection
    Dim ret As Long
    
    If SapModel Is Nothing Then Set GetNodesFromElement_Core = result: Exit Function
    
    On Error Resume Next
    
    Select Case UCase(objType)
        Case "FRAME"
            Dim pt1 As String, pt2 As String
            ret = SapModel.frameObj.GetPoints(elemName, pt1, pt2)
            If ret = 0 Then
                If Trim$(pt1) <> "" Then result.Add pt1
                If Trim$(pt2) <> "" And pt2 <> pt1 Then result.Add pt2
            End If
            
        Case "AREA"
            Dim numPts As Long
            Dim ptNames() As String
            ' SAP2000 API requires dynamic array for Area GetPoints
            ret = SapModel.AreaObj.GetPoints(elemName, numPts, ptNames)
            
            If ret = 0 And numPts > 0 Then
                Dim i As Long
                ' Use a temp dictionary to avoid duplicates within the same area (rare but possible in bad geometry)
                Dim tempDict As Object
                Set tempDict = CreateObject("Scripting.Dictionary")
                
                For i = 0 To numPts - 1
                    Dim pName As String
                    pName = Trim$(ptNames(i))
                    If pName <> "" Then
                        If Not tempDict.exists(pName) Then
                            tempDict.Add pName, True
                            result.Add pName
                        End If
                    End If
                Next i
                Set tempDict = Nothing
            End If
            
        Case "CABLE"
            Dim cpt1 As String, cpt2 As String
            ret = SapModel.CableObj.GetPoints(elemName, cpt1, cpt2)
            If ret = 0 Then
                If Trim$(cpt1) <> "" Then result.Add cpt1
                If Trim$(cpt2) <> "" And cpt2 <> cpt1 Then result.Add cpt2
            End If
            
        Case "TENDON"
            Dim tpt1 As String, tpt2 As String
            ret = SapModel.TendonObj.GetPoints(elemName, tpt1, tpt2)
            If ret = 0 Then
                If Trim$(tpt1) <> "" Then result.Add tpt1
                If Trim$(tpt2) <> "" And tpt2 <> tpt1 Then result.Add tpt2
            End If
            
        Case "SOLID"
            Dim solidPts() As String
            ret = SapModel.SolidObj.GetPoints(elemName, solidPts)
            If ret = 0 Then
                Dim j As Long
                ' Solid points can be up to 8, ensure uniqueness
                Dim tempDictSol As Object
                Set tempDictSol = CreateObject("Scripting.Dictionary")
                
                If (Not Not solidPts) <> 0 Then ' Check if array is initialized
                    For j = LBound(solidPts) To UBound(solidPts)
                        Dim spt As String
                        spt = Trim$(CStr(solidPts(j)))
                        If spt <> "" Then
                            If Not tempDictSol.exists(spt) Then
                                tempDictSol.Add spt, True
                                result.Add spt
                            End If
                        End If
                    Next j
                End If
                Set tempDictSol = Nothing
            End If
            
        Case "LINK"
            Dim lpt1 As String, lpt2 As String
            ret = SapModel.LinkObj.GetPoints(elemName, lpt1, lpt2)
            If ret = 0 Then
                If Trim$(lpt1) <> "" Then result.Add lpt1
                ' Link can be 1-point or 2-point
                If Trim$(lpt2) <> "" And lpt2 <> lpt1 Then result.Add lpt2
            End If
            
    End Select
    
    On Error GoTo 0
    Set GetNodesFromElement_Core = result
End Function

' ===============================================================
' CORE LOGIC FUNCTIONS (extracted from UserForm)
' ===============================================================

Public Function GetElementsByPropertyValue_Core( _
    ByVal attrName As String, _
    ByVal propValue As String, _
    ByVal objType As String _
) As Collection

    Dim result As New Collection
    Dim ret As Long
    Dim cnt As Long
    Dim names() As String
    Dim i As Long
    
    Dim pvNorm As String
    pvNorm = Trim$(CStr(propValue))
    
    ' Handle "Any" - return all elements
    If StrComp(LCase$(pvNorm), "any", vbTextCompare) = 0 Then
        If LCase$(attrName) <> "support" And LCase$(attrName) <> "spring" Then
            Set result = GetAllElements_Core(objType)
            Set GetElementsByPropertyValue_Core = result
            Exit Function
        End If
    End If
    
    Select Case attrName
        Case "Name"
            result.Add propValue
            
        Case "GUID"
            Set result = GetElementsByGUID_Core(propValue, objType)
            
        Case "Group"
            Dim savedCount As Long
            Dim savedObjTypeArr() As Long
            Dim savedObjNameArr() As String
            savedCount = 0
            SaveSelectionState_Core savedCount, savedObjTypeArr, savedObjNameArr
            
            ret = SapModel.SelectObj.ClearSelection()
            On Error Resume Next
            ret = SapModel.SelectObj.Group(propValue, False)
            On Error GoTo 0
            Set result = GetSelectedElements_Core(objType)
            
            RestoreSelectionState_Core savedCount, savedObjTypeArr, savedObjNameArr
            
        Case "Section"
            Set result = GetElementsBySection_Core(pvNorm, objType)
            
        Case "Material"
            Set result = GetElementsByMaterial_Core(propValue, objType)
            
        Case "Constraint"
            Set result = GetElementsByConstraint_Core(propValue)
            
        Case "Support"
            Set result = GetElementsBySupport_Core(propValue, objType)
            
        Case "Spring"
            Set result = GetElementsBySpring_Core(propValue, objType)
            
        Case "Design-Type"
            Set result = GetElementsByDesignType_Core(propValue, objType)
            
        Case "AutoSection"
            Set result = GetElementsByAutoSection_Core(propValue, objType)
            
        Case "Area-Type"
            Set result = GetAllElements_Core(objType)
            
        Case "Property"
            Set result = GetElementsByLinkProperty_Core(propValue, objType)
            
        Case "Coordinate-X"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis_Core(objType, "X", propValue)
            End If
            
        Case "Coordinate-Y"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis_Core(objType, "Y", propValue)
            End If
            
        Case "Coordinate-Z"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis_Core(objType, "Z", propValue)
            End If
    End Select
    
    Set GetElementsByPropertyValue_Core = result
End Function

Private Function GetElementsBySection_Core(ByVal sectionName As String, ByVal objType As String) As Collection
    Dim result As New Collection
    Dim ret As Long
    
    ' Handle "None" or empty section
    If sectionName = "" Or StrComp(LCase$(sectionName), "none", vbTextCompare) = 0 Then
        Dim allElems As Collection
        Set allElems = GetAllElements_Core(objType)
        Dim e As Variant
        Dim pName As String, pAuto As String
        
        For Each e In allElems
            On Error Resume Next
            If UCase(objType) = "FRAME" Then
                ret = SapModel.frameObj.GetSection(CStr(e), pName, pAuto)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                        result.Add CStr(e)
                    End If
                End If
            ElseIf UCase(objType) = "AREA" Then
                ret = SapModel.AreaObj.GetProperty(CStr(e), pName)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                        result.Add CStr(e)
                    End If
                End If
            End If
        Next e
        
        Set GetElementsBySection_Core = result
        Exit Function
    End If
    
    ' Use selection API
    Dim sSavedCnt As Long
    Dim sSavedTypes() As Long
    Dim sSavedNames() As String
    sSavedCnt = 0
    SaveSelectionState_Core sSavedCnt, sSavedTypes, sSavedNames
    
    ret = SapModel.SelectObj.ClearSelection()
    Select Case objType
        Case "Frame": On Error Resume Next: ret = SapModel.SelectObj.PropertyFrame(sectionName, False): On Error GoTo 0
        Case "Cable": On Error Resume Next: ret = SapModel.SelectObj.PropertyCable(sectionName, False): On Error GoTo 0
        Case "Tendon": On Error Resume Next: ret = SapModel.SelectObj.PropertyTendon(sectionName, False): On Error GoTo 0
        Case "Area": On Error Resume Next: ret = SapModel.SelectObj.PropertyArea(sectionName, False): On Error GoTo 0
        Case "Solid": On Error Resume Next: ret = SapModel.SelectObj.PropertySolid(sectionName, False): On Error GoTo 0
        Case "Link": On Error Resume Next: ret = SapModel.SelectObj.PropertyLink(sectionName, False): On Error GoTo 0
    End Select
    
    Set result = GetSelectedElements_Core(objType)
    RestoreSelectionState_Core sSavedCnt, sSavedTypes, sSavedNames
    
    Set GetElementsBySection_Core = result
End Function

Private Function GetElementsByMaterial_Core(ByVal materialName As String, ByVal objType As String) As Collection
    Dim result As New Collection
    Dim matSavedCnt As Long
    Dim matSavedTypes() As Long
    Dim matSavedNames() As String
    matSavedCnt = 0
    
    SaveSelectionState_Core matSavedCnt, matSavedTypes, matSavedNames
    
    On Error Resume Next
    SapModel.SelectObj.ClearSelection
    SapModel.SelectObj.PropertyMaterial materialName, False
    On Error GoTo 0
    
    Set result = GetSelectedElements_Core(objType)
    RestoreSelectionState_Core matSavedCnt, matSavedTypes, matSavedNames
    
    Set GetElementsByMaterial_Core = result
End Function

Private Function GetElementsByConstraint_Core(ByVal constraintName As String) As Collection
    Dim result As New Collection
    Dim cSavedCnt As Long
    Dim cSavedTypes() As Long
    Dim cSavedNames() As String
    cSavedCnt = 0
    
    SaveSelectionState_Core cSavedCnt, cSavedTypes, cSavedNames
    
    On Error Resume Next
    SapModel.SelectObj.ClearSelection
    SapModel.SelectObj.Constraint constraintName, False
    On Error GoTo 0
    
    Set result = GetSelectedElements_Core("Node")
    RestoreSelectionState_Core cSavedCnt, cSavedTypes, cSavedNames
    
    Set GetElementsByConstraint_Core = result
End Function

Private Function GetElementsBySupport_Core(ByVal supportValue As String, ByVal objType As String) As Collection
    Dim result As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetElementsBySupport_Core = result
        Exit Function
    End If
    
    Dim subtype As String
    subtype = Trim$(supportValue)
    
    If LCase(subtype) = "any" Then
        Set result = GetAllElements_Core(objType)
        Set GetElementsBySupport_Core = result
        Exit Function
    End If
    
    If UCase(objType) = "NODE" Then
        Set result = GetNodesMatchingSupport_Core(subtype)
        Set GetElementsBySupport_Core = result
        Exit Function
    End If
    
    ' For non-node types: get matching nodes then map to elements
    Dim matchingNodes As Collection
    Set matchingNodes = GetNodesMatchingSupport_Core(subtype)
    
    Dim elems As Collection
    Set elems = GetAllElements_Core(objType)
    
    Dim addedElems As Object
    Set addedElems = CreateObject("Scripting.Dictionary")
    
    Dim e As Variant, n As Variant
    For Each e In elems
        Dim elemNodes As Collection
        Set elemNodes = GetNodesFromElement_Core(CStr(e), objType)
        
        Dim foundMatch As Boolean
        foundMatch = False
        For Each n In elemNodes
            If IsInCollection_Core(matchingNodes, CStr(n)) Then
                foundMatch = True
                Exit For
            End If
        Next n
        
        If foundMatch Then
            If Not addedElems.exists(CStr(e)) Then
                addedElems.Add CStr(e), True
                result.Add CStr(e)
            End If
        End If
    Next e
    
    Set GetElementsBySupport_Core = result
End Function

Private Function GetElementsBySpring_Core(ByVal springValue As String, ByVal objType As String) As Collection
    ' Similar to GetElementsBySupport_Core but for springs
    Dim result As New Collection
    Set GetElementsBySpring_Core = GetElementsBySupport_Core(springValue, objType)
End Function

Private Function GetElementsByDesignType_Core(ByVal designType As String, ByVal objType As String) As Collection
    Dim result As New Collection
    
    If objType <> "Frame" Then
        Set GetElementsByDesignType_Core = result
        Exit Function
    End If
    
    Dim allFrames As Collection
    Set allFrames = GetAllElements_Core("Frame")
    
    Dim fr As Variant
    Dim sectionName As String, sectionAuto As String
    Dim typeRebar As Long
    Dim mappedType As String
    Dim ret As Long
    
    For Each fr In allFrames
        sectionName = "": sectionAuto = ""
        On Error Resume Next
        ret = SapModel.frameObj.GetSection(CStr(fr), sectionName, sectionAuto)
        On Error GoTo 0
        
        If ret <> 0 Or Trim$(sectionName) = "" Then
            mappedType = "Program Determined"
        Else
            typeRebar = -1
            On Error Resume Next
            ret = SapModel.PropFrame.GetTypeRebar(sectionName, typeRebar)
            On Error GoTo 0
            
            If ret = 0 Then
                Select Case typeRebar
                    Case 1: mappedType = "Column"
                    Case 2: mappedType = "Beam"
                    Case Else: mappedType = "Program Determined"
                End Select
            Else
                mappedType = "Program Determined"
            End If
        End If
        
        If StrComp(designType, mappedType, vbTextCompare) = 0 Then
            result.Add CStr(fr)
        End If
    Next fr
    
    Set GetElementsByDesignType_Core = result
End Function

Private Function GetElementsByAutoSection_Core(ByVal autoSectionName As String, ByVal objType As String) As Collection
    Dim result As New Collection
    
    If objType <> "Frame" Then
        Set GetElementsByAutoSection_Core = result
        Exit Function
    End If
    
    Dim allFrs As Collection
    Set allFrs = GetAllElements_Core("Frame")
    
    Dim fItem As Variant
    Dim ret As Long
    Dim secName As String, secAuto As String
    
    For Each fItem In allFrs
        On Error Resume Next
        ret = SapModel.frameObj.GetSection(CStr(fItem), secName, secAuto)
        On Error GoTo 0
        
        If ret = 0 Then
            If Trim$(secName) = autoSectionName Or Trim$(secAuto) = autoSectionName Then
                result.Add CStr(fItem)
            End If
        End If
    Next fItem
    
    Set GetElementsByAutoSection_Core = result
End Function

Private Function GetElementsByLinkProperty_Core(ByVal propName As String, ByVal objType As String) As Collection
    Dim result As New Collection
    
    If objType <> "Link" Then
        Set GetElementsByLinkProperty_Core = result
        Exit Function
    End If
    
    Dim allLinks As Collection
    Set allLinks = GetAllElements_Core("Link")
    
    Dim l As Variant
    Dim ret As Long
    Dim linkProp As String
    
    For Each l In allLinks
        On Error Resume Next
        ret = SapModel.LinkObj.GetProperty(CStr(l), linkProp)
        On Error GoTo 0
        
        If ret = 0 And Trim$(linkProp) = propName Then
            result.Add CStr(l)
        End If
    Next l
    
    Set GetElementsByLinkProperty_Core = result
End Function

Private Function GetElementsByGUID_Core(ByVal guidValue As String, ByVal objType As String) As Collection
    Dim result As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetElementsByGUID_Core = result
        Exit Function
    End If
    
    Dim ret As Long, cnt As Long, names() As String, i As Long
    Dim objGuid As String
    
    Select Case objType
        Case "Node": ret = SapModel.pointObj.GetNameList(cnt, names)
        Case "Frame": ret = SapModel.frameObj.GetNameList(cnt, names)
        Case "Cable": ret = SapModel.CableObj.GetNameList(cnt, names)
        Case "Tendon": ret = SapModel.TendonObj.GetNameList(cnt, names)
        Case "Area": ret = SapModel.AreaObj.GetNameList(cnt, names)
        Case "Solid": ret = SapModel.SolidObj.GetNameList(cnt, names)
        Case "Link": ret = SapModel.LinkObj.GetNameList(cnt, names)
        Case Else: ret = -1
    End Select
    
    If ret = 0 And cnt > 0 Then
        For i = 0 To cnt - 1
            If Trim$(names(i)) <> "" Then
                objGuid = ""
                On Error Resume Next
                
                Select Case objType
                    Case "Node": ret = SapModel.pointObj.GetGUID(names(i), objGuid)
                    Case "Frame": ret = SapModel.frameObj.GetGUID(names(i), objGuid)
                    Case "Cable": ret = SapModel.CableObj.GetGUID(names(i), objGuid)
                    Case "Tendon": ret = SapModel.TendonObj.GetGUID(names(i), objGuid)
                    Case "Area": ret = SapModel.AreaObj.GetGUID(names(i), objGuid)
                    Case "Solid": ret = SapModel.SolidObj.GetGUID(names(i), objGuid)
                    Case "Link": ret = SapModel.LinkObj.GetGUID(names(i), objGuid)
                End Select
                
                On Error GoTo 0
                
                If ret = 0 And StrComp(Trim$(objGuid), Trim$(guidValue), vbBinaryCompare) = 0 Then
                    result.Add names(i)
                End If
            End If
        Next i
    End If
    
    Set GetElementsByGUID_Core = result
End Function

Private Function GetElementsByCoordinateAxis_Core(ByVal objType As String, ByVal axis As String, ByVal propValue As String) As Collection
    Dim result As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetElementsByCoordinateAxis_Core = result
        Exit Function
    End If
    
    Dim target As Double
    On Error GoTo CleanExit
    
    Dim numericStr As String
    numericStr = NumericStringFromDisplay_Core(propValue)
    target = CDbl(numericStr)
    
    Dim elems As Collection
    Set elems = GetAllElements_Core(objType)
    
    Dim e As Variant
    Dim ret As Long
    Dim px As Double, py As Double, pz As Double
    
    For Each e In elems
        px = 0: py = 0: pz = 0
        Dim found As Boolean
        found = False
        
        If objType = "Node" Then
            Dim nx As Double, ny As Double, nz As Double
            On Error Resume Next
            ret = SapModel.pointObj.GetCoordCartesian(CStr(e), nx, ny, nz)
            On Error GoTo 0
            If ret = 0 Then
                px = nx: py = ny: pz = nz
                found = True
            End If
            
        ElseIf objType = "Frame" Or objType = "Cable" Or objType = "Tendon" Then
            Dim sp As String, ep As String
            On Error Resume Next
            Select Case objType
                Case "Frame": ret = SapModel.frameObj.GetPoints(CStr(e), sp, ep)
                Case "Cable": ret = SapModel.CableObj.GetPoints(CStr(e), sp, ep)
                Case "Tendon": ret = SapModel.TendonObj.GetPoints(CStr(e), sp, ep)
                Case Else: ret = -1
            End Select
            On Error GoTo 0
            
            If ret = 0 And Trim$(sp) <> "" Then
                Dim sx As Double, sy As Double, sz As Double
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(sp, sx, sy, sz)
                On Error GoTo 0
                
                If ret = 0 Then
                    If Trim$(ep) <> "" Then
                        Dim ex As Double, ey As Double, ez As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(ep, ex, ey, ez)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = (sx + ex) / 2: py = (sy + ey) / 2: pz = (sz + ez) / 2
                            found = True
                        Else
                            px = sx: py = sy: pz = sz
                            found = True
                        End If
                    Else
                        px = sx: py = sy: pz = sz
                        found = True
                    End If
                End If
            End If
        End If
        
        If found Then
            Dim val As Double
            Select Case UCase(axis)
                Case "X": val = px
                Case "Y": val = py
                Case "Z": val = pz
                Case Else: val = px
            End Select
            
            If Abs(Round(val, 6) - Round(target, 6)) < 0.0000005 Then
                result.Add CStr(e)
            End If
        End If
    Next e
    
CleanExit:
    On Error GoTo 0
    Set GetElementsByCoordinateAxis_Core = result
End Function

Private Function GetNodesMatchingSupport_Core(ByVal supportSubtype As String) As Collection
    Dim nodesResult As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetNodesMatchingSupport_Core = nodesResult
        Exit Function
    End If
    
    Dim s As String
    s = LCase(Trim$(supportSubtype))
    
    Dim reqDOF(5) As Boolean
    Dim checkExact As Boolean
    Dim includeRestraints As Boolean
    Dim includeSprings As Boolean
    Dim i As Long
    
    includeRestraints = False
    includeSprings = False
    checkExact = False
    
    For i = 0 To 5: reqDOF(i) = False: Next i
    
    Select Case True
        Case InStr(s, "support - fixed") > 0 Or s = "fixed"
            For i = 0 To 5: reqDOF(i) = True: Next i
            includeRestraints = True
            checkExact = True
            
        Case InStr(s, "support - pinned") > 0 Or s = "pinned"
            reqDOF(0) = True: reqDOF(1) = True: reqDOF(2) = True
            includeRestraints = True
            checkExact = True
            
        Case InStr(s, "support - roller") > 0 Or s = "roller"
            reqDOF(1) = True
            includeRestraints = True
            checkExact = True
            
        Case InStr(s, "spring") > 0
            includeSprings = True
    End Select
    
    Dim DOF() As Boolean
    ReDim DOF(5)
    For i = 0 To 5: DOF(i) = reqDOF(i): Next i
    
    Dim added As Object
    Set added = CreateObject("Scripting.Dictionary")
    
    On Error GoTo FinishNodes
    
    If includeRestraints Then
        Dim savedCnt As Long
        Dim savedTypes() As Long
        Dim savedNames() As String
        savedCnt = 0
        SaveSelectionState_Core savedCnt, savedTypes, savedNames
        
        On Error Resume Next
        SapModel.SelectObj.ClearSelection
        SapModel.SelectObj.SupportedPoints DOF, "Local", False, True, False, False, False, False, False
        On Error GoTo 0
        
        Dim candidateRes As Collection
        Set candidateRes = GetSelectedElements_Core("Node")
        RestoreSelectionState_Core savedCnt, savedTypes, savedNames
        
        Dim nod As Variant
        For Each nod In candidateRes
            Dim restraints() As Boolean
            Dim ret As Long
            On Error Resume Next
            ret = SapModel.pointObj.GetRestraint(CStr(nod), restraints)
            On Error GoTo 0
            
            If ret = 0 And IsArray(restraints) Then
                Dim lb As Long, ub As Long
                lb = LBound(restraints): ub = UBound(restraints)
                Dim ok As Boolean
                ok = True
                
                For i = 0 To 5
                    If reqDOF(i) Then
                        If i < lb Or i > ub Then ok = False: Exit For
                        If restraints(i) = False Then ok = False: Exit For
                    ElseIf checkExact Then
                        If i >= lb And i <= ub Then
                            If restraints(i) = True Then ok = False: Exit For
                        End If
                    End If
                Next i
                
                If ok Then
                    If Not added.exists(CStr(nod)) Then
                        added.Add CStr(nod), True
                        nodesResult.Add CStr(nod)
                    End If
                End If
            End If
        Next nod
    End If
    
    If includeSprings Then
        Dim sSavedCnt As Long
        Dim sSavedTypes() As Long
        Dim sSavedNames() As String
        sSavedCnt = 0
        SaveSelectionState_Core sSavedCnt, sSavedTypes, sSavedNames
        
        On Error Resume Next
        SapModel.SelectObj.ClearSelection
        SapModel.SelectObj.SupportedPoints DOF, "Local", False, False, True, True, True, True, True
        On Error GoTo 0
        
        Dim candidateSpr As Collection
        Set candidateSpr = GetSelectedElements_Core("Node")
        RestoreSelectionState_Core sSavedCnt, sSavedTypes, sSavedNames
        
        Dim nk As Variant
        For Each nk In candidateSpr
            If Not added.exists(CStr(nk)) Then
                Dim k() As Double
                Dim ret2 As Long
                On Error Resume Next
                ReDim k(5)
                ret2 = SapModel.pointObj.GetSpring(CStr(nk), k)
                On Error GoTo 0
                
                Dim foundSpr As Boolean
                foundSpr = False
                If ret2 = 0 And IsArray(k) Then
                    For i = 0 To 5
                        If k(i) <> 0# Then
                            foundSpr = True
                            Exit For
                        End If
                    Next i
                End If
                
                If foundSpr Then
                    added.Add CStr(nk), True
                    nodesResult.Add CStr(nk)
                End If
            End If
        Next nk
    End If
    
FinishNodes:
    On Error GoTo 0
    Set GetNodesMatchingSupport_Core = nodesResult
End Function

Private Function FilterByCoordinateRange_Core( _
    ByVal inputColl As Collection, _
    ByVal objType As String, _
    ByVal axis As String, _
    ByVal minVal As Double, _
    ByVal maxVal As Double _
) As Collection
    
    Dim result As New Collection
    Dim elem As Variant
    
    For Each elem In inputColl
        Dim coordVal As Double
        coordVal = GetRepresentativeCoordinate_Core(CStr(elem), objType, axis)
        
        If coordVal >= minVal And coordVal <= maxVal Then
            result.Add CStr(elem)
        End If
    Next elem
    
    Set FilterByCoordinateRange_Core = result
End Function

Private Function GetRepresentativeCoordinate_Core( _
    ByVal elemName As String, _
    ByVal objType As String, _
    ByVal axis As String _
) As Double
    
    Dim ret As Long
    Dim px As Double, py As Double, pz As Double
    px = 0: py = 0: pz = 0
    
    On Error Resume Next
    
    If objType = "Node" Then
        ret = SapModel.pointObj.GetCoordCartesian(elemName, px, py, pz)
        
    ElseIf objType = "Frame" Or objType = "Cable" Or objType = "Tendon" Then
        Dim sp As String, ep As String
        Select Case objType
            Case "Frame": ret = SapModel.frameObj.GetPoints(elemName, sp, ep)
            Case "Cable": ret = SapModel.CableObj.GetPoints(elemName, sp, ep)
            Case "Tendon": ret = SapModel.TendonObj.GetPoints(elemName, sp, ep)
        End Select
        
        If ret = 0 And Trim$(sp) <> "" Then
            Dim sx As Double, sy As Double, sz As Double
            ret = SapModel.pointObj.GetCoordCartesian(sp, sx, sy, sz)
            
            If Trim$(ep) <> "" Then
                Dim ex As Double, ey As Double, ez As Double
                ret = SapModel.pointObj.GetCoordCartesian(ep, ex, ey, ez)
                px = (sx + ex) / 2: py = (sy + ey) / 2: pz = (sz + ez) / 2
            Else
                px = sx: py = sy: pz = sz
            End If
        End If
    End If
    
    On Error GoTo 0
    
    Select Case UCase(axis)
        Case "X": GetRepresentativeCoordinate_Core = px
        Case "Y": GetRepresentativeCoordinate_Core = py
        Case "Z": GetRepresentativeCoordinate_Core = pz
        Case Else: GetRepresentativeCoordinate_Core = px
    End Select
End Function

' ===============================================================
' HELPER FUNCTIONS
' ===============================================================

Public Function GetAllElements_Core(ByVal objType As String) As Collection
    Dim result As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetAllElements_Core = result
        Exit Function
    End If
    
    Dim ret As Long, cnt As Long, names() As String, i As Long
    
    Select Case objType
        Case "Node": ret = SapModel.pointObj.GetNameList(cnt, names)
        Case "Frame": ret = SapModel.frameObj.GetNameList(cnt, names)
        Case "Cable": ret = SapModel.CableObj.GetNameList(cnt, names)
        Case "Tendon": ret = SapModel.TendonObj.GetNameList(cnt, names)
        Case "Area": ret = SapModel.AreaObj.GetNameList(cnt, names)
        Case "Solid": ret = SapModel.SolidObj.GetNameList(cnt, names)
        Case "Link": ret = SapModel.LinkObj.GetNameList(cnt, names)
        Case Else: ret = -1
    End Select
    
    If ret = 0 And cnt > 0 Then
        For i = 0 To cnt - 1
            If Trim$(names(i)) <> "" Then result.Add names(i)
        Next i
    End If
    
    Set GetAllElements_Core = result
End Function

Public Function GetSelectedElements_Core(ByVal itemType As String) As Collection
    Dim result As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetSelectedElements_Core = result
        Exit Function
    End If
    
    Dim ret As Long
    Dim selCount As Long
    Dim objTypes() As Long
    Dim objNames() As String
    
    On Error Resume Next
    ret = SapModel.SelectObj.GetSelected(selCount, objTypes, objNames)
    On Error GoTo 0
    
    If ret <> 0 Or selCount <= 0 Then
        Set GetSelectedElements_Core = result
        Exit Function
    End If
    
    Dim wantTypeCode As Long
    wantTypeCode = GetObjectTypeCode_Core(itemType)
    
    Dim i As Long
    For i = 0 To selCount - 1
        If objTypes(i) = wantTypeCode Then
            If Trim$(objNames(i)) <> "" Then
                result.Add objNames(i)
            End If
        End If
    Next i
    
    Set GetSelectedElements_Core = result
End Function

Public Sub SaveSelectionState_Core(ByRef outCount As Long, ByRef outTypes() As Long, ByRef outNames() As String)
    outCount = 0
    On Error Resume Next
    If SapModel Is Nothing Then Exit Sub
    
    Dim ret As Long
    ret = SapModel.SelectObj.GetSelected(outCount, outTypes, outNames)
    
    If ret <> 0 Or outCount <= 0 Then
        outCount = 0
        Erase outTypes
        Erase outNames
    End If
    On Error GoTo 0
End Sub

Public Sub RestoreSelectionState_Core(ByVal inCount As Long, ByRef inTypes() As Long, ByRef inNames() As String)
    On Error Resume Next
    If SapModel Is Nothing Then Exit Sub
    
    SapModel.SelectObj.ClearSelection
    
    Dim i As Long
    For i = 0 To inCount - 1
        Dim tcode As Long
        Dim nm As String
        tcode = inTypes(i)
        nm = inNames(i)
        
        Dim tName As String
        tName = ObjectTypeCodeToName_Core(tcode)
        
        If tName <> "" Then
            Call SelectSingleObject_Core(tName, CStr(nm), True)
        End If
    Next i
    On Error GoTo 0
End Sub

Public Function SelectSingleObject_Core(ByVal objType As String, ByVal objName As String, ByVal selectFlag As Boolean) As Long
    Dim ret As Long
    If Not EnsureSapModelAvailable() Then
        SelectSingleObject_Core = -1
        Exit Function
    End If
    
    On Error Resume Next
    Select Case objType
        Case "Node": ret = SapModel.pointObj.SetSelected(objName, selectFlag)
        Case "Frame": ret = SapModel.frameObj.SetSelected(objName, selectFlag)
        Case "Cable": ret = SapModel.CableObj.SetSelected(objName, selectFlag)
        Case "Tendon": ret = SapModel.TendonObj.SetSelected(objName, selectFlag)
        Case "Area": ret = SapModel.AreaObj.SetSelected(objName, selectFlag)
        Case "Solid": ret = SapModel.SolidObj.SetSelected(objName, selectFlag)
        Case "Link": ret = SapModel.LinkObj.SetSelected(objName, selectFlag)
        Case Else: ret = -1
    End Select
    On Error GoTo 0
    
    SelectSingleObject_Core = ret
End Function

Public Function GetObjectTypeCode_Core(ByVal objType As String) As Long
    Select Case objType
        Case "Node": GetObjectTypeCode_Core = 1
        Case "Frame": GetObjectTypeCode_Core = 2
        Case "Cable": GetObjectTypeCode_Core = 3
        Case "Tendon": GetObjectTypeCode_Core = 4
        Case "Area": GetObjectTypeCode_Core = 5
        Case "Solid": GetObjectTypeCode_Core = 6
        Case "Link": GetObjectTypeCode_Core = 7
        Case Else: GetObjectTypeCode_Core = 0
    End Select
End Function

Public Function ObjectTypeCodeToName_Core(ByVal Code As Long) As String
    Select Case Code
        Case 1: ObjectTypeCodeToName_Core = "Node"
        Case 2: ObjectTypeCodeToName_Core = "Frame"
        Case 3: ObjectTypeCodeToName_Core = "Cable"
        Case 4: ObjectTypeCodeToName_Core = "Tendon"
        Case 5: ObjectTypeCodeToName_Core = "Area"
        Case 6: ObjectTypeCodeToName_Core = "Solid"
        Case 7: ObjectTypeCodeToName_Core = "Link"
        Case Else: ObjectTypeCodeToName_Core = ""
    End Select
End Function

Private Function NumericStringFromDisplay_Core(ByVal displayValue As String) As String
    Dim s As String
    s = Trim$(displayValue)
    If s = "" Then
        NumericStringFromDisplay_Core = "0"
        Exit Function
    End If
    
    Dim pos As Long
    pos = InStr(s, "(")
    If pos > 0 Then
        s = Trim$(Left$(s, pos - 1))
    Else
        pos = InStr(s, " ")
        If pos > 0 Then s = Trim$(Left$(s, pos - 1))
    End If
    
    NumericStringFromDisplay_Core = s
End Function

Private Function IsInCollection_Core(ByVal col As Collection, ByVal key As String) As Boolean
    Dim item As Variant
    IsInCollection_Core = False
    For Each item In col
        If CStr(item) = key Then
            IsInCollection_Core = True
            Exit Function
        End If
    Next item
End Function
' ---------------- Helper: ensure SapModel is connected ----------------
Public Function EnsureSapModelAvailable() As Boolean
    On Error Resume Next
    If Not (Not SapModel Is Nothing) Then
        ' Try to connect via central ConnectSAP2000
        If Not ConnectSAP2000() Then
            EnsureSapModelAvailable = False
            Exit Function
        End If
    End If
    EnsureSapModelAvailable = Not (SapModel Is Nothing)
End Function
' ----------------------------
' AutoCAD Connection
' ----------------------------

Public Function GetOrCreateAutoCAD() As Object
    On Error Resume Next
    
    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application")
    End If
    
    If Not acadApp Is Nothing Then
        acadApp.Visible = True
    End If
    
    Set GetOrCreateAutoCAD = acadApp
    On Error GoTo 0
End Function

Public Sub ShowExcelOnTop()
    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo 0
End Sub

' ----------------------------
' AutoCAD View Helpers
' ----------------------------

Public Sub ZoomAll(acadDoc As Object)
    On Error Resume Next
    acadDoc.SendCommand ".ZOOM A " & vbCr
    
    Dim t As Double
    t = Timer
    Do While Timer < t + 1.2
        DoEvents
    Loop
    
    acadDoc.Regen 0
    On Error GoTo 0
End Sub

Public Sub EnableViewCube(acadApp As Object, acadDoc As Object)
    On Error Resume Next
    
    acadDoc.SetVariable "NAVVCUBE", 1
    acadDoc.SetVariable "NAVVCUBEDISPLAY", 3
    
    Dim t As Double
    t = Timer
    Do While Timer < t + 0.2
        DoEvents
    Loop
    
    On Error GoTo 0
End Sub

Public Sub SetInsUnits(acadDoc As Object, units As Long)
    On Error Resume Next
    acadDoc.SetVariable "INSUNITS", units
    On Error GoTo 0
End Sub

' ----------------------------
' Status Callback Helper - FIXED FOR VBA ENCODING
' ----------------------------

Public Sub SetStatusToForm(msg As String)
    On Error Resume Next
    
    ' Remove special characters that VBA can't handle
    msg = CleanMessageForVBA(msg)
    
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "frmSyncCADSAP" Then
            With uf
                .txtStatus.text = msg & vbCrLf & .txtStatus.text
                DoEvents
            End With
            Exit For
        End If
    Next uf
    
    On Error GoTo 0
End Sub

' Clean message from special characters
Private Function CleanMessageForVBA(msg As String) As String
    Dim result As String
    result = msg
    
    ' Replace special arrows and symbols
    result = Replace(result, "?", "->")
    result = Replace(result, "?", "<-")
    result = Replace(result, "?", "<->")
    result = Replace(result, "?", "=>")
    result = Replace(result, "?", "<=")
    result = Replace(result, "?", "<=>")
    result = Replace(result, "?", ">")
    result = Replace(result, "?", "<")
    result = Replace(result, "?", ">")
    result = Replace(result, "?", "<")
    result = Replace(result, "?", "->")  ' Your specific case
    
    ' Replace box drawing characters
    result = Replace(result, "-", "=")
    result = Replace(result, "?", "-")
    result = Replace(result, "-", "-")
    result = Replace(result, "¦", "|")
    result = Replace(result, "+", "+")
    result = Replace(result, "+", "+")
    result = Replace(result, "+", "+")
    result = Replace(result, "+", "+")
    
    ' Replace bullets and special symbols
    result = Replace(result, "•", "*")
    result = Replace(result, "?", "*")
    result = Replace(result, "?", "o")
    result = Replace(result, "¦", "#")
    result = Replace(result, "?", "#")
    result = Replace(result, "?", "OK")
    result = Replace(result, "?", "X")
    result = Replace(result, "?", "X")
    
    CleanMessageForVBA = result
End Function

'' ----------------------------
'' Array Utilities
'' ----------------------------
'
'Public Function IsArrayEmpty(arr As Variant) As Boolean
'    On Error Resume Next
'    IsArrayEmpty = (LBound(arr) > UBound(arr))
'    If err.number <> 0 Then IsArrayEmpty = True
'    On Error GoTo 0
'End Function

' ----------------------------
' String Utilities
' ----------------------------

Public Function FormatCoords(X As Double, Y As Double, Z As Double) As String
    FormatCoords = "(" & Format(X, "0.00") & "," & Format(Y, "0.00") & "," & Format(Z, "0.00") & ")"
End Function

Public Function FormatCoordsArray(coords() As Double, Optional nDim As Long = 3) As String
    If IsArrayEmpty(coords) Then
        FormatCoordsArray = "(empty)"
        Exit Function
    End If
    
    Dim result As String
    result = "("
    
    Dim i As Long
    For i = LBound(coords) To UBound(coords)
        If i > LBound(coords) Then result = result & ","
        result = result & Format(coords(i), "0.00")
    Next i
    
    result = result & ")"
    FormatCoordsArray = result
End Function

' ----------------------------
' Geometry Utilities
' ----------------------------

Public Function Distance3D(x1 As Double, y1 As Double, z1 As Double, _
                           x2 As Double, y2 As Double, z2 As Double) As Double
    Dim dx As Double, dy As Double, dz As Double
    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1
    Distance3D = Sqr(dx * dx + dy * dy + dz * dz)
End Function

Public Function DistanceSquared3D(x1 As Double, y1 As Double, z1 As Double, _
                                  x2 As Double, y2 As Double, z2 As Double) As Double
    Dim dx As Double, dy As Double, dz As Double
    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1
    DistanceSquared3D = dx * dx + dy * dy + dz * dz
End Function

' Check if points form horizontal plane (within tolerance)
Public Function IsHorizontalPlane(coords3D() As Double, Optional toleranceMM As Double = 10) As Boolean
    IsHorizontalPlane = False
    
    If IsArrayEmpty(coords3D) Then Exit Function
    
    Dim nPts As Long
    nPts = (UBound(coords3D) + 1) / 3
    
    If nPts < 3 Then Exit Function
    
    Dim minZ As Double, maxZ As Double
    minZ = coords3D(2) ' First Z
    maxZ = coords3D(2)
    
    Dim i As Long
    For i = 1 To nPts - 1
        Dim Z As Double
        Z = coords3D(i * 3 + 2)
        If Z < minZ Then minZ = Z
        If Z > maxZ Then maxZ = Z
    Next i
    
    IsHorizontalPlane = (Abs(maxZ - minZ) < toleranceMM)
End Function

' ----------------------------
' Dictionary Utilities
' ----------------------------

Public Function CreateDict() As Object
    Set CreateDict = CreateObject("Scripting.Dictionary")
End Function

' ----------------------------
' Validation
' ----------------------------

Public Function ValidatePositiveNumber(Value As Variant, defaultValue As Double) As Double
    Dim num As Double
    num = val(Value)
    
    If num <= 0 Then
        ValidatePositiveNumber = defaultValue
    Else
        ValidatePositiveNumber = num
    End If
End Function

Public Function ValidateTolerance(Value As Variant) As Double
    ValidateTolerance = ValidatePositiveNumber(Value, 1#)
End Function

Public Function ValidateScaleFactor(Value As Variant) As Double
    ValidateScaleFactor = ValidatePositiveNumber(Value, 1#)
End Function

' ----------------------------
' Logging
' ----------------------------

Public Sub LogMessage(msg As String)
    ' Can be extended to write to file or debug window
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
End Sub

Public Sub LogError(procName As String, errNum As Long, errDesc As String)
    LogMessage "ERROR in " & procName & ": #" & errNum & " - " & errDesc
End Sub

' ----------------------------
' SAP2000 Connection Check
' ----------------------------

Public Function IsSAPConnected(SapModel As Object) As Boolean
    On Error Resume Next
    IsSAPConnected = False
    
    If SapModel Is Nothing Then Exit Function
    
    ' Try to get model info
    Dim modelName As String
    modelName = SapModel.GetModelFilename
    
    If err.number = 0 Then
        IsSAPConnected = True
    End If
    
    On Error GoTo 0
End Function
Public Sub RegisterDataApp(acadDoc As Object)
    On Error Resume Next
    acadDoc.RegisteredApplications.Add "DTS_APP"
    On Error GoTo 0
End Sub
Public Sub EnsureLayerExists(acadDoc As Object, layerName As String, colorIndex As Long)
    On Error Resume Next
    Dim lay As Object
    Set lay = acadDoc.layers.item(layerName)
    If err.number <> 0 Then
        err.Clear
        Set lay = acadDoc.layers.Add(layerName)
        lay.color = colorIndex
    End If
    On Error GoTo 0
End Sub

