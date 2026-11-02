Attribute VB_Name = "FrameCutListProbe"
Option Explicit

' Read-only geometry/cut-list probe. No dimensions, rebuilds, configuration
' switches, cut-list updates, selection changes or document saves.
Private Const PREFIX As String = "FrameCutListProbe v1.0"
Private app As SldWorks.SldWorks
Private drawing As SldWorks.ModelDoc2
Private view As SldWorks.View
Private frameModel As SldWorks.ModelDoc2
Private frameInstance As SldWorks.Component2
Private visited As Collection
Private folders As Long, bodiesReported As Long, errors As Long
Private output As Object
Private outputPath As String

Public Sub main()
    On Error GoTo Failed
    Set output = Nothing
    Set frameInstance = Nothing
    Set visited = New Collection
    folders = 0: bodiesReported = 0: errors = 0
    outputPath = ""
    OpenLog
    LogLine PREFIX & " START " & CStr(Now)
    Set app = Application.SldWorks
    Set drawing = app.ActiveDoc
    If drawing Is Nothing Then Err.Raise vbObjectError + 1, , "Open the assembly drawing."
    If drawing.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."
    Dim sm As SldWorks.SelectionMgr, selectedEdge As SldWorks.Edge
    Set sm = drawing.SelectionManager
    If sm.GetSelectedObjectCount2(-1) <> 1 Then _
        Err.Raise vbObjectError + 3, , "Select ONE visible edge of the multibody FRAME part in the drawing view."
    If sm.GetSelectedObjectType3(1, -1) <> swSelEDGES Then _
        Err.Raise vbObjectError + 4, , "Select an edge of the FRAME part, not the view border or an accessory."
    Set selectedEdge = sm.GetSelectedObject6(1, -1)
    Set view = sm.GetSelectedObjectsDrawingView2(1, -1)
    If view Is Nothing Then Err.Raise vbObjectError + 5, , "Selected edge has no drawing view."
    Dim reference As SldWorks.ModelDoc2
    Set reference = view.ReferencedDocument
    If reference Is Nothing Then Err.Raise vbObjectError + 6, , "Resolve/open the referenced model first."
    LogLine "VIEW=[" & view.GetName2 & "]; reference=[" & reference.GetTitle & "]; config=[" & view.ReferencedConfiguration & "]"
    CheckConfiguration reference, view.ReferencedConfiguration
    If reference.GetType = swDocASSEMBLY Then
        Dim entity As SldWorks.Entity, dc As SldWorks.DrawingComponent, drawingComp As SldWorks.Component2
        Set entity = selectedEdge
        Set dc = entity.GetDrawingComponent(view)
        If dc Is Nothing Then Err.Raise vbObjectError + 7, , "Cannot obtain drawing component of selected edge."
        Set drawingComp = dc.Component
        If drawingComp Is Nothing Then Err.Raise vbObjectError + 8, , "Selected drawing component unavailable."
        Set frameInstance = ResolveFrameInstance(reference, drawingComp)
        Set frameModel = frameInstance.GetModelDoc2
        If frameModel Is Nothing Then Err.Raise vbObjectError + 9, , "Resolve/open the selected frame part first."
        CheckConfiguration frameModel, frameInstance.ReferencedConfiguration
        LogLine "FRAME INSTANCE=[" & frameInstance.Name2 & "]; config=[" & frameInstance.ReferencedConfiguration & "]"
    ElseIf reference.GetType = swDocPART Then
        Set frameModel = reference
    Else
        Err.Raise vbObjectError + 10, , "Only assembly and part references supported."
    End If
    If frameModel.GetType <> swDocPART Then Err.Raise vbObjectError + 11, , "Selected component must be the multibody FRAME PART."
    LogLine "FRAME MODEL=[" & frameModel.GetTitle & "]; path=[" & frameModel.GetPathName & "]"
    Dim framePart As SldWorks.PartDoc, solids As Variant, solidCount As Long
    Set framePart = frameModel
    solids = framePart.GetBodies2(swSolidBody, False)
    If HasItems(solids) Then solidCount = UBound(solids) - LBound(solids) + 1
    LogLine "FRAME solid bodies=" & solidCount
    LogLine "ONLY this selected part is inspected. Property values below have NO assumed units."
    DescribeEdge selectedEdge, "SELECTED VIEW EDGE"
    DescribeSourceEdge selectedEdge
    Dim feature As SldWorks.Feature
    Set feature = frameModel.FirstFeature
    Do While Not feature Is Nothing
        VisitFeature feature
        Set feature = feature.GetNextFeature
    Loop
    LogLine "END folders=" & folders & "; body entries=" & bodiesReported & "; errors=" & errors
    CloseLog
    MsgBox PREFIX & ": folders=" & folders & "; body entries=" & bodiesReported & "; errors=" & errors & vbCrLf & _
        "No dimensions created. No model saved or rebuilt." & vbCrLf & "Log: " & outputPath, vbInformation
    Exit Sub
Failed:
    Dim message As String
    message = CStr(Err.Number) & ": " & Err.Description
    LogLine "STOP " & message
    CloseLog
    MsgBox PREFIX & " stopped: " & message & vbCrLf & "Log: " & outputPath, vbExclamation
End Sub

Private Sub OpenLog()
    On Error GoTo NoFile
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    outputPath = fs.BuildPath(fs.GetSpecialFolder(2), "FrameCutListProbe-" & fs.GetTempName & ".txt")
    ' Unicode (UTF-16) preserves Russian component and property names.
    ' overwrite=False protects existing files.
    Set output = fs.CreateTextFile(outputPath, False, True)
    Exit Sub
NoFile:
    outputPath = "VBA Immediate window (Ctrl+G); log file unavailable"
    Debug.Print PREFIX & " log file error: " & Err.Description
End Sub

Private Sub LogLine(ByVal message As String)
    Debug.Print PREFIX & " " & message
    If Not output Is Nothing Then WriteLogLine message
End Sub

Private Sub WriteLogLine(ByVal message As String)
    On Error GoTo WriteFailed
    output.WriteLine PREFIX & " " & message
    Exit Sub
WriteFailed:
    Debug.Print PREFIX & " log file write failed: " & Err.Description
    Set output = Nothing
End Sub

Private Sub CloseLog()
    On Error GoTo CloseFailed
    If Not output Is Nothing Then output.Close
    Set output = Nothing
    Exit Sub
CloseFailed:
    Debug.Print PREFIX & " log file close failed: " & Err.Description
    Set output = Nothing
End Sub

Private Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasItems = (UBound(values) >= LBound(values))
EmptyArray:
End Function

Private Sub CheckConfiguration(ByVal model As SldWorks.ModelDoc2, ByVal requested As String)
    Dim active As SldWorks.Configuration
    Set active = model.ConfigurationManager.ActiveConfiguration
    If active Is Nothing Then Err.Raise vbObjectError + 12, , "Active source configuration unavailable."
    LogLine "CONFIG model=[" & model.GetTitle & "]; active=[" & active.Name & "]; requested=[" & requested & "]"
    If Len(Trim$(requested)) > 0 Then
        If StrComp(active.Name, requested, vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 13, , "Configuration mismatch. Probe does not switch configurations."
    End If
End Sub

Private Function ResolveFrameInstance(ByVal model As SldWorks.ModelDoc2, _
        ByVal requested As SldWorks.Component2) As SldWorks.Component2
    Dim assembly As SldWorks.AssemblyDoc, all As Variant, i As Long, matches As Long
    Dim candidate As SldWorks.Component2, found As SldWorks.Component2
    Dim root As SldWorks.DrawingComponent, rootComp As SldWorks.Component2
    Dim rootName As String, rootAlias As String, sameName As Boolean
    Dim path As String, config As String
    Set root = view.RootDrawingComponent
    If Not root Is Nothing Then
        rootName = root.Name
        Set rootComp = root.Component
        If Not rootComp Is Nothing Then rootAlias = rootComp.Name2
    End If
    Set assembly = model
    all = assembly.GetComponents(False)
    If Not HasItems(all) Then Err.Raise vbObjectError + 14, , "Referenced assembly component tree unavailable."
    path = requested.GetPathName: config = requested.ReferencedConfiguration
    LogLine "DRAWING COMPONENT=[" & requested.Name2 & "]; root=[" & rootName & "]; root alias=[" & rootAlias & "]"
    For i = LBound(all) To UBound(all)
        Set candidate = all(i)
        If Not candidate Is Nothing Then
            sameName = (StrComp(requested.Name2, candidate.Name2, vbTextCompare) = 0)
            If Len(rootName) > 0 Then sameName = sameName Or (StrComp(requested.Name2, rootName & "/" & candidate.Name2, vbTextCompare) = 0)
            If Len(rootAlias) > 0 Then sameName = sameName Or (StrComp(requested.Name2, rootAlias & "/" & candidate.Name2, vbTextCompare) = 0)
            If sameName Then
                If Len(path) > 0 Then sameName = (StrComp(Replace(path, "/", "\"), Replace(candidate.GetPathName, "/", "\"), vbTextCompare) = 0)
                If Len(Trim$(config)) > 0 Then sameName = sameName And (StrComp(config, candidate.ReferencedConfiguration, vbTextCompare) = 0)
                If sameName Then
                    matches = matches + 1
                    Set found = candidate
                End If
            End If
        End If
    Next i
    If matches <> 1 Then Err.Raise vbObjectError + 15, , "Frame instance resolution: accepted=" & matches
    Set ResolveFrameInstance = found
End Function

Private Function FirstVisit(ByVal feature As SldWorks.Feature) As Boolean
    On Error GoTo AlreadySeen
    ' Feature names are unique in one part; use them as keys to avoid traversing
    ' duplicate top-level/subfeature presentations of the same cut-list folder.
    visited.Add feature.Name, feature.Name
    FirstVisit = True
    Exit Function
AlreadySeen:
    If Err.Number <> 457 Then Err.Raise Err.Number, , Err.Description
End Function

Private Sub VisitFeature(ByVal feature As SldWorks.Feature)
    If Not FirstVisit(feature) Then Exit Sub
    If StrComp(feature.GetTypeName2, "CutListFolder", vbTextCompare) = 0 Then ReportFolder feature
    Dim child As SldWorks.Feature
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        VisitFeature child
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Sub ReportFolder(ByVal feature As SldWorks.Feature)
    On Error GoTo FolderFailed
    folders = folders + 1
    Dim folder As SldWorks.BodyFolder, bodies As Variant, i As Long, body As SldWorks.Body2
    Set folder = feature.GetSpecificFeature2
    If folder Is Nothing Then Err.Raise vbObjectError + 16, , "Cut-list body folder unavailable."
    LogLine "CUTLIST=[" & feature.Name & "]; API body count=" & folder.GetBodyCount
    ReportProperties feature
    bodies = folder.GetBodies
    If Not HasItems(bodies) Then
        LogLine "CUTLIST has no accessible bodies"
        Exit Sub
    End If
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        If Not body Is Nothing Then
            bodiesReported = bodiesReported + 1
            ReportBody body
        End If
    Next i
    Exit Sub
FolderFailed:
    errors = errors + 1
    LogLine "CUTLIST ERROR [" & feature.Name & "]: " & Err.Number & ": " & Err.Description
End Sub

Private Sub ReportProperties(ByVal feature As SldWorks.Feature)
    On Error GoTo PropertiesFailed
    Dim manager As SldWorks.CustomPropertyManager, names As Variant, i As Long
    Set manager = feature.CustomPropertyManager
    names = manager.GetNames
    If Not HasItems(names) Then
        LogLine "  No cached cut-list properties"
        Exit Sub
    End If
    For i = LBound(names) To UBound(names)
        ReportProperty manager, CStr(names(i))
    Next i
    Exit Sub
PropertiesFailed:
    errors = errors + 1
    LogLine "  PROPERTY ENUMERATION ERROR: " & Err.Number & ": " & Err.Description
End Sub

Private Sub ReportProperty(ByVal manager As SldWorks.CustomPropertyManager, ByVal name As String)
    On Error GoTo PropertyFailed
    Dim raw As String, resolved As String, wasResolved As Boolean, linked As Boolean, status As Long
    status = manager.Get6(name, True, raw, resolved, wasResolved, linked)
    LogLine "  PROPERTY [" & name & "]; raw=[" & raw & "]; resolved=[" & resolved & _
        "]; status=" & status & "; wasResolved=" & wasResolved & "; linked=" & linked
    Exit Sub
PropertyFailed:
    errors = errors + 1
    LogLine "  PROPERTY ERROR [" & name & "]: " & Err.Number & ": " & Err.Description
End Sub

Private Sub ReportBody(ByVal body As SldWorks.Body2)
    On Error GoTo BodyFailed
    Dim edges As Variant, edge As SldWorks.Edge, longest As SldWorks.Edge
    Dim curve As SldWorks.Curve, a As SldWorks.Vertex, b As SldWorks.Vertex
    Dim pa As Variant, pb As Variant, i As Long, straight As Long, length As Double, maximum As Double
    LogLine "  BODY=[" & body.Name & "]"
    edges = body.GetEdges
    If Not HasItems(edges) Then Exit Sub
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        Set curve = edge.GetCurve
        If Not curve Is Nothing Then
            If CBool(curve.IsLine) Then
                straight = straight + 1
                Set a = edge.GetStartVertex: Set b = edge.GetEndVertex
                If Not a Is Nothing And Not b Is Nothing Then
                    pa = a.GetPoint: pb = b.GetPoint
                    length = Sqr((pb(0) - pa(0)) ^ 2 + (pb(1) - pa(1)) ^ 2 + (pb(2) - pa(2)) ^ 2)
                    If length > maximum Then
                        maximum = length
                        Set longest = edge
                    End If
                End If
            End If
        End If
    Next i
    LogLine "  BODY edges=" & (UBound(edges) - LBound(edges) + 1) & "; straight=" & straight & "; max straight edge m=" & CStr(maximum)
    If Not longest Is Nothing Then DescribeEdge longest, "  SOURCE BODY LONGEST EDGE (not assumed equal to cut-list length)"
    Exit Sub
BodyFailed:
    errors = errors + 1
    LogLine "  BODY ERROR: " & Err.Number & ": " & Err.Description
End Sub

Private Sub DescribeSourceEdge(ByVal selected As SldWorks.Edge)
    On Error GoTo ConversionFailed
    Dim converted As Object, edge As SldWorks.Edge
    Set converted = frameModel.Extension.GetCorrespondingEntity2(selected)
    If converted Is Nothing Then
        LogLine "SOURCE CONVERSION: Nothing"
    ElseIf TypeOf converted Is SldWorks.Edge Then
        Set edge = converted
        DescribeEdge edge, "REVERSE-CONVERTED SOURCE EDGE"
    Else
        LogLine "SOURCE CONVERSION: unexpected type=" & TypeName(converted)
    End If
    Exit Sub
ConversionFailed:
    errors = errors + 1
    LogLine "SOURCE CONVERSION ERROR: " & Err.Number & ": " & Err.Description
End Sub

Private Sub DescribeEdge(ByVal edge As SldWorks.Edge, ByVal label As String)
    On Error GoTo EdgeFailed
    Dim curve As SldWorks.Curve, a As SldWorks.Vertex, b As SldWorks.Vertex
    Dim pa As Variant, pb As Variant
    Set curve = edge.GetCurve
    If curve Is Nothing Then Err.Raise vbObjectError + 17, , "Curve unavailable"
    LogLine label & ": curve Identity=" & curve.Identity & "; IsLine=" & CBool(curve.IsLine)
    Set a = edge.GetStartVertex: Set b = edge.GetEndVertex
    If a Is Nothing Or b Is Nothing Then
        LogLine label & ": vertices unavailable"
        Exit Sub
    End If
    pa = a.GetPoint: pb = b.GetPoint
    LogLine label & ": p0=(" & pa(0) & ", " & pa(1) & ", " & pa(2) & "); p1=(" & pb(0) & ", " & pb(1) & ", " & pb(2) & ")"
    Exit Sub
EdgeFailed:
    errors = errors + 1
    LogLine label & " ERROR: " & Err.Number & ": " & Err.Description
End Sub
