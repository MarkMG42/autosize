Attribute VB_Name = "AutoDimensionFrame"
Option Explicit

' Requires SolidWorks type libraries (normal SolidWorks VBA macro project).
' Distances and coordinates are in metres. No document is saved automatically.
Private Const MACRO_VERSION As String = "2.4.3"
Private Const LOG_PREFIX As String = "AutoFrameDim v" & MACRO_VERSION
Private Const OFFSET_M As Double = 0.004
Private Const OUTLINE_TOL_M As Double = 0.001
Private Const MIN_LENGTH_M As Double = 0.00000001
Private Const MIN_PROJECTED_M As Double = 0.00001
Private Const LENGTH_TOL_M As Double = 0.000001

Private swApp As SldWorks.SldWorks
Private doc As SldWorks.ModelDoc2
Private targetView As SldWorks.View
Private created As Long, skipped As Long, failed As Long
Private blocked As Boolean
Private preferenceCaptured As Boolean, oldInputPreference As Boolean
Private viewOutline As Variant
Private assemblyView As Boolean
Private sheetWidth As Double, sheetHeight As Double
Private occupied As Collection
Private localEdges As Collection, instanceEdges As Collection, assemblyEdges As Collection
Private viewModel As SldWorks.ModelDoc2, sourceModel As SldWorks.ModelDoc2
Private sourceInstance As SldWorks.Component2
Private modelComponents As Variant
Private drawingRootName As String, drawingRootAlias As String
Private reverseMatches As Long, forwardMatches As Long, missingMatches As Long
Private correspondenceErrors As Long, forwardBuilt As Boolean
Private reverseErrors As Long, assemblyErrors As Long, viewErrors As Long
Private reverseNulls As Long, reverseWrongTypes As Long, reverseNonLocal As Long
Private bodiesSeen As Long, unresolved As Long
Private Const TEXT_HEIGHT_M As Double = 0.004
Private Const LAYOUT_GAP_M As Double = 0.001
Private Const POSITION_TOL_M As Double = 0.00001

Public Sub main()
    On Error GoTo Fatal
    preferenceCaptured = False
    created = 0: skipped = 0: failed = 0: blocked = False
    bodiesSeen = 0: unresolved = 0
    Set occupied = New Collection
    Set doc = Nothing
    Set targetView = Nothing
    Set viewModel = Nothing
    Set sourceModel = Nothing
    Set sourceInstance = Nothing
    modelComponents = Empty
    drawingRootName = "": drawingRootAlias = ""
    Debug.Print LOG_PREFIX & " START " & CStr(Now)
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Err.Raise vbObjectError + 1, , "Open a drawing first."
    If doc.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."

    Dim sm As SldWorks.SelectionMgr
    Set sm = doc.SelectionManager
    If sm.GetSelectedObjectCount2(-1) <> 1 Then Err.Raise vbObjectError + 3, , "Select exactly one drawing view (its border)."
    If sm.GetSelectedObjectType3(1, -1) <> swSelDRAWINGVIEWS Then Err.Raise vbObjectError + 4, , "Select the view border, not an edge."
    Set targetView = sm.GetSelectedObject6(1, -1)
    If targetView Is Nothing Then Err.Raise vbObjectError + 5, , "Cannot obtain selected view."
    If CBool(targetView.ProjectedDimensions) Then Err.Raise vbObjectError + 6, , "Set this view's dimension type to True before running. Projected dimensions are not supported."

    viewOutline = targetView.GetOutline
    If Not IsArray(viewOutline) Then Err.Raise vbObjectError + 19, , "View outline unavailable."
    If UBound(viewOutline) - LBound(viewOutline) <> 3 Then Err.Raise vbObjectError + 19, , "Invalid view outline."
    If viewOutline(2) <= viewOutline(0) Or viewOutline(3) <= viewOutline(1) Then _
        Err.Raise vbObjectError + 19, , "Empty view outline."
    Dim referenced As SldWorks.ModelDoc2
    Set referenced = targetView.ReferencedDocument
    If referenced Is Nothing Then Err.Raise vbObjectError + 7, , "Referenced model unavailable. Open/resolve the model first."
    Set viewModel = referenced
    assemblyView = (referenced.GetType = swDocASSEMBLY)
    If Not assemblyView And referenced.GetType <> swDocPART Then _
        Err.Raise vbObjectError + 8, , "Only part and assembly views are supported."
    Dim drawing As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Set drawing = doc
    Set sheet = drawing.GetCurrentSheet
    sheet.GetSize sheetWidth, sheetHeight
    If sheetWidth <= 0 Or sheetHeight <= 0 Then Err.Raise vbObjectError + 9, , "Invalid sheet size."
    Debug.Print LOG_PREFIX & ": existing dimensions preserved; rerun can create duplicates"
    SeedOccupied
    CheckSourceConfiguration referenced, targetView.ReferencedConfiguration, "view model"
    If assemblyView Then
        Dim assembly As SldWorks.AssemblyDoc
        Set assembly = referenced
        modelComponents = assembly.GetComponents(False)
        If Not HasItems(modelComponents) Then Err.Raise vbObjectError + 33, , "Referenced assembly has no accessible model components"
        CaptureDrawingRoot
    End If
    Dim comps As Variant
    comps = targetView.GetVisibleComponents

    oldInputPreference = CBool(swApp.GetUserPreferenceToggle(swInputDimValOnCreate))
    preferenceCaptured = True
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, False
    If CBool(swApp.GetUserPreferenceToggle(swInputDimValOnCreate)) Then _
        Err.Raise vbObjectError + 17, , "Cannot disable dimension value input preference."

    Dim i As Long, comp As SldWorks.Component2
    Debug.Print LOG_PREFIX & " view: " & targetView.GetName2
    If assemblyView Then
        If Not HasItems(comps) Then Err.Raise vbObjectError + 10, , "No visible components. Resolve the assembly first."
        For i = LBound(comps) To UBound(comps)
            Set comp = comps(i)
            If Not comp Is Nothing Then ProcessComponent comp
            If blocked Then Exit For
        Next i
    Else
        ' A part view can expose a root component; it is not an assembly transform.
        Set comp = Nothing
        If HasItems(comps) Then Set comp = comps(LBound(comps))
        If comp Is Nothing Then
            Dim root As SldWorks.DrawingComponent
            Set root = targetView.RootDrawingComponent
            If Not root Is Nothing Then Set comp = root.Component
        End If
        If comp Is Nothing Then Err.Raise vbObjectError + 31, , "Part view root component unavailable. Reopen/resolve the drawing and part."
        ProcessComponent comp
    End If
    doc.ClearSelection2 True
    doc.GraphicsRedraw2
    RestorePreference
    If blocked Then MsgBox LOG_PREFIX & ": STOP - dimension creation or cleanup has an uncertain outcome. Inspect this drawing manually before rerunning.", vbCritical
    Debug.Print LOG_PREFIX & " END created=" & created & "; skipped=" & skipped & "; failed=" & failed & "; bodies=" & bodiesSeen & "; unresolved=" & unresolved & "; blocked=" & blocked
    MsgBox LOG_PREFIX & ": created=" & created & "; skipped=" & skipped & "; failed=" & failed & vbCrLf & _
        "Bodies=" & bodiesSeen & "; unresolved=" & unresolved & vbCrLf & _
        "Details: VBA Immediate window (Ctrl+G). Review dimensions before saving.", vbInformation
    Exit Sub
Fatal:
    Dim message As String
    message = Err.Description
    RestorePreference
    ClearSelectionSafely
    Debug.Print LOG_PREFIX & " STOP: " & message
    MsgBox LOG_PREFIX & " stopped: " & message, vbExclamation
End Sub

Private Sub RestorePreference()
    On Error GoTo RestoreFailed
    If preferenceCaptured Then
        swApp.SetUserPreferenceToggle swInputDimValOnCreate, oldInputPreference
        If CBool(swApp.GetUserPreferenceToggle(swInputDimValOnCreate)) <> oldInputPreference Then _
            Err.Raise vbObjectError + 18, , "Preference restore failed"
        preferenceCaptured = False
    End If
    Exit Sub
RestoreFailed:
    Debug.Print LOG_PREFIX & ": preference restore failed: " & Err.Description
    MsgBox LOG_PREFIX & ": Could not restore Input dimension value preference. Check SolidWorks options.", vbExclamation
End Sub

Private Sub ClearSelectionSafely()
    On Error GoTo ClearFailed
    If Not doc Is Nothing Then doc.ClearSelection2 True
    Exit Sub
ClearFailed:
    Debug.Print LOG_PREFIX & ": could not clear selection: " & Err.Description
End Sub

Private Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasItems = (UBound(values) >= LBound(values))
EmptyArray:
End Function

Private Function ComponentName(ByVal comp As SldWorks.Component2) As String
    ComponentName = "[part]"
    If Not comp Is Nothing Then ComponentName = comp.Name2
End Function

Private Sub CheckSourceConfiguration(ByVal model As SldWorks.ModelDoc2, _
        ByVal requestedConfiguration As String, ByVal context As String)
    Dim active As SldWorks.Configuration, sourceConfiguration As String
    Set active = model.ConfigurationManager.ActiveConfiguration
    If active Is Nothing Then Err.Raise vbObjectError + 32, , "Source active configuration unavailable: " & context
    sourceConfiguration = active.Name
    If Len(Trim$(sourceConfiguration)) = 0 Then _
        Err.Raise vbObjectError + 32, , "Source active configuration name unavailable: " & context
    If Len(Trim$(requestedConfiguration)) = 0 Then
        ' An empty API result is UNKNOWN, not a different configuration name.
        ' Do not switch configurations or assume an empty string means default.
        ' Assembly source edges must still pass SourceEdge/completeMaps, and
        ' every newly created dimension must pass the true-length readback.
        Debug.Print LOG_PREFIX & " " & context & ": referenced configuration not reported; source=[" & _
            sourceConfiguration & "]; proceeding with geometric validation"
    ElseIf StrComp(sourceConfiguration, requestedConfiguration, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 32, , "Configuration mismatch: source=[" & sourceConfiguration & _
            "]; referenced=[" & requestedConfiguration & "]; context=" & context & _
            ". Activate the referenced configuration in the source model first."
    End If
End Sub

Private Sub CaptureDrawingRoot()
    Dim drawingRoot As SldWorks.DrawingComponent, rootComponent As SldWorks.Component2
    Set drawingRoot = targetView.RootDrawingComponent
    If Not drawingRoot Is Nothing Then
        drawingRootName = drawingRoot.Name
        Set rootComponent = drawingRoot.Component
        If Not rootComponent Is Nothing Then drawingRootAlias = rootComponent.Name2
    End If
    Debug.Print LOG_PREFIX & ": drawing root=[" & drawingRootName & "]; root component=[" & drawingRootAlias & "]"
    Debug.Print LOG_PREFIX & ": model component count=" & (UBound(modelComponents) - LBound(modelComponents) + 1)
    Dim i As Long, candidate As SldWorks.Component2
    For i = LBound(modelComponents) To UBound(modelComponents)
        If i - LBound(modelComponents) >= 8 Then Exit For
        Set candidate = modelComponents(i)
        If Not candidate Is Nothing Then Debug.Print LOG_PREFIX & ": model tree sample=[" & candidate.Name2 & "]"
    Next i
End Sub

Private Function ComponentNameMatches(ByVal requested As String, ByVal actual As String, _
        ByVal rootName As String, ByVal rootAlias As String) As Boolean
    ' Keep actual intact. Compare both legitimate drawing representations.
    ' Never strip arbitrary ancestors, match basenames or discard nesting.
    ComponentNameMatches = (StrComp(requested, actual, vbTextCompare) = 0)
    If Len(rootName) > 0 Then
        ComponentNameMatches = ComponentNameMatches Or (StrComp(requested, rootName & "/" & actual, vbTextCompare) = 0)
    End If
    If Len(rootAlias) > 0 Then
        ComponentNameMatches = ComponentNameMatches Or (StrComp(requested, rootAlias & "/" & actual, vbTextCompare) = 0)
    End If
End Function

Private Function ComponentMetadataMatches(ByVal candidate As SldWorks.Component2, ByVal requestedPath As String, _
        ByVal requestedConfiguration As String, ByRef pathRejected As Long, ByRef configRejected As Long) As Boolean
    If Len(requestedPath) > 0 Then
        If StrComp(Replace(candidate.GetPathName, "/", "\"), Replace(requestedPath, "/", "\"), vbTextCompare) <> 0 Then
            pathRejected = pathRejected + 1
            Debug.Print LOG_PREFIX & ": rejected model path: instance=[" & candidate.Name2 & "]; actual=[" & _
                candidate.GetPathName & "]; drawing=[" & requestedPath & "]"
            Exit Function
        End If
    End If
    If Len(Trim$(requestedConfiguration)) > 0 Then
        If StrComp(candidate.ReferencedConfiguration, requestedConfiguration, vbTextCompare) <> 0 Then
            configRejected = configRejected + 1
            Debug.Print LOG_PREFIX & ": rejected model configuration: instance=[" & candidate.Name2 & "]; actual=[" & _
                candidate.ReferencedConfiguration & "]; drawing=[" & requestedConfiguration & "]"
            Exit Function
        End If
    End If
    ComponentMetadataMatches = True
End Function

Private Function ResolveModelComponent(ByVal drawingComponent As SldWorks.Component2) As SldWorks.Component2
    Dim normalized As Object, candidate As SldWorks.Component2, found As SldWorks.Component2
    Dim i As Long, matches As Long, requestedPath As String, requestedConfiguration As String
    Dim requestedName As String, nameMatches As Long, pathRejected As Long, configRejected As Long
    requestedName = drawingComponent.Name2
    requestedPath = drawingComponent.GetPathName
    requestedConfiguration = drawingComponent.ReferencedConfiguration
    Set normalized = TryCorrespondingObject(viewModel, drawingComponent)
    If Not normalized Is Nothing Then
        For i = LBound(modelComponents) To UBound(modelComponents)
            Set candidate = modelComponents(i)
            If Not candidate Is Nothing Then
                If swApp.IsSame(normalized, candidate) = swObjectSame Then
                    If ComponentMetadataMatches(candidate, requestedPath, requestedConfiguration, pathRejected, configRejected) Then
                        Set ResolveModelComponent = candidate
                        Debug.Print LOG_PREFIX & ": resolved by API: drawing=[" & requestedName & "]; model=[" & candidate.Name2 & "]"
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    pathRejected = 0: configRejected = 0
    ' Count the UNION of exact and verified-root-prefixed matches once per
    ' actual instance. A root-name collision must remain ambiguous, not first-win.
    For i = LBound(modelComponents) To UBound(modelComponents)
        Set candidate = modelComponents(i)
        If Not candidate Is Nothing Then
            If ComponentNameMatches(requestedName, candidate.Name2, drawingRootName, drawingRootAlias) Then
                nameMatches = nameMatches + 1
                If ComponentMetadataMatches(candidate, requestedPath, requestedConfiguration, pathRejected, configRejected) Then
                    matches = matches + 1
                    Set found = candidate
                End If
            End If
        End If
    Next i
    Debug.Print LOG_PREFIX & ": component resolution: drawing=[" & requestedName & "]; name matches=" & nameMatches & _
        "; path rejected=" & pathRejected & "; config rejected=" & configRejected & "; accepted=" & matches
    If matches <> 1 Then Err.Raise vbObjectError + 34, , _
        "Cannot uniquely resolve drawing component: " & requestedName & "; matches=" & matches & _
        "; name matches=" & nameMatches & "; path rejected=" & pathRejected & "; config rejected=" & configRejected
    Set ResolveModelComponent = found
    Debug.Print LOG_PREFIX & ": resolved by full hierarchy: drawing=[" & requestedName & "]; model=[" & found.Name2 & "]"
End Function

Private Function TryCorrespondingObject(ByVal model As SldWorks.ModelDoc2, ByVal item As Object) As Object
    On Error GoTo NoCorrespondence
    Set TryCorrespondingObject = model.Extension.GetCorresponding2(item)
    Exit Function
NoCorrespondence:
    Debug.Print LOG_PREFIX & ": model corresponding-object API error " & Err.Number & ": " & Err.Description
End Function

Private Sub BuildLocalLookup(ByVal comp As SldWorks.Component2)
    Set localEdges = New Collection
    Set instanceEdges = New Collection
    Set assemblyEdges = New Collection
    Set sourceInstance = Nothing
    reverseMatches = 0: forwardMatches = 0: missingMatches = 0
    correspondenceErrors = 0: forwardBuilt = False
    reverseErrors = 0: assemblyErrors = 0: viewErrors = 0
    reverseNulls = 0: reverseWrongTypes = 0: reverseNonLocal = 0
    If assemblyView Then
        If comp Is Nothing Then Err.Raise vbObjectError + 20, , "Missing drawing component"
        Set sourceInstance = ResolveModelComponent(comp)
        Set sourceModel = sourceInstance.GetModelDoc2
        If sourceModel Is Nothing Then Err.Raise vbObjectError + 21, , "Resolve the component model first"
        Debug.Print LOG_PREFIX & " model instance=" & sourceInstance.Name2 & "; config=[" & sourceInstance.ReferencedConfiguration & "]"
        CheckSourceConfiguration sourceModel, sourceInstance.ReferencedConfiguration, sourceInstance.Name2
    Else
        Set sourceModel = viewModel
    End If
    If sourceModel.GetType <> swDocPART Then Err.Raise vbObjectError + 22, , "Expected a leaf part model"
    Dim part As SldWorks.PartDoc
    Set part = sourceModel
    Dim bs As Variant, es As Variant, i As Long, j As Long
    Dim body As SldWorks.Body2, edge As SldWorks.Edge
    bs = part.GetBodies2(swAllBodies, False)
    If Not HasItems(bs) Then Err.Raise vbObjectError + 23, , "Source part has no accessible bodies"
    For i = LBound(bs) To UBound(bs)
        Set body = bs(i)
        es = body.GetEdges
        If HasItems(es) Then
            For j = LBound(es) To UBound(es)
                Set edge = es(j)
                localEdges.Add edge
            Next j
        End If
    Next i
    Debug.Print LOG_PREFIX & ": source edges=" & localEdges.Count
End Sub

Private Function LocalMember(ByVal item As Object) As SldWorks.Edge
    If item Is Nothing Then Exit Function
    If Not TypeOf item Is SldWorks.Edge Then Exit Function
    Dim edge As SldWorks.Edge, i As Long
    For i = 1 To localEdges.Count
        Set edge = localEdges(i)
        If swApp.IsSame(item, edge) = swObjectSame Then
            Set LocalMember = edge
            Exit Function
        End If
    Next i
End Function

Private Function ReverseCorrespondence(ByVal visibleEdge As SldWorks.Edge) As Object
    On Error GoTo ConversionFailed
    ' The receiver is the TARGET source document. Do not compare the converted
    ' source entity's identity directly with the drawing-context entity.
    Set ReverseCorrespondence = sourceModel.Extension.GetCorrespondingEntity2(visibleEdge)
    Exit Function
ConversionFailed:
    correspondenceErrors = correspondenceErrors + 1
    reverseErrors = reverseErrors + 1
    If reverseErrors = 1 Then Debug.Print LOG_PREFIX & ": source entity conversion error " & Err.Number & ": " & Err.Description
End Function

Private Sub BuildForwardLookup()
    ' Exact fallback: local -> actual assembly instance -> selected drawing view.
    Dim edge As SldWorks.Edge, inAssembly As Object, inView As Object, i As Long
    Dim assemblyCount As Long, viewCount As Long
    For i = 1 To localEdges.Count
        Set edge = localEdges(i)
        Set inAssembly = ForwardAssemblyEntity(edge)
        Set inView = Nothing
        If Not inAssembly Is Nothing Then
            assemblyCount = assemblyCount + 1
            Set inView = ForwardViewEntity(inAssembly)
        End If
        If Not inView Is Nothing Then viewCount = viewCount + 1
        assemblyEdges.Add inAssembly
        instanceEdges.Add inView
    Next i
    forwardBuilt = True
    Debug.Print LOG_PREFIX & ": forward lookup: source=" & localEdges.Count & _
        "; assembly/model=" & assemblyCount & "; drawing view=" & viewCount
End Sub

Private Function ForwardAssemblyEntity(ByVal edge As SldWorks.Edge) As Object
    On Error GoTo ConversionFailed
    If assemblyView Then
        Set ForwardAssemblyEntity = sourceInstance.GetCorresponding(edge)
    Else
        Set ForwardAssemblyEntity = edge
    End If
    Exit Function
ConversionFailed:
    correspondenceErrors = correspondenceErrors + 1
    assemblyErrors = assemblyErrors + 1
    If assemblyErrors = 1 Then Debug.Print LOG_PREFIX & ": assembly entity conversion error " & Err.Number & ": " & Err.Description
End Function

Private Function ForwardViewEntity(ByVal item As Object) As Object
    On Error GoTo ConversionFailed
    Set ForwardViewEntity = targetView.GetCorresponding(item)
    Exit Function
ConversionFailed:
    correspondenceErrors = correspondenceErrors + 1
    viewErrors = viewErrors + 1
    If viewErrors = 1 Then Debug.Print LOG_PREFIX & ": view entity conversion error " & Err.Number & ": " & Err.Description
End Function

Private Function SourceEdge(ByVal visibleEdge As SldWorks.Edge) As SldWorks.Edge
    Dim localEdge As SldWorks.Edge, converted As Object
    Set converted = ReverseCorrespondence(visibleEdge)
    If converted Is Nothing Then
        reverseNulls = reverseNulls + 1
    ElseIf Not TypeOf converted Is SldWorks.Edge Then
        reverseWrongTypes = reverseWrongTypes + 1
        If reverseWrongTypes = 1 Then Debug.Print LOG_PREFIX & ": reverse conversion unexpected type=" & TypeName(converted)
    Else
        Set localEdge = LocalMember(converted)
        If localEdge Is Nothing Then
            reverseNonLocal = reverseNonLocal + 1
            If reverseNonLocal = 1 Then Debug.Print LOG_PREFIX & ": reverse edge is not in active source-part geometry"
        End If
    End If
    If localEdge Is Nothing Then Set localEdge = LocalMember(visibleEdge)
    If Not localEdge Is Nothing Then
        reverseMatches = reverseMatches + 1
        Set SourceEdge = localEdge
        Exit Function
    End If
    If Not forwardBuilt Then BuildForwardLookup
    Dim i As Long, mapped As Object
    For i = 1 To localEdges.Count
        Set mapped = instanceEdges(i)
        If Not mapped Is Nothing Then
            If swApp.IsSame(visibleEdge, mapped) = swObjectSame Then
                forwardMatches = forwardMatches + 1
                Set SourceEdge = localEdges(i)
                Exit Function
            End If
        End If
        Set mapped = assemblyEdges(i)
        If Not mapped Is Nothing Then
            If swApp.IsSame(visibleEdge, mapped) = swObjectSame Then
                forwardMatches = forwardMatches + 1
                Set SourceEdge = localEdges(i)
                Exit Function
            End If
        End If
    Next i
    missingMatches = missingMatches + 1
End Function

Private Function IsStraightEdge(ByVal edge As SldWorks.Edge) As Boolean
    ' Propagate unexpected API errors to the component handler. A curved edge
    ' is not a missing straight-length candidate and must not block its body.
    Dim curve As SldWorks.Curve
    Set curve = edge.GetCurve
    If curve Is Nothing Then Err.Raise vbObjectError + 35, , "Visible edge curve unavailable"
    IsStraightEdge = CBool(curve.IsLine)
End Function

Private Function StraightLength(ByVal edge As SldWorks.Edge, ByRef length As Double, _
        ByRef reason As String) As Boolean
    On Error GoTo BadEdge
    reason = "edge/curve/vertices unavailable"
    If edge Is Nothing Then Exit Function
    Dim curve As SldWorks.Curve, va As SldWorks.Vertex, vb As SldWorks.Vertex
    Set curve = edge.GetCurve
    If curve Is Nothing Then Exit Function
    reason = "not straight"
    If Not CBool(curve.IsLine) Then Exit Function
    Set va = edge.GetStartVertex: Set vb = edge.GetEndVertex
    reason = "vertices unavailable"
    If va Is Nothing Then Exit Function
    If vb Is Nothing Then Exit Function
    Dim a As Variant, b As Variant
    a = va.GetPoint: b = vb.GetPoint
    length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2 + (b(2) - a(2)) ^ 2)
    reason = "degenerate edge"
    If length <= MIN_LENGTH_M Then Exit Function
    StraightLength = True
    reason = ""
    Exit Function
BadEdge:
    reason = "geometry error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub ProcessComponent(ByVal comp As SldWorks.Component2)
    On Error GoTo FailedComponent
    Dim edges As Variant, componentLabel As String
    componentLabel = ComponentName(comp)
    Debug.Print LOG_PREFIX & " component=" & componentLabel
    BuildLocalLookup comp
    edges = targetView.GetVisibleEntities2(comp, swViewEntityType_Edge)
    If Not HasItems(edges) Then
        skipped = skipped + 1
        Debug.Print LOG_PREFIX & " " & componentLabel & ": no accessible visible edges"
        Exit Sub
    End If
    Dim bodies As New Collection, edgeLists As New Collection, maxLengths As New Collection
    Dim completeMaps As New Collection
    Dim i As Long, j As Long, index As Long, length As Double, reason As String
    Dim edge As SldWorks.Edge, localEdge As SldWorks.Edge
    Dim body As SldWorks.Body2, oldBody As SldWorks.Body2, candidates As Collection
    For i = LBound(edges) To UBound(edges)
        If TypeOf edges(i) Is SldWorks.Edge Then
            Set edge = edges(i)
            Set body = edge.GetBody
            If Not body Is Nothing Then
                index = 0
                For j = 1 To bodies.Count
                    Set oldBody = bodies(j)
                    If swApp.IsSame(body, oldBody) = swObjectSame Then
                        index = j
                        Exit For
                    End If
                Next j
                If index = 0 Then
                    bodies.Add body
                    Set candidates = New Collection
                    edgeLists.Add candidates
                    maxLengths.Add 0#
                    completeMaps.Add True
                    index = bodies.Count
                End If
                If IsStraightEdge(edge) Then
                    Set localEdge = SourceEdge(edge)
                    If localEdge Is Nothing Then
                        completeMaps.Add False, Before:=index
                        completeMaps.Remove index + 1
                        Debug.Print LOG_PREFIX & " " & componentLabel & " body=" & index & _
                            " edge=" & i & ": SOURCE MAPPING FAILED after source/view context conversion"
                    ElseIf StraightLength(localEdge, length, reason) Then
                        Set candidates = edgeLists(index)
                        If length > CDbl(maxLengths(index)) + LENGTH_TOL_M Then
                            ' Keep only alternatives of the SAME maximum length.
                            Do While candidates.Count > 0
                                candidates.Remove candidates.Count
                            Loop
                            maxLengths.Add length, Before:=index
                            maxLengths.Remove index + 1
                        End If
                        If Abs(length - CDbl(maxLengths(index))) <= LENGTH_TOL_M Then candidates.Add edge
                    Else
                        completeMaps.Add False, Before:=index
                        completeMaps.Remove index + 1
                        Debug.Print LOG_PREFIX & " " & componentLabel & " edge=" & i & " unusable straight geometry: " & reason
                    End If
                End If
            Else
                Debug.Print LOG_PREFIX & " " & componentLabel & " edge=" & i & ": body unavailable"
            End If
        End If
    Next i
    Debug.Print LOG_PREFIX & " " & componentLabel & ": straight-edge mapping: reverse/local=" & reverseMatches & _
        "; forward=" & forwardMatches & "; missing=" & missingMatches & "; API errors=" & correspondenceErrors
    Debug.Print LOG_PREFIX & ": reverse results: Nothing/error=" & reverseNulls & "; wrong type=" & reverseWrongTypes & _
        "; not in source geometry=" & reverseNonLocal
    For i = 1 To bodies.Count
        bodiesSeen = bodiesSeen + 1
        Set candidates = edgeLists(i)
        Debug.Print LOG_PREFIX & " " & componentLabel & " body=" & i & "; maximum mm=" & _
            Format$(CDbl(maxLengths(i)) * 1000#, "0.######") & "; alternatives=" & candidates.Count
        Dim done As Boolean
        done = False
        If Not CBool(completeMaps(i)) Then
            Debug.Print LOG_PREFIX & " " & componentLabel & " body=" & i & ": incomplete source mapping; refusing a potentially short substitute"
        Else
            For j = 1 To candidates.Count
                Set edge = candidates(j)
                If AddEdgeDimension(edge, comp) Then
                    done = True
                    Exit For
                End If
                If blocked Then Exit For
            Next j
        End If
        If Not done Then
            unresolved = unresolved + 1
            Debug.Print LOG_PREFIX & " " & componentLabel & " body=" & i & ": UNRESOLVED; no shorter substitute used"
        End If
        If blocked Then Exit For
    Next i
    Exit Sub
FailedComponent:
    failed = failed + 1
    Debug.Print LOG_PREFIX & " " & componentLabel & ": component failed: " & Err.Description
End Sub

Private Function Candidate(ByVal edge As SldWorks.Edge, ByVal comp As SldWorks.Component2, ByRef length As Double, _
        ByRef a As Variant, ByRef b As Variant, ByRef reason As String) As Boolean
    On Error GoTo Unusable
    Dim localEdge As SldWorks.Edge
    Set localEdge = SourceEdge(edge)
    reason = "cannot map visible edge to source part"
    If localEdge Is Nothing Then Exit Function
    If Not StraightLength(localEdge, length, reason) Then Exit Function
    Dim va As SldWorks.Vertex, vb As SldWorks.Vertex
    Set va = localEdge.GetStartVertex: Set vb = localEdge.GetEndVertex
    a = va.GetPoint: b = vb.GetPoint
    reason = "source-part to sheet projection failed"
    If Not ProjectEndpoints(a, b, comp, a, b) Then Exit Function
    reason = "projected endpoints outside selected view; unsupported view/transform"
    If Not PointsInsideOutline(a, b) Then Exit Function
    reason = "edge is end-on / projected length too small"
    If Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2) < MIN_PROJECTED_M Then Exit Function
    Candidate = True
    reason = ""
    Exit Function
Unusable:
    reason = reason & ": VBA error " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Function ProjectEndpoints(ByVal a As Variant, ByVal b As Variant, ByVal comp As SldWorks.Component2, _
        ByRef outA As Variant, ByRef outB As Variant) As Boolean
    On Error GoTo ProjectionFailed
    Dim mu As SldWorks.MathUtility, pa As SldWorks.MathPoint, pb As SldWorks.MathPoint
    Dim transform As SldWorks.MathTransform
    Set mu = swApp.GetMathUtility
    Set pa = mu.CreatePoint(a): Set pb = mu.CreatePoint(b)
    If assemblyView Then
        If sourceInstance Is Nothing Then Exit Function
        Set transform = sourceInstance.Transform2
        If transform Is Nothing Then Exit Function
        Set pa = pa.MultiplyTransform(transform)
        Set pb = pb.MultiplyTransform(transform)
    End If
    Set transform = targetView.ModelToViewTransform
    If transform Is Nothing Then Exit Function
    Set pa = pa.MultiplyTransform(transform)
    Set pb = pb.MultiplyTransform(transform)
    outA = pa.ArrayData: outB = pb.ArrayData
    ProjectEndpoints = True
    Exit Function
ProjectionFailed:
    ProjectEndpoints = False
End Function

Private Function PointsInsideOutline(ByVal a As Variant, ByVal b As Variant) As Boolean
    PointsInsideOutline = (a(0) >= viewOutline(0) - OUTLINE_TOL_M And a(0) <= viewOutline(2) + OUTLINE_TOL_M And _
        a(1) >= viewOutline(1) - OUTLINE_TOL_M And a(1) <= viewOutline(3) + OUTLINE_TOL_M And _
        b(0) >= viewOutline(0) - OUTLINE_TOL_M And b(0) <= viewOutline(2) + OUTLINE_TOL_M And _
        b(1) >= viewOutline(1) - OUTLINE_TOL_M And b(1) <= viewOutline(3) + OUTLINE_TOL_M)
End Function

Private Sub SeedOccupied()
    ' Preserve all existing dimensions. Reserve conservative text rectangles.
    Dim dd As SldWorks.DisplayDimension, ann As SldWorks.Annotation, p As Variant
    Set dd = targetView.GetFirstDisplayDimension5
    Do While Not dd Is Nothing
        Set ann = dd.GetAnnotation
        If Not ann Is Nothing Then
            p = ann.GetPosition
            If HasItems(p) Then occupied.Add Array(CDbl(p(0)), CDbl(p(1)), 0.015, TEXT_HEIGHT_M)
        End If
        Set dd = dd.GetNext5
    Loop
End Sub

Private Function TextWidth(ByVal length As Double, ByVal height As Double) As Double
    ' Conservative estimate, not an API-provided text bounding box.
    ' Return HALF width: estimate digits plus allowance for unit/rounding.
    TextWidth = (Len(Format$(length * 1000#, "0.###")) + 4) * height * 0.4 + LAYOUT_GAP_M
    If TextWidth < 0.006 Then TextWidth = 0.006
End Function

Private Function PositionFits(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double) As Boolean
    ' width/height are estimated HALF extents around the text origin.
    ' This is a layout estimate, not a measured glyph/extension-line bounding box.
    If x - width < viewOutline(0) Or x + width > viewOutline(2) Then Exit Function
    If y - height < viewOutline(1) Or y + height > viewOutline(3) Then Exit Function
    If x - width < LAYOUT_GAP_M Or x + width > sheetWidth - LAYOUT_GAP_M Then Exit Function
    If y - height < LAYOUT_GAP_M Or y + height > sheetHeight - LAYOUT_GAP_M Then Exit Function
    Dim r As Variant
    For Each r In occupied
        If Abs(x - CDbl(r(0))) < width + CDbl(r(2)) + LAYOUT_GAP_M And _
           Abs(y - CDbl(r(1))) < height + CDbl(r(3)) + LAYOUT_GAP_M Then Exit Function
    Next r
    PositionFits = True
End Function

Private Function FindPosition(ByVal a As Variant, ByVal b As Variant, ByVal width As Double, ByVal height As Double, _
        ByRef x As Double, ByRef y As Double) As Boolean
    Dim dx As Double, dy As Double, projected As Double
    dx = b(0) - a(0): dy = b(1) - a(1)
    projected = Sqr(dx * dx + dy * dy)
    If projected < MIN_PROJECTED_M Then Exit Function
    ' Search only near the measured edge; never send text to an arbitrary sheet corner.
    Dim level As Long, side As Long, shift As Long, t As Double, offset As Double
    For level = 0 To 4
        offset = OFFSET_M + level * 0.002
        For shift = 0 To 2
            t = 0.5
            If shift = 1 Then t = 0.35
            If shift = 2 Then t = 0.65
            For side = -1 To 1 Step 2
                x = a(0) + t * dx - side * dy / projected * offset
                y = a(1) + t * dy + side * dx / projected * offset
                If PositionFits(x, y, width, height) Then
                    FindPosition = True
                    Exit Function
                End If
            Next side
        Next shift
    Next level
End Function

Private Function AddEdgeDimension(ByVal edge As SldWorks.Edge, ByVal comp As SldWorks.Component2) As Boolean
    On Error GoTo FailedDimension
    Dim length As Double, x As Double, y As Double, reason As String, componentLabel As String
    Dim a As Variant, b As Variant, width As Double, height As Double
    Dim dd As SldWorks.DisplayDimension, creationAttempted As Boolean
    Dim ann As SldWorks.Annotation, dimValue As SldWorks.Dimension
    componentLabel = ComponentName(comp)
    If Not Candidate(edge, comp, length, a, b, reason) Then Err.Raise vbObjectError + 12, , reason
    height = TEXT_HEIGHT_M
    width = TextWidth(length, height)
    If Not FindPosition(a, b, width, height, x, y) Then Err.Raise vbObjectError + 24, , "No free near-edge text slot inside this view; enlarge view or arrange manually"
    doc.ClearSelection2 True
    If Not CBool(targetView.SelectEntity(edge, False)) Then Err.Raise vbObjectError + 13, , "Cannot select edge in view"
    creationAttempted = True
    Set dd = doc.AddDimension2(x, y, 0#)
    If dd Is Nothing Then Err.Raise vbObjectError + 14, , "AddDimension2 returned Nothing"
    Set dimValue = dd.GetDimension2(0)
    If dimValue Is Nothing Then Err.Raise vbObjectError + 15, , "Dimension readback unavailable"
    Dim actual As Double
    actual = dimValue.SystemValue
    If Abs(actual - length) > LENGTH_TOL_M Then Err.Raise vbObjectError + 16, , _
        "Not true edge length: expected mm=" & CStr(length * 1000#) & "; actual mm=" & CStr(actual * 1000#)
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Err.Raise vbObjectError + 25, , "Dimension annotation unavailable"
    Dim fmt As SldWorks.TextFormat
    Set fmt = ann.GetTextFormat(0)
    If Not fmt Is Nothing Then
        If fmt.CharHeight > height Then height = fmt.CharHeight
    End If
    width = TextWidth(length, height)
    If Not FindPosition(a, b, width, height, x, y) Then Err.Raise vbObjectError + 24, , "No text slot for actual font height"
    If Not CBool(ann.SetPosition2(x, y, 0#)) Then Err.Raise vbObjectError + 26, , "Cannot move dimension annotation"
    Dim p As Variant
    p = ann.GetPosition
    If Not HasItems(p) Then Err.Raise vbObjectError + 27, , "Annotation position readback unavailable"
    If Abs(CDbl(p(0)) - x) > POSITION_TOL_M Or Abs(CDbl(p(1)) - y) > POSITION_TOL_M Then _
        Err.Raise vbObjectError + 28, , "SolidWorks did not retain requested annotation position"
    If Not PositionFits(CDbl(p(0)), CDbl(p(1)), width, height) Then _
        Err.Raise vbObjectError + 29, , "Readback position conflicts with placement limits"
    doc.ClearSelection2 True
    occupied.Add Array(CDbl(p(0)), CDbl(p(1)), width, height)
    created = created + 1
    AddEdgeDimension = True
    Debug.Print LOG_PREFIX & " " & componentLabel & ": CREATED mm=" & Format$(actual * 1000#, "0.######") & _
        "; sheet text x/y mm=" & CStr(p(0) * 1000#) & "/" & CStr(p(1) * 1000#)
    Exit Function
FailedDimension:
    reason = Err.Description
    failed = failed + 1
    Debug.Print LOG_PREFIX & " " & componentLabel & ": attempt rejected: " & reason
    If Not dd Is Nothing Then
        If Not RemoveNewDimension(dd) Then
            blocked = True
            Debug.Print LOG_PREFIX & ": STOP: failed to remove current attempt; inspect manually"
        End If
    ElseIf creationAttempted Then
        blocked = True
        Debug.Print LOG_PREFIX & ": STOP: creation outcome uncertain; inspect manually"
    End If
    ClearSelectionSafely
End Function

Private Function RemoveNewDimension(ByVal dd As SldWorks.DisplayDimension) As Boolean
    On Error GoTo CannotRemove
    Dim ann As SldWorks.Annotation
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Exit Function
    doc.ClearSelection2 True
    If CBool(ann.Select3(False, Nothing)) = False Then Exit Function
    RemoveNewDimension = CBool(doc.Extension.DeleteSelection2(0))
CannotRemove:
End Function
