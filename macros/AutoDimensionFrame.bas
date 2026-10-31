Attribute VB_Name = "AutoDimensionFrame"
Option Explicit

' Requires SolidWorks type libraries (normal SolidWorks VBA macro project).
' Distances and coordinates are in metres. No document is saved automatically.
Private Const MACRO_VERSION As String = "2.2"
Private Const LOG_PREFIX As String = "AutoFrameDim v" & MACRO_VERSION
Private Const OFFSET_M As Double = 0.010
Private Const OUTLINE_TOL_M As Double = 0.001
Private Const COORD_EQUAL_TOL_M As Double = 0.0001
Private Const MIN_LENGTH_M As Double = 0.02
Private Const MIN_PROJECTED_M As Double = 0.001
Private Const LENGTH_TOL_M As Double = 0.000001

Private swApp As SldWorks.SldWorks
Private doc As SldWorks.ModelDoc2
Private targetView As SldWorks.View
Private created As Long, skipped As Long, failed As Long
Private blocked As Boolean
Private preferenceCaptured As Boolean, oldInputPreference As Boolean
Private viewOutline As Variant
Private coordinateRejectLogCount As Long

Public Sub main()
    On Error GoTo Fatal
    preferenceCaptured = False
    created = 0: skipped = 0: failed = 0: blocked = False
    Set doc = Nothing
    Set targetView = Nothing
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
    Dim viewPosition As Variant
    viewPosition = targetView.Position
    Debug.Print LOG_PREFIX & " outline mm: " & SheetPoint(viewOutline(0), viewOutline(1)) & " to " & SheetPoint(viewOutline(2), viewOutline(3))
    Debug.Print LOG_PREFIX & " view position mm: " & SheetPoint(viewPosition(0), viewPosition(1)) & "; scale=" & targetView.ScaleDecimal & "; offset mm=" & OFFSET_M * 1000#

    Dim comps As Variant
    comps = targetView.GetVisibleComponents
    If IsEmpty(comps) Then Err.Raise vbObjectError + 7, , "No visible assembly components. Part-only views are not supported."
    If Not IsArray(comps) Then Err.Raise vbObjectError + 8, , "Visible component array unavailable."

    oldInputPreference = CBool(swApp.GetUserPreferenceToggle(swInputDimValOnCreate))
    preferenceCaptured = True
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, False
    If CBool(swApp.GetUserPreferenceToggle(swInputDimValOnCreate)) Then _
        Err.Raise vbObjectError + 17, , "Cannot disable dimension value input preference."

    Dim i As Long, comp As SldWorks.Component2
    Debug.Print LOG_PREFIX & " view: " & targetView.GetName2
    For i = LBound(comps) To UBound(comps)
        Set comp = comps(i)
        If Not comp Is Nothing Then ProcessComponent comp
        If blocked Then Exit For
    Next i
    doc.ClearSelection2 True
    doc.GraphicsRedraw2
    RestorePreference
    If blocked Then MsgBox LOG_PREFIX & ": STOP - dimension creation or cleanup has an uncertain outcome. Inspect this drawing manually before rerunning.", vbCritical
    Debug.Print LOG_PREFIX & " END created=" & created & "; skipped=" & skipped & "; failed=" & failed & "; blocked=" & blocked
    MsgBox LOG_PREFIX & ": created=" & created & "; skipped=" & skipped & "; failed=" & failed & vbCrLf & _
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

Private Sub ProcessComponent(ByVal comp As SldWorks.Component2)
    On Error GoTo FailedComponent
    coordinateRejectLogCount = 0
    Dim edges As Variant
    edges = targetView.GetVisibleEntities2(comp, swViewEntityType_Edge)
    If IsEmpty(edges) Then
        skipped = skipped + 1
        Debug.Print LOG_PREFIX & " " & comp.Name2 & ": no visible edges"
        Exit Sub
    End If
    If Not IsArray(edges) Then Err.Raise vbObjectError + 11, , "Visible edges unavailable"

    ' One candidate per visible body, preserving each component instance.
    Dim bodies As New Collection, bestEdges As New Collection
    Dim lengths As New Collection
    Dim i As Long, j As Long, index As Long
    Dim edge As SldWorks.Edge, body As SldWorks.Body2, oldBody As SldWorks.Body2
    Dim length As Double, x As Double, y As Double, reason As String
    Dim stats As Object, key As Variant
    Set stats = CreateObject("Scripting.Dictionary")
    For i = LBound(edges) To UBound(edges)
        If TypeOf edges(i) Is SldWorks.Edge Then
        Set edge = edges(i)
        If Candidate(edge, comp, length, x, y, reason) Then
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
                    bestEdges.Add edge
                    lengths.Add length
                ElseIf length > CDbl(lengths(index)) + LENGTH_TOL_M Then
                    bestEdges.Add edge, Before:=index
                    bestEdges.Remove index + 1
                    lengths.Add length, Before:=index
                    lengths.Remove index + 1
                End If
            Else
                reason = "body unavailable"
            End If
        End If
        Else
            reason = "not an Edge"
        End If
        If Len(reason) > 0 Then
            If Not stats.Exists(reason) Then stats.Add reason, 0
            stats(reason) = CLng(stats(reason)) + 1
        End If
        reason = ""
    Next i
    If bestEdges.Count = 0 Then
        skipped = skipped + 1
        Debug.Print LOG_PREFIX & " " & comp.Name2 & ": no usable visible straight edge; edge count=" & (UBound(edges) - LBound(edges) + 1)
    End If
    Debug.Print LOG_PREFIX & " " & comp.Name2 & ": candidates=" & bestEdges.Count & "; visible edges=" & (UBound(edges) - LBound(edges) + 1)
    For Each key In stats.Keys
        Debug.Print LOG_PREFIX & "  rejected " & CStr(key) & ": " & stats(key)
    Next key
    For i = 1 To bestEdges.Count
        Set edge = bestEdges(i)
        AddEdgeDimension edge, comp
        If blocked Then Exit For
    Next i
    Exit Sub
FailedComponent:
    failed = failed + 1
    Debug.Print LOG_PREFIX & " " & comp.Name2 & ": component failed: " & Err.Description
End Sub

Private Function Candidate(ByVal edge As SldWorks.Edge, ByVal comp As SldWorks.Component2, ByRef length As Double, _
        ByRef labelX As Double, ByRef labelY As Double, ByRef reason As String, Optional ByVal reportCoordinates As Boolean = False) As Boolean
    On Error GoTo Unusable
    reason = "edge unavailable"
    If edge Is Nothing Then Exit Function
    Dim curve As SldWorks.Curve
    reason = "curve unavailable"
    Set curve = edge.GetCurve
    If curve Is Nothing Then Exit Function
    Dim lineResult As Variant
    reason = "IsLine call"
    lineResult = curve.IsLine
    If CBool(lineResult) = False Then
        reason = "not straight (IsLine=" & CStr(lineResult) & ", type=" & TypeName(lineResult) & ")"
        Exit Function
    End If
    reason = "vertex unavailable"
    Dim va As SldWorks.Vertex, vb As SldWorks.Vertex
    Set va = edge.GetStartVertex: Set vb = edge.GetEndVertex
    If va Is Nothing Then Exit Function
    If vb Is Nothing Then Exit Function
    Dim a As Variant, b As Variant
    reason = "vertex coordinates"
    a = va.GetPoint: b = vb.GetPoint
    length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2 + (b(2) - a(2)) ^ 2)
    reason = "shorter than MIN_LENGTH_M"
    If length < MIN_LENGTH_M Then Exit Function

    ' Determine a plausible coordinate path without silently assuming edge context.
    ' GetOutline is only a guard, NOT proof that points belong to this edge.
    Dim directA As Variant, directB As Variant, componentA As Variant, componentB As Variant
    Dim directOK As Boolean, componentOK As Boolean, path As String
    reason = "view coordinate projection"
    directOK = ProjectEndpoints(a, b, comp, False, directA, directB)
    componentOK = ProjectEndpoints(a, b, comp, True, componentA, componentB)
    Dim directInside As Boolean, componentInside As Boolean
    If directOK Then directInside = PointsInsideOutline(directA, directB)
    If componentOK Then componentInside = PointsInsideOutline(componentA, componentB)
    If directInside And componentInside Then
        If Not SameSheetPoints(directA, directB, componentA, componentB) Then
            reason = "ambiguous component coordinate context"
            LogRejectedCoordinates comp, reason, directOK, directA, directB, componentOK, componentA, componentB
            Exit Function
        End If
        a = directA: b = directB: path = "direct (paths agree)"
    ElseIf directInside Then
        a = directA: b = directB: path = "direct"
    ElseIf componentInside Then
        a = componentA: b = componentB: path = "Component.Transform2 then view"
    Else
        reason = "neither coordinate path fits view outline"
        LogRejectedCoordinates comp, reason, directOK, directA, directB, componentOK, componentA, componentB
        Exit Function
    End If
    Dim dx As Double, dy As Double, projected As Double
    dx = b(0) - a(0): dy = b(1) - a(1)
    projected = Sqr(dx * dx + dy * dy)
    reason = "projection shorter than MIN_PROJECTED_M"
    If projected < MIN_PROJECTED_M Then Exit Function
    ' Stable normal independent of the edge's vertex ordering.
    If dx < 0 Or (Abs(dx) < 0.000000001 And dy < 0) Then
        dx = -dx: dy = -dy
    End If
    labelX = (a(0) + b(0)) / 2 - dy / projected * OFFSET_M
    labelY = (a(1) + b(1)) / 2 + dx / projected * OFFSET_M
    If reportCoordinates Then
        Debug.Print LOG_PREFIX & " " & comp.Name2 & " coordinate path: " & path
        Debug.Print LOG_PREFIX & " sheet mm A=" & SheetPoint(a(0), a(1)) & "; B=" & SheetPoint(b(0), b(1))
        Debug.Print LOG_PREFIX & " midpoint mm=" & SheetPoint((a(0) + b(0)) / 2, (a(1) + b(1)) / 2) & "; requested label mm=" & SheetPoint(labelX, labelY)
    End If
    Candidate = True
    reason = ""
    Exit Function
Unusable:
    reason = reason & ": VBA error " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Function ProjectEndpoints(ByVal a As Variant, ByVal b As Variant, ByVal comp As SldWorks.Component2, _
        ByVal applyComponent As Boolean, ByRef outA As Variant, ByRef outB As Variant) As Boolean
    On Error GoTo ProjectionFailed
    Dim mu As SldWorks.MathUtility, pa As SldWorks.MathPoint, pb As SldWorks.MathPoint
    Dim transform As SldWorks.MathTransform
    Set mu = swApp.GetMathUtility
    If mu Is Nothing Then Exit Function
    Set pa = mu.CreatePoint(a): Set pb = mu.CreatePoint(b)
    If pa Is Nothing Then Exit Function
    If pb Is Nothing Then Exit Function
    If applyComponent Then
        Set transform = comp.Transform2
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

Private Function SameSheetPoints(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant, ByVal d As Variant) As Boolean
    SameSheetPoints = (Abs(a(0) - c(0)) <= COORD_EQUAL_TOL_M And Abs(a(1) - c(1)) <= COORD_EQUAL_TOL_M And _
        Abs(b(0) - d(0)) <= COORD_EQUAL_TOL_M And Abs(b(1) - d(1)) <= COORD_EQUAL_TOL_M)
End Function

Private Function SheetPoint(ByVal x As Double, ByVal y As Double) As String
    SheetPoint = "(" & Format$(x * 1000#, "0.###") & ", " & Format$(y * 1000#, "0.###") & ")"
End Function

Private Sub LogRejectedCoordinates(ByVal comp As SldWorks.Component2, ByVal reason As String, _
        ByVal directOK As Boolean, ByVal a As Variant, ByVal b As Variant, _
        ByVal componentOK As Boolean, ByVal c As Variant, ByVal d As Variant)
    ' Limit detailed samples; summary counters still include every rejected edge.
    If coordinateRejectLogCount >= 5 Then Exit Sub
    coordinateRejectLogCount = coordinateRejectLogCount + 1
    Debug.Print LOG_PREFIX & " " & comp.Name2 & ": " & reason
    If directOK Then Debug.Print LOG_PREFIX & " direct sheet mm A=" & SheetPoint(a(0), a(1)) & "; B=" & SheetPoint(b(0), b(1))
    If componentOK Then Debug.Print LOG_PREFIX & " component sheet mm A=" & SheetPoint(c(0), c(1)) & "; B=" & SheetPoint(d(0), d(1))
End Sub

Private Sub AddEdgeDimension(ByVal edge As SldWorks.Edge, ByVal comp As SldWorks.Component2)
    On Error GoTo FailedDimension
    Dim length As Double, x As Double, y As Double, reason As String, componentName As String
    componentName = comp.Name2
    If Not Candidate(edge, comp, length, x, y, reason, True) Then Err.Raise vbObjectError + 12, , "Candidate became unavailable: " & reason
    doc.ClearSelection2 True
    If CBool(targetView.SelectEntity(edge, False)) = False Then Err.Raise vbObjectError + 13, , "Cannot select edge in view"
    Dim dd As SldWorks.DisplayDimension, creationAttempted As Boolean
    creationAttempted = True
    Set dd = doc.AddDimension2(x, y, 0#)
    If dd Is Nothing Then Err.Raise vbObjectError + 14, , "AddDimension2 returned Nothing"
    Dim dimValue As SldWorks.Dimension
    Set dimValue = dd.GetDimension2(0)
    If dimValue Is Nothing Then Err.Raise vbObjectError + 15, , "Dimension readback unavailable"
    Dim actual As Double
    actual = dimValue.SystemValue
    If Abs(actual - length) > LENGTH_TOL_M Then Err.Raise vbObjectError + 16, , "Dimension is not the true edge length"
    doc.ClearSelection2 True
    Debug.Print LOG_PREFIX & " " & componentName & ": created, length=" & Format$(actual * 1000#, "0.###") & " mm"
    created = created + 1
    Exit Sub
FailedDimension:
    reason = Err.Description
    failed = failed + 1
    Debug.Print LOG_PREFIX & " " & componentName & ": failed: " & reason
    If Not dd Is Nothing Then
        If Not RemoveNewDimension(dd) Then
            blocked = True
            Debug.Print LOG_PREFIX & ": STOP: failed to remove the new dimension. Inspect manually."
        End If
    ElseIf creationAttempted Then
        blocked = True
        Debug.Print LOG_PREFIX & ": STOP: AddDimension2 was attempted but no dimension handle is available. Inspect manually."
    End If
    ClearSelectionSafely
End Sub

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
