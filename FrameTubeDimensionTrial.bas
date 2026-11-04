Attribute VB_Name = "FrameTubeDimensionTrial"
Option Explicit

' ONE tube boundary dimension. The user selects an edge or silhouette of the tube
' body; the macro finds a visible full-length plane/cylinder boundary after
' temporarily showing tangent edges. Original view mode/settings are restored.
' Standard SldWorks/SwConst only. No rebuild, model edit, override or save.
Private Const PREFIX As String = "FrameTubeDimensionTrial v1.2"
Private Const LENGTH_TOL_M As Double = 0.000001
Private Const POSITION_TOL_M As Double = 0.00001
Private Const CUTLIST_MM_TO_M As Double = 0.001
Private Const SHEET_MARGIN_M As Double = 0.001
Private Const PROBE_STEP_M As Double = 0.001
Private Type BoundaryCandidate
    Visible As SldWorks.Edge
    Source As SldWorks.Edge
    Score As Double
End Type
Private candidates() As BoundaryCandidate
Private candidateCount As Long
Private app As SldWorks.SldWorks
Private drawing As SldWorks.ModelDoc2
Private view As SldWorks.View
Private frameModel As SldWorks.ModelDoc2
Private frameInstance As SldWorks.Component2
Private frameDrawingComponent As SldWorks.Component2
Private selectedBody As SldWorks.Body2
Private visited As Collection, occupied As Collection
Private output As Object
Private outputPath As String
Private matchedFolder As SldWorks.Feature
Private folderMatches As Long
Private expectedLength As Double
Private preferenceCaptured As Boolean, oldInputPreference As Boolean
Private tangentCaptured As Boolean, oldTangentMode As Long
Private sheetWidth As Double, sheetHeight As Double
Private sheetA As Variant, sheetB As Variant, sheetCentre As Variant
Private sourceCentre As Variant

Public Sub main()
    On Error GoTo Failed
    Dim dd As SldWorks.DisplayDimension, attempted As Boolean, selectionTouched As Boolean
    CloseLog
    Set drawing = Nothing: Set view = Nothing: Set frameModel = Nothing
    Set frameInstance = Nothing: Set frameDrawingComponent = Nothing
    Set selectedBody = Nothing: Set matchedFolder = Nothing
    Set visited = New Collection: Set occupied = New Collection
    Erase candidates
    candidateCount = 0: folderMatches = 0: expectedLength = 0#
    preferenceCaptured = False: tangentCaptured = False: outputPath = ""
    OpenLog
    LogLine "START " & CStr(Now)
    LogLine "ONE tube boundary trial; initial edge/silhouette only identifies the body. Rerun may add a duplicate."
    Set app = Application.SldWorks
    Set drawing = app.ActiveDoc
    If drawing Is Nothing Then Err.Raise vbObjectError + 1, , "Open the drawing first."
    If drawing.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."
    Dim sm As SldWorks.SelectionMgr, picked As Object, selectedIndex As Long
    Set sm = drawing.SelectionManager
    Set picked = ReadSelectedTubeGeometry(sm, view, selectedIndex)
    If ApiBoolean(view.ProjectedDimensions) Then Err.Raise vbObjectError + 30, , "Set view dimension type to True, not Projected. View dimension type was not changed."
    Dim seed As Object, converted As Object, identifyingEdge As SldWorks.Edge
    Dim silhouette As SldWorks.SilhouetteEdge, identifyingFace As SldWorks.Face2, silhouetteView As SldWorks.View
    Dim silhouetteSeed As Boolean
    If TypeOf picked Is SldWorks.Edge Then
        Set identifyingEdge = picked
        ResolveContext identifyingEdge
        Set seed = identifyingEdge
    ElseIf TypeOf picked Is SldWorks.SilhouetteEdge Then
        silhouetteSeed = True
        Set silhouette = picked
        Set silhouetteView = silhouette.GetView
        If silhouetteView Is Nothing Then Err.Raise vbObjectError + 90, , "Selected silhouette has no view."
        If StrComp(silhouetteView.GetName2, view.GetName2, vbBinaryCompare) <> 0 Then _
            Err.Raise vbObjectError + 91, , "Silhouette view does not match selection view."
        Set identifyingFace = silhouette.GetFace
        If identifyingFace Is Nothing Then Err.Raise vbObjectError + 92, , "Selected silhouette has no supporting face."
        LogLine "SILHOUETTE seed index=" & selectedIndex & "; supporting face identifies body only; silhouette will NOT be dimensioned"
        ResolveSilhouetteContext identifyingFace, sm, selectedIndex
        Set seed = identifyingFace
    Else
        Err.Raise vbObjectError + 93, , "Unsupported tube selection object."
    End If
    Set converted = frameModel.Extension.GetCorrespondingEntity2(seed)
    If converted Is Nothing Then
        If silhouetteSeed Then
            ' Only the seed may already be part-context. Do not use name or
            ' geometry guesses: its body must have exactly one native identity
            ' match in the active, configuration-checked target part.
            If FaceBodyIsInTargetPart(identifyingFace) Then
                Set converted = identifyingFace
                LogLine "SILHOUETTE reverse returned Nothing; supporting face body verified in target part"
            End If
        End If
    End If
    If converted Is Nothing Then Err.Raise vbObjectError + 31, , "Selected seed has no verified source-part correspondence."
    If silhouetteSeed Then
        If Not TypeOf converted Is SldWorks.Face2 Then Err.Raise vbObjectError + 32, , "Silhouette face did not map to a source Face2."
        Set identifyingFace = converted
        Set selectedBody = identifyingFace.GetBody
    Else
        If Not TypeOf converted Is SldWorks.Edge Then Err.Raise vbObjectError + 32, , "Selected edge did not map to a source Edge."
        Set identifyingEdge = converted
        Set selectedBody = identifyingEdge.GetBody
    End If
    If selectedBody Is Nothing Then Err.Raise vbObjectError + 33, , "Selected source body unavailable."
    LogLine "TARGET BODY=[" & selectedBody.Name & "]; native self identity=" & app.IsSame(selectedBody, selectedBody)
    If app.IsSame(selectedBody, selectedBody) <> swObjectSame Then Err.Raise vbObjectError + 34, , "Body identity self-check failed."
    FindCutListMembership
    If folderMatches <> 1 Then Err.Raise vbObjectError + 35, , "Expected exactly one cut-list membership; found=" & folderMatches
    ReadCutListLength matchedFolder
    LogLine "TARGET cutlist=[" & matchedFolder.Name & "]; LENGTH mm=" & expectedLength * 1000#
    Dim box As Variant
    box = selectedBody.GetBodyBox
    If HasItems(box) = False Then Err.Raise vbObjectError + 36, , "Body box unavailable."
    sourceCentre = Array((box(0) + box(3)) / 2#, (box(1) + box(4)) / 2#, (box(2) + box(5)) / 2#)
    sheetCentre = ProjectPoint(sourceCentre)
    CollectBoundaries "before tangent-mode change", False
    ' Capture flag is set BEFORE the first mutating setter, including when the
    ' original mode is already Visible. A failed setter still enters restoration.
    oldTangentMode = view.GetDisplayTangentEdges2
    If oldTangentMode < 0 Or oldTangentMode > 2 Then Err.Raise vbObjectError + 61, , "Original tangent display mode is outside the supported enum range."
    tangentCaptured = True
    LogLine "TANGENT original mode=" & oldTangentMode & "; requested visible mode=" & swTangentEdgesVisible
    view.SetDisplayTangentEdges2 swTangentEdgesVisible
    If view.GetDisplayTangentEdges2 <> swTangentEdgesVisible Then _
        Err.Raise vbObjectError + 37, , "Cannot enable visible tangent edges."
    drawing.GraphicsRedraw2
    CollectBoundaries "after tangent edges visible", True
    If candidateCount = 0 Then Err.Raise vbObjectError + 38, , "No visible full-LENGTH straight plane/cylinder boundary of the selected body."
    Dim drawingDoc As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Set drawingDoc = drawing
    Set sheet = drawingDoc.GetCurrentSheet
    sheet.GetSize sheetWidth, sheetHeight
    If sheetWidth <= 0 Or sheetHeight <= 0 Then Err.Raise vbObjectError + 39, , "Invalid sheet size."
    SeedOccupied
    Dim chosen As Long, x As Double, y As Double, height As Double, halfWidth As Double
    height = 0.004: halfWidth = EstimateHalfWidth(expectedLength, height)
    chosen = ChooseBoundary(halfWidth, height, x, y)
    If chosen = 0 Then Err.Raise vbObjectError + 40, , "No free sheet-bounded near-edge text slot for any eligible boundary."
    LogLine "CHOSEN boundary=" & chosen & "; radial score m2=" & candidates(chosen).Score & "; length mm=" & expectedLength * 1000#
    oldInputPreference = ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate))
    preferenceCaptured = True
    app.SetUserPreferenceToggle swInputDimValOnCreate, False
    If ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate)) Then Err.Raise vbObjectError + 41, , "Cannot disable dimension value input."
    selectionTouched = True
    drawing.ClearSelection2 True
    If ApiBoolean(view.SelectEntity(candidates(chosen).Visible, False)) = False Then Err.Raise vbObjectError + 42, , "Cannot select chosen visible boundary."
    VerifySingleSelection sm
    attempted = True
    Set dd = drawing.AddDimension2(x, y, 0#)
    If dd Is Nothing Then Err.Raise vbObjectError + 43, , "AddDimension2 returned Nothing; outcome uncertain."
    Dim value As SldWorks.Dimension, ann As SldWorks.Annotation, fmt As SldWorks.TextFormat, p As Variant
    Set value = dd.GetDimension2(0)
    If value Is Nothing Then Err.Raise vbObjectError + 44, , "Dimension readback unavailable."
    If Abs(value.SystemValue - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 45, , "Created dimension is not the true cut-list length."
    If ApiBoolean(dd.GetOverride) Then Err.Raise vbObjectError + 46, , "Unexpected displayed-value override."
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Err.Raise vbObjectError + 47, , "New annotation unavailable."
    Set fmt = ann.GetTextFormat(0)
    If Not fmt Is Nothing Then
        If fmt.CharHeight > height Then height = fmt.CharHeight
    End If
    halfWidth = EstimateHalfWidth(expectedLength, height)
    If FindPosition(halfWidth, height, x, y) = False Then Err.Raise vbObjectError + 48, , "No text slot for actual font height."
    p = MoveAnnotationToTarget(ann, x, y)
    CheckFinalDimension dd, value, ann, p, x, y, halfWidth, height, "before tangent restore"
    If RestoreTangentMode = False Then Err.Raise vbObjectError + 49, , "Cannot restore original tangent edge display mode."
    p = ann.GetPosition
    CheckFinalDimension dd, value, ann, p, x, y, halfWidth, height, "AFTER tangent restore"
    drawing.ClearSelection2 True
    If RestorePreference = False Then Err.Raise vbObjectError + 50, , "Cannot restore original dimension value input preference."
    LogLine "END created=1; target body=[" & selectedBody.Name & "]; cutlist=[" & matchedFolder.Name & "]; length mm=" & expectedLength * 1000#
    LogLine "Original tangent display mode restored. Radial outer-boundary ranking and text bounds are heuristics; inspect the drawing."
    CloseLog
    MsgBox PREFIX & ": ONE tube length dimension created, mm=" & CStr(expectedLength * 1000#) & vbCrLf & _
        "Original display/settings restored. No document saved." & vbCrLf & "Log: " & outputPath, vbInformation
    Exit Sub
Failed:
    Dim message As String, cleanupOK As Boolean, tangentOK As Boolean, preferenceOK As Boolean
    message = CStr(Err.Number) & ": " & Err.Description
    LogLine "STOP " & message
    If Not dd Is Nothing Then
        cleanupOK = RemoveNewDimension(dd)
        LogLine "CLEANUP current new dimension removed=" & cleanupOK
        If cleanupOK = False Then LogLine "UNCERTAIN OUTCOME: current new dimension could not be removed; inspect manually."
    ElseIf attempted Then
        LogLine "UNCERTAIN OUTCOME: creation attempted without a handle; inspect drawing manually."
    Else
        cleanupOK = True
        LogLine "No dimension creation attempted."
    End If
    tangentOK = RestoreTangentMode
    preferenceOK = RestorePreference
    If selectionTouched Then ClearSelectionSafely
    LogLine "END stopped; tangent mode restored=" & tangentOK & "; input preference restored=" & preferenceOK
    CloseLog
    MsgBox PREFIX & " stopped: " & message & vbCrLf & "Tangent mode restored=" & tangentOK & _
        "; input setting restored=" & preferenceOK & vbCrLf & "Inspect drawing/log before rerunning. Log: " & outputPath, vbExclamation
End Sub

Private Sub CollectBoundaries(ByVal phase As String, ByVal keep As Boolean)
    Dim edges As Variant, i As Long, boundary As SldWorks.Edge, source As SldWorks.Edge
    Dim a As Variant, b As Variant, score As Double, passed As Long, rejected As Long
    Dim axisX As Double, axisY As Double, axisZ As Double, length As Double, dot As Double
    If keep Then
        Erase candidates
        candidateCount = 0
    End If
    edges = view.GetVisibleEntities2(frameDrawingComponent, swViewEntityType_Edge)
    If HasItems(edges) = False Then
        LogLine "CANDIDATES " & phase & ": visible edges=0; passed=0"
        Exit Sub
    End If
    For i = LBound(edges) To UBound(edges)
        If TypeOf edges(i) Is SldWorks.Edge Then
            Set boundary = edges(i)
            Set source = Nothing
            If EligibleBoundary(boundary, source, a, b, score) Then
                length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2 + (b(2) - a(2)) ^ 2)
                If passed = 0 Then
                    axisX = (b(0) - a(0)) / length: axisY = (b(1) - a(1)) / length: axisZ = (b(2) - a(2)) / length
                Else
                    dot = ((b(0) - a(0)) * axisX + (b(1) - a(1)) * axisY + (b(2) - a(2)) * axisZ) / length
                    If Abs(dot) < 0.99999 Then Err.Raise vbObjectError + 51, , "Ambiguous tube length direction: eligible boundaries are not parallel."
                End If
                passed = passed + 1
                If keep Then
                    candidateCount = candidateCount + 1
                    If candidateCount = 1 Then
                        ReDim candidates(1 To 1)
                    Else
                        ReDim Preserve candidates(1 To candidateCount)
                    End If
                    Set candidates(candidateCount).Visible = boundary
                    Set candidates(candidateCount).Source = source
                    candidates(candidateCount).Score = score
                End If
            Else
                rejected = rejected + 1
            End If
        End If
    Next i
    LogLine "CANDIDATES " & phase & ": visible edges=" & (UBound(edges) - LBound(edges) + 1) & _
        "; passed full-LENGTH plane/cylinder boundaries=" & passed & "; other/unusable=" & rejected
End Sub

Private Function EligibleBoundary(ByVal visible As SldWorks.Edge, ByRef source As SldWorks.Edge, _
        ByRef a As Variant, ByRef b As Variant, ByRef score As Double) As Boolean
    On Error GoTo NotEligible
    ' The visible edge's curve classification is deliberately NOT used: only
    ' its real source counterpart supplies the authoritative geometry.
    Dim converted As Object, body As SldWorks.Body2, length As Double, projectedA As Variant, projectedB As Variant
    Set converted = frameModel.Extension.GetCorrespondingEntity2(visible)
    If converted Is Nothing Then Exit Function
    If Not TypeOf converted Is SldWorks.Edge Then Exit Function
    Set source = converted
    Set body = source.GetBody
    If body Is Nothing Then Exit Function
    If app.IsSame(body, selectedBody) <> swObjectSame Then Exit Function
    If StrComp(body.Name, selectedBody.Name, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 52, , "Same-identity body name inconsistency."
    If StraightGeometry(source, a, b, length) = False Then Exit Function
    If Abs(length - expectedLength) > LENGTH_TOL_M Then Exit Function
    If PlaneCylinderBoundary(source) = False Then Exit Function
    projectedA = ProjectPoint(a): projectedB = ProjectPoint(b)
    If Sqr((projectedB(0) - projectedA(0)) ^ 2 + (projectedB(1) - projectedA(1)) ^ 2) < 0.00001 Then Exit Function
    score = RadialScore(a, b, length)
    EligibleBoundary = True
    Exit Function
NotEligible:
    LogLine "BOUNDARY rejected by API/geometry: " & Err.Number & ": " & Err.Description
End Function

Private Function PlaneCylinderBoundary(ByVal edge As SldWorks.Edge) As Boolean
    Dim faces As Variant, first As SldWorks.Face2, second As SldWorks.Face2
    Dim surface1 As SldWorks.Surface, surface2 As SldWorks.Surface
    Dim plane1 As Boolean, plane2 As Boolean, cylinder1 As Boolean, cylinder2 As Boolean
    faces = edge.GetTwoAdjacentFaces2
    If HasItems(faces) = False Then Exit Function
    If UBound(faces) - LBound(faces) <> 1 Then Exit Function
    Set first = faces(LBound(faces)): Set second = faces(LBound(faces) + 1)
    If first Is Nothing Or second Is Nothing Then Exit Function
    Set surface1 = first.GetSurface: Set surface2 = second.GetSurface
    If surface1 Is Nothing Or surface2 Is Nothing Then Exit Function
    plane1 = ApiBoolean(surface1.IsPlane): plane2 = ApiBoolean(surface2.IsPlane)
    cylinder1 = ApiBoolean(surface1.IsCylinder): cylinder2 = ApiBoolean(surface2.IsCylinder)
    PlaneCylinderBoundary = ((plane1 And cylinder2) Or (plane2 And cylinder1))
End Function

Private Function RadialScore(ByVal a As Variant, ByVal b As Variant, ByVal length As Double) As Double
    ' Heuristic only: outer boundaries usually lie farther from body-box centre
    ' perpendicular to the longitudinal axis than inner hollow-tube boundaries.
    Dim dx As Double, dy As Double, dz As Double, rx As Double, ry As Double, rz As Double, along As Double
    dx = (b(0) - a(0)) / length: dy = (b(1) - a(1)) / length: dz = (b(2) - a(2)) / length
    rx = (a(0) + b(0)) / 2# - sourceCentre(0)
    ry = (a(1) + b(1)) / 2# - sourceCentre(1)
    rz = (a(2) + b(2)) / 2# - sourceCentre(2)
    along = rx * dx + ry * dy + rz * dz
    RadialScore = rx * rx + ry * ry + rz * rz - along * along
    If RadialScore < 0# Then RadialScore = 0#
End Function

Private Function ChooseBoundary(ByVal halfWidth As Double, ByVal halfHeight As Double, ByRef x As Double, ByRef y As Double) As Long
    Dim used() As Boolean, i As Long, trial As Long, best As Long, bestScore As Double
    Dim a As Variant, b As Variant, length As Double
    ReDim used(1 To candidateCount)
    For trial = 1 To candidateCount
        best = 0: bestScore = -1#
        For i = 1 To candidateCount
            If used(i) = False And candidates(i).Score > bestScore Then
                best = i: bestScore = candidates(i).Score
            End If
        Next i
        If best = 0 Then Exit Function
        used(best) = True
        If StraightGeometry(candidates(best).Source, a, b, length) Then
            If Abs(length - expectedLength) <= LENGTH_TOL_M Then
                sheetA = ProjectPoint(a): sheetB = ProjectPoint(b)
                If FindPosition(halfWidth, halfHeight, x, y) Then
                    LogLine "CHOSEN source points m=(" & a(0) & ", " & a(1) & ", " & a(2) & ") -> (" & b(0) & ", " & b(1) & ", " & b(2) & ")"
                    LogLine "CHOSEN sheet points m=(" & sheetA(0) & ", " & sheetA(1) & ") -> (" & sheetB(0) & ", " & sheetB(1) & ")"
                    ChooseBoundary = best
                    Exit Function
                End If
            End If
        End If
    Next trial
End Function

Private Sub CheckFinalDimension(ByVal dd As SldWorks.DisplayDimension, ByVal value As SldWorks.Dimension, _
        ByVal ann As SldWorks.Annotation, ByVal p As Variant, ByVal x As Double, ByVal y As Double, _
        ByVal halfWidth As Double, ByVal height As Double, ByVal phase As String)
    RequirePosition p
    Dim actual As Double, dangling As Boolean, overridden As Boolean
    actual = value.SystemValue
    dangling = ApiBoolean(ann.IsDangling)
    overridden = ApiBoolean(dd.GetOverride)
    LogLine "CHECK " & phase & ": physical m=" & actual & "; expected m=" & expectedLength & "; dangling=" & dangling & _
        "; override=" & overridden & "; text XYZ m=(" & p(0) & ", " & p(1) & ", " & p(2) & ")"
    If Abs(actual - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 53, , "Physical dimension no longer equals cut-list LENGTH."
    If dangling Then Err.Raise vbObjectError + 54, , "New dimension is dangling after display-mode change."
    If overridden Then Err.Raise vbObjectError + 55, , "Unexpected displayed-value override."
    If Abs(p(0) - x) > POSITION_TOL_M Or Abs(p(1) - y) > POSITION_TOL_M Then Err.Raise vbObjectError + 56, , "Annotation position did not retain its fixed target."
    If PositionFits(CDbl(p(0)), CDbl(p(1)), halfWidth, height) = False Then Err.Raise vbObjectError + 57, , "Final text estimate is outside placement limits."
End Sub

Private Function RestoreTangentMode() As Boolean
    On Error GoTo RestoreFailed
    If tangentCaptured Then
        view.SetDisplayTangentEdges2 oldTangentMode
        If view.GetDisplayTangentEdges2 <> oldTangentMode Then Err.Raise vbObjectError + 58, , "Tangent display restore readback mismatch."
        drawing.GraphicsRedraw2
        tangentCaptured = False
        LogLine "TANGENT original mode restored=" & oldTangentMode
    End If
    RestoreTangentMode = True
    Exit Function
RestoreFailed:
    LogLine "TANGENT RESTORE ERROR: " & Err.Number & ": " & Err.Description
End Function

Private Sub CheckFolderMembership(ByVal feature As SldWorks.Feature)
    Dim folder As SldWorks.BodyFolder, bodies As Variant, body As SldWorks.Body2, i As Long
    Set folder = feature.GetSpecificFeature2
    If folder Is Nothing Then Err.Raise vbObjectError + 59, , "Cut-list body folder unavailable."
    bodies = folder.GetBodies
    If HasItems(bodies) = False Then Exit Sub
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        If Not body Is Nothing Then
            If app.IsSame(selectedBody, body) = swObjectSame Then
                If StrComp(selectedBody.Name, body.Name, vbBinaryCompare) <> 0 Then _
                    Err.Raise vbObjectError + 60, , "Same-identity source body returned inconsistent names."
                folderMatches = folderMatches + 1
                Set matchedFolder = feature
                LogLine "CUTLIST BODY membership=[" & feature.Name & "]; body=[" & body.Name & "]; native identity confirmed"
            End If
        End If
    Next i
End Sub

Private Sub SeedOccupied()
    Dim dd As SldWorks.DisplayDimension, ann As SldWorks.Annotation, p As Variant
    Set dd = view.GetFirstDisplayDimension5
    Do While Not dd Is Nothing
        Set ann = dd.GetAnnotation
        If Not ann Is Nothing Then
            p = ann.GetPosition
            If HasItems(p) Then occupied.Add Array(CDbl(p(0)), CDbl(p(1)), 0.015, 0.006)
        End If
        Set dd = dd.GetNext5
    Loop
End Sub

Private Sub FindCutListMembership()
    Dim feature As SldWorks.Feature
    Set feature = frameModel.FirstFeature
    Do While Not feature Is Nothing
        VisitFeature feature
        Set feature = feature.GetNextFeature
    Loop
End Sub

Private Sub VisitFeature(ByVal feature As SldWorks.Feature)
    If Not FirstVisit(feature) Then Exit Sub
    If StrComp(feature.GetTypeName2, "CutListFolder", vbTextCompare) = 0 Then CheckFolderMembership feature
    Dim child As SldWorks.Feature
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        VisitFeature child
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Sub ReadCutListLength(ByVal feature As SldWorks.Feature)
    Dim manager As SldWorks.CustomPropertyManager, names As Variant, i As Long, matches As Long
    Dim raw As String, resolved As String, wasResolved As Boolean, linked As Boolean, status As Long
    Set manager = feature.CustomPropertyManager
    names = manager.GetNames
    If Not HasItems(names) Then Err.Raise vbObjectError + 58, , "Cut-list properties unavailable."
    For i = LBound(names) To UBound(names)
        raw = "": resolved = "": wasResolved = False: linked = False
        status = manager.Get6(CStr(names(i)), True, raw, resolved, wasResolved, linked)
        If StrComp(PropertyToken(raw), "LENGTH", vbTextCompare) = 0 Then
            LogLine "LENGTH property=[" & CStr(names(i)) & "]; raw=[" & raw & "]; resolved=[" & resolved & _
                "]; status=" & status & "; wasResolved=" & wasResolved & "; linked=" & linked
            If ApiBoolean(wasResolved) = False Then Err.Raise vbObjectError + 59, , "Cut-list LENGTH is not resolved in the cached properties."
            If ApiBoolean(linked) = False Then Err.Raise vbObjectError + 79, , "Cut-list LENGTH property is not linked."
            matches = matches + 1
            expectedLength = ParsePlainPositiveNumber(resolved) * CUTLIST_MM_TO_M
        End If
    Next i
    If matches <> 1 Then Err.Raise vbObjectError + 60, , "Expected one linked LENGTH property; found=" & matches
    LogLine "LENGTH input unit required by this tube trial=mm; scaled to metres; geometry cross-check is mandatory."
End Sub

Private Function MoveAnnotationToTarget(ByVal ann As SldWorks.Annotation, _
        ByVal targetX As Double, ByVal targetY As Double) As Variant
    ' At most four commands: baseline, independent positive X/Y probes, inverse-J
    ' target command. Probe errors may increase: only final target is accepted.
    Dim initial As Variant, baseline As Variant, probeX As Variant, probeY As Variant, final As Variant
    initial = ann.GetPosition
    RequirePosition initial
    Dim z As Double, h As Double, j11 As Double, j12 As Double, j21 As Double, j22 As Double
    Dim det As Double, conditionBound As Double
    Dim errorX As Double, errorY As Double, correctionX As Double, correctionY As Double
    Dim commandX As Double, commandY As Double
    z = CDbl(initial(2)): h = PROBE_STEP_M
    If Abs(CDbl(initial(0)) - targetX) <= POSITION_TOL_M And Abs(CDbl(initial(1)) - targetY) <= POSITION_TOL_M Then
        MoveAnnotationToTarget = initial
        Exit Function
    End If
    baseline = CommandAnnotation(ann, targetX, targetY, z, "baseline")
    If Abs(CDbl(baseline(0)) - targetX) <= POSITION_TOL_M And Abs(CDbl(baseline(1)) - targetY) <= POSITION_TOL_M Then
        MoveAnnotationToTarget = baseline
        Exit Function
    End If
    If h <= POSITION_TOL_M * 10# Then Err.Raise vbObjectError + 59, , "Calibration step is not positive/sufficiently large."
    probeX = CommandAnnotation(ann, targetX + h, targetY, z, "X probe")
    probeY = CommandAnnotation(ann, targetX, targetY + h, z, "Y probe")
    j11 = (CDbl(probeX(0)) - CDbl(baseline(0))) / h
    j21 = (CDbl(probeX(1)) - CDbl(baseline(1))) / h
    j12 = (CDbl(probeY(0)) - CDbl(baseline(0))) / h
    j22 = (CDbl(probeY(1)) - CDbl(baseline(1))) / h
    errorX = targetX - CDbl(baseline(0)): errorY = targetY - CDbl(baseline(1))
    Dim solution As Variant
    solution = SolvePositionCorrection(j11, j12, j21, j22, errorX, errorY)
    correctionX = solution(0): correctionY = solution(1)
    det = solution(2): conditionBound = solution(3)
    commandX = targetX + correctionX: commandY = targetY + correctionY
    If commandX < -0.05 Or commandY < -0.05 Or commandX > sheetWidth + 0.05 Or commandY > sheetHeight + 0.05 Then _
        Err.Raise vbObjectError + 64, , "Calibrated command is outside bounded sheet-local range."
    LogLine "TEXT J=(" & j11 & ", " & j12 & "; " & j21 & ", " & j22 & "); determinant=" & det & "; condition bound=" & conditionBound
    final = CommandAnnotation(ann, commandX, commandY, z, "calibrated final")
    If Abs(CDbl(final(0)) - targetX) > POSITION_TOL_M Or Abs(CDbl(final(1)) - targetY) > POSITION_TOL_M Then _
        Err.Raise vbObjectError + 65, , "Calibrated annotation position did not meet fixed 10 micrometre XY tolerance."
    MoveAnnotationToTarget = final
End Function

Private Function SolvePositionCorrection(ByVal j11 As Double, ByVal j12 As Double, ByVal j21 As Double, _
        ByVal j22 As Double, ByVal errorX As Double, ByVal errorY As Double) As Variant
    ' Pure arithmetic: inverse of the measured 2D response, with scale,
    ' angular conditioning and local-command guards. No COM or global sheet state.
    Dim det As Double, columnX As Double, columnY As Double, normSquared As Double, conditionBound As Double
    Dim correctionX As Double, correctionY As Double
    det = j11 * j22 - j12 * j21
    columnX = Sqr(j11 * j11 + j21 * j21): columnY = Sqr(j12 * j12 + j22 * j22)
    normSquared = j11 * j11 + j12 * j12 + j21 * j21 + j22 * j22
    If columnX < 0.00001 Or columnY < 0.00001 Or columnX > 100000# Or columnY > 100000# Then _
        Err.Raise vbObjectError + 60, , "Calibration response scale is invalid or constrained."
    If Abs(det) / (columnX * columnY) < 0.0001 Then Err.Raise vbObjectError + 61, , "Calibration directions are nearly dependent."
    conditionBound = normSquared / Abs(det)
    If conditionBound > 10000# Then Err.Raise vbObjectError + 62, , "Calibration matrix is poorly conditioned."
    correctionX = (j22 * errorX - j12 * errorY) / det
    correctionY = (-j21 * errorX + j11 * errorY) / det
    If Sqr(correctionX * correctionX + correctionY * correctionY) > 0.05 Then _
        Err.Raise vbObjectError + 63, , "Calibration correction exceeds 50 mm."
    SolvePositionCorrection = Array(correctionX, correctionY, det, conditionBound)
End Function

Private Function CommandAnnotation(ByVal ann As SldWorks.Annotation, ByVal x As Double, ByVal y As Double, _
        ByVal z As Double, ByVal label As String) As Variant
    If ApiBoolean(ann.SetPosition2(x, y, z)) = False Then Err.Raise vbObjectError + 66, , "Annotation command rejected: " & label
    drawing.GraphicsRedraw2
    Dim p As Variant
    p = ann.GetPosition
    RequirePosition p
    LogLine "TEXT " & label & "; command XYZ m=(" & x & ", " & y & ", " & z & "); actual XYZ m=(" & p(0) & ", " & p(1) & ", " & p(2) & ")"
    CommandAnnotation = p
End Function

Private Function FindPosition(ByVal halfWidth As Double, ByVal halfHeight As Double, ByRef x As Double, ByRef y As Double) As Boolean
    Dim dx As Double, dy As Double, projected As Double, nx As Double, ny As Double
    Dim mx As Double, my As Double, dot As Double, side As Long, level As Long, offset As Double
    Dim shift As Long, t As Double
    dx = sheetB(0) - sheetA(0): dy = sheetB(1) - sheetA(1)
    projected = Sqr(dx * dx + dy * dy)
    If projected < 0.00001 Then Exit Function
    nx = -dy / projected: ny = dx / projected
    mx = (sheetA(0) + sheetB(0)) / 2#: my = (sheetA(1) + sheetB(1)) / 2#
    dot = (mx - sheetCentre(0)) * nx + (my - sheetCentre(1)) * ny
    If dot < 0 Then nx = -nx: ny = -ny
    For level = 0 To 6
        offset = 0.006 + level * 0.002
        For shift = 0 To 4
            Select Case shift
                Case 0: t = 0.5
                Case 1: t = 0.35
                Case 2: t = 0.65
                Case 3: t = 0.2
                Case Else: t = 0.8
            End Select
            mx = sheetA(0) + t * dx: my = sheetA(1) + t * dy
            For side = 1 To -1 Step -2
                x = mx + side * nx * offset: y = my + side * ny * offset
                If PositionFits(x, y, halfWidth, halfHeight) Then
                    FindPosition = True
                    Exit Function
                End If
            Next side
        Next shift
    Next level
End Function

Private Function FaceBodyIsInTargetPart(ByVal face As SldWorks.Face2) As Boolean
    Dim body As SldWorks.Body2, candidate As SldWorks.Body2, part As SldWorks.PartDoc
    Dim bodies As Variant, i As Long, matches As Long
    If face Is Nothing Then Exit Function
    If frameModel Is Nothing Then Exit Function
    If frameModel.GetType <> swDocPART Then Exit Function
    Set body = face.GetBody
    If body Is Nothing Then Exit Function
    Set part = frameModel
    bodies = part.GetBodies2(swSolidBody, False)
    If HasItems(bodies) = False Then Exit Function
    For i = LBound(bodies) To UBound(bodies)
        Set candidate = bodies(i)
        If Not candidate Is Nothing Then
            If app.IsSame(body, candidate) = swObjectSame Then matches = matches + 1
        End If
    Next i
    LogLine "SILHOUETTE local-face body identity matches in target part=" & matches
    FaceBodyIsInTargetPart = (matches = 1)
End Function

Private Sub ResolveSilhouetteContext(ByVal face As SldWorks.Face2, _
        ByVal sm As SldWorks.SelectionMgr, ByVal selectedIndex As Long)
    If face Is Nothing Then Err.Raise vbObjectError + 92, , "Silhouette supporting face unavailable."
    Dim reference As SldWorks.ModelDoc2, rawComponent As Object
    Dim component As SldWorks.Component2, dc As SldWorks.DrawingComponent
    Set reference = view.ReferencedDocument
    If reference Is Nothing Then Err.Raise vbObjectError + 6, , "Resolve/open the referenced model first."
    LogLine "VIEW=[" & view.GetName2 & "]; reference=[" & reference.GetTitle & "]; config=[" & view.ReferencedConfiguration & "]"
    CheckConfiguration reference, view.ReferencedConfiguration
    If reference.GetType = swDocASSEMBLY Then
        ' The supporting face may be part-context and cannot identify which of
        ' several identical instances was picked. Use the actual selection index.
        Set rawComponent = sm.GetSelectedObjectsComponent4(selectedIndex, -1)
        If rawComponent Is Nothing Then Err.Raise vbObjectError + 94, , "Selected silhouette component unavailable; choose an ordinary tube end edge."
        Dim selectedWrapper As SldWorks.DrawingComponent, wrapperView As SldWorks.View
        LogLine "SILHOUETTE component result type=" & TypeName(rawComponent) & _
            "; supports Component2=" & (TypeOf rawComponent Is SldWorks.Component2) & _
            "; supports DrawingComponent=" & (TypeOf rawComponent Is SldWorks.DrawingComponent)
        If TypeOf rawComponent Is SldWorks.DrawingComponent Then
            Set selectedWrapper = rawComponent
            Set wrapperView = selectedWrapper.View
            If wrapperView Is Nothing Then Err.Raise vbObjectError + 97, , "Selected DrawingComponent has no view."
            If StrComp(wrapperView.GetName2, view.GetName2, vbBinaryCompare) <> 0 Then _
                Err.Raise vbObjectError + 98, , "Selected DrawingComponent belongs to a different view."
            Set component = selectedWrapper.Component
            If component Is Nothing Then Err.Raise vbObjectError + 99, , "Selected DrawingComponent has no model Component2."
            LogLine "SILHOUETTE unwrapped DrawingComponent in view=[" & wrapperView.GetName2 & "]"
        ElseIf TypeOf rawComponent Is SldWorks.Component2 Then
            Set component = rawComponent
        Else
            Err.Raise vbObjectError + 95, , "Silhouette owner is neither Component2 nor DrawingComponent."
        End If
        LogLine "SILHOUETTE selected component=[" & component.Name2 & "]; index=" & selectedIndex
        Set frameInstance = ResolveFrameInstance(reference, component)
        Set frameModel = frameInstance.GetModelDoc2
        If frameModel Is Nothing Then Err.Raise vbObjectError + 9, , "Resolve/open the selected frame part first."
        CheckConfiguration frameModel, frameInstance.ReferencedConfiguration
        If selectedWrapper Is Nothing Then
            Set dc = frameInstance.GetDrawingComponent(view)
        Else
            Set dc = selectedWrapper
        End If
        If dc Is Nothing Then Err.Raise vbObjectError + 96, , "Selected frame instance has no drawing component in this view."
        Set frameDrawingComponent = dc.Component
        If frameDrawingComponent Is Nothing Then Err.Raise vbObjectError + 8, , "Selected frame drawing component unavailable."
        Dim verifiedOwner As SldWorks.Component2
        Set verifiedOwner = ResolveFrameInstance(reference, frameDrawingComponent)
        If app.IsSame(verifiedOwner, frameInstance) <> swObjectSame Then _
            Err.Raise vbObjectError + 100, , "Drawing wrapper resolved to a different frame instance."
        LogLine "FRAME INSTANCE=[" & frameInstance.Name2 & "]; config=[" & frameInstance.ReferencedConfiguration & "]"
    ElseIf reference.GetType = swDocPART Then
        Set frameModel = reference
        Set dc = view.RootDrawingComponent
        If Not dc Is Nothing Then Set frameDrawingComponent = dc.Component
        If frameDrawingComponent Is Nothing Then Err.Raise vbObjectError + 67, , "Part-view drawing component unavailable."
    Else
        Err.Raise vbObjectError + 10, , "Only assembly and part references supported."
    End If
    If frameModel.GetType <> swDocPART Then Err.Raise vbObjectError + 11, , "Selected frame must be a PART."
    LogLine "FRAME MODEL=[" & frameModel.GetTitle & "]; path=[" & frameModel.GetPathName & "]"
End Sub

Private Sub ResolveContext(ByVal selectedEdge As SldWorks.Edge)
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
        Set frameDrawingComponent = drawingComp
        If drawingComp Is Nothing Then Err.Raise vbObjectError + 8, , "Selected drawing component unavailable."
        Set frameInstance = ResolveFrameInstance(reference, drawingComp)
        Set frameModel = frameInstance.GetModelDoc2
        If frameModel Is Nothing Then Err.Raise vbObjectError + 9, , "Resolve/open selected frame part first."
        CheckConfiguration frameModel, frameInstance.ReferencedConfiguration
        LogLine "FRAME INSTANCE=[" & frameInstance.Name2 & "]; config=[" & frameInstance.ReferencedConfiguration & "]"
    ElseIf reference.GetType = swDocPART Then
        Set frameModel = reference
        Set entity = selectedEdge
        Set dc = entity.GetDrawingComponent(view)
        If Not dc Is Nothing Then Set frameDrawingComponent = dc.Component
        If frameDrawingComponent Is Nothing Then
            Set dc = view.RootDrawingComponent
            If Not dc Is Nothing Then Set frameDrawingComponent = dc.Component
        End If
        If frameDrawingComponent Is Nothing Then Err.Raise vbObjectError + 67, , "Part-view drawing component unavailable."
    Else
        Err.Raise vbObjectError + 10, , "Only assembly and part references supported."
    End If
    If frameModel.GetType <> swDocPART Then Err.Raise vbObjectError + 11, , "Selected frame must be a PART."
    LogLine "FRAME MODEL=[" & frameModel.GetTitle & "]; path=[" & frameModel.GetPathName & "]"
End Sub

Private Sub RequirePosition(ByVal coordinates As Variant)
    If HasItems(coordinates) = False Then Err.Raise vbObjectError + 52, , "Annotation XYZ position unavailable."
    If LBound(coordinates) <> 0 Or UBound(coordinates) <> 2 Then _
        Err.Raise vbObjectError + 52, , "Expected annotation position array with three XYZ coordinates."
End Sub

Private Function ApiBoolean(ByVal raw As Variant) As Boolean
    ' A comparison produces a canonical VBA Boolean for either COM True
    ' representation (1 or -1), rather than applying bitwise Not to raw bits.
    ApiBoolean = (CBool(raw) <> False)
End Function

Private Function PropertyToken(ByVal raw As String) As String
    Dim s As String, separator As Long
    s = Trim$(raw)
    If Len(s) > 0 Then
        If Left$(s, 1) = Chr$(34) Then s = Mid$(s, 2)
    End If
    separator = InStr(1, s, "@@@", vbBinaryCompare)
    If separator > 1 Then PropertyToken = Left$(s, separator - 1)
End Function

Private Function ParsePlainPositiveNumber(ByVal text As String) As Double
    Dim s As String, i As Long, digit As String, separators As Long, digits As Long
    s = Trim$(text)
    For i = 1 To Len(s)
        digit = Mid$(s, i, 1)
        If digit >= "0" And digit <= "9" Then
            digits = digits + 1
        ElseIf digit = "." Or digit = "," Then
            separators = separators + 1
        Else
            Err.Raise vbObjectError + 61, , "LENGTH must be a plain positive decimal in millimetres: [" & text & "]"
        End If
    Next i
    If digits = 0 Or separators > 1 Then Err.Raise vbObjectError + 62, , "Ambiguous numeric LENGTH: [" & text & "]"
    ParsePlainPositiveNumber = Val(Replace(s, ",", "."))
    If ParsePlainPositiveNumber <= 0 Then Err.Raise vbObjectError + 63, , "LENGTH must be positive."
End Function

Private Function ProjectPoint(ByVal coordinates As Variant) As Variant
    Dim utility As SldWorks.MathUtility, point As SldWorks.MathPoint, transform As SldWorks.MathTransform
    Set utility = app.GetMathUtility
    Set point = utility.CreatePoint(coordinates)
    If point Is Nothing Then Err.Raise vbObjectError + 64, , "Cannot create projection point."
    If Not frameInstance Is Nothing Then
        Set transform = frameInstance.Transform2
        If transform Is Nothing Then Err.Raise vbObjectError + 65, , "Component transform unavailable."
        Set point = point.MultiplyTransform(transform)
    End If
    Set transform = view.ModelToViewTransform
    If transform Is Nothing Then Err.Raise vbObjectError + 66, , "View transform unavailable."
    Set point = point.MultiplyTransform(transform)
    ProjectPoint = point.ArrayData
End Function

Private Sub VerifySingleSelection(ByVal sm As SldWorks.SelectionMgr)
    If sm.GetSelectedObjectCount2(-1) <> 1 Then Err.Raise vbObjectError + 67, , "Dimension selection is not exactly one object."
    If sm.GetSelectedObjectType3(1, -1) <> swSelEDGES Then Err.Raise vbObjectError + 68, , "Dimension selection is not an edge."
    Dim selectedView As SldWorks.View
    Set selectedView = sm.GetSelectedObjectsDrawingView2(1, -1)
    If selectedView Is Nothing Then Err.Raise vbObjectError + 69, , "Dimension selection has no view."
    If StrComp(selectedView.GetName2, view.GetName2, vbBinaryCompare) <> 0 Then _
        Err.Raise vbObjectError + 70, , "Dimension selection belongs to a different view."
    Dim picked As Object, localObject As Object, localEdge As SldWorks.Edge, body As SldWorks.Body2
    Set picked = sm.GetSelectedObject6(1, -1)
    If picked Is Nothing Then Err.Raise vbObjectError + 73, , "Dimension selection object unavailable."
    If Not TypeOf picked Is SldWorks.Edge Then Err.Raise vbObjectError + 74, , "Dimension selection object is not an Edge."
    Dim ownerOther As Boolean, ownerKnown As Boolean
    ownerKnown = AttachmentOwner(picked, ownerOther)
    If ownerKnown = False Or ownerOther Then _
        Err.Raise vbObjectError + 79, , "Dimension selection frame INSTANCE could not be verified."
    Set localObject = frameModel.Extension.GetCorrespondingEntity2(picked)
    If localObject Is Nothing Then Err.Raise vbObjectError + 75, , "Dimension selection has no source correspondence."
    If Not TypeOf localObject Is SldWorks.Edge Then Err.Raise vbObjectError + 76, , "Dimension selection source is not an Edge."
    Set localEdge = localObject
    Set body = localEdge.GetBody
    If body Is Nothing Then Err.Raise vbObjectError + 77, , "Dimension selection source body unavailable."
    LogLine "BOUND SELECTION BODY=[" & body.Name & "]; IsSame target=" & app.IsSame(body, selectedBody)
    If app.IsSame(body, selectedBody) <> swObjectSame Or StrComp(body.Name, selectedBody.Name, vbBinaryCompare) <> 0 Then _
        Err.Raise vbObjectError + 78, , "Dimension selection is not attached to the target frame body."
End Sub

Private Function AttachmentOwner(ByVal item As Object, ByRef ownerOther As Boolean) As Boolean
    On Error GoTo Unconfirmed
    If frameInstance Is Nothing Then
        Dim reference As SldWorks.ModelDoc2
        Set reference = view.ReferencedDocument
        If Not reference Is Nothing Then AttachmentOwner = (app.IsSame(reference, frameModel) = swObjectSame)
        Exit Function
    End If
    Dim entity As SldWorks.Entity, dc As SldWorks.DrawingComponent, component As SldWorks.Component2
    Set entity = item
    Set dc = entity.GetDrawingComponent(view)
    If dc Is Nothing Then Exit Function
    Set component = dc.Component
    If component Is Nothing Then Exit Function
    Set reference = view.ReferencedDocument
    Set component = ResolveFrameInstance(reference, component)
    If app.IsSame(component, frameInstance) = swObjectSame Then
        If StrComp(component.Name2, frameInstance.Name2, vbBinaryCompare) <> 0 Then _
            Err.Raise vbObjectError + 58, , "Existing attachment component identity/name mismatch."
        AttachmentOwner = True
    Else
        ownerOther = True
    End If
    Exit Function
Unconfirmed:
    LogLine "EXISTING attachment owner unconfirmed: " & Err.Number & ": " & Err.Description
End Function

Private Function EstimateHalfWidth(ByVal length As Double, ByVal height As Double) As Double
    ' Approximate half width, not measured glyph/extension-line bounds.
    EstimateHalfWidth = (Len(Format$(length * 1000#, "0.###")) + 4) * height * 0.4 + 0.001
    If EstimateHalfWidth < 0.006 Then EstimateHalfWidth = 0.006
End Function

Private Function PositionFits(ByVal x As Double, ByVal y As Double, ByVal halfWidth As Double, ByVal halfHeight As Double) As Boolean
    If x - halfWidth < SHEET_MARGIN_M Or x + halfWidth > sheetWidth - SHEET_MARGIN_M Then Exit Function
    If y - halfHeight < SHEET_MARGIN_M Or y + halfHeight > sheetHeight - SHEET_MARGIN_M Then Exit Function
    Dim rectangle As Variant
    For Each rectangle In occupied
        If Abs(x - rectangle(0)) < halfWidth + rectangle(2) + 0.001 And _
            Abs(y - rectangle(1)) < halfHeight + rectangle(3) + 0.001 Then Exit Function
    Next rectangle
    PositionFits = True
End Function

Private Function RestorePreference() As Boolean
    On Error GoTo RestoreFailed
    If preferenceCaptured Then
        app.SetUserPreferenceToggle swInputDimValOnCreate, oldInputPreference
        If ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate)) <> oldInputPreference Then _
            Err.Raise vbObjectError + 71, , "Input dimension preference restore failed."
        preferenceCaptured = False
        LogLine "SETTING Input dimension value restored=" & oldInputPreference
    End If
    RestorePreference = True
    Exit Function
RestoreFailed:
    LogLine "SETTING RESTORE ERROR: " & Err.Number & ": " & Err.Description
End Function

Private Function RemoveNewDimension(ByVal dd As SldWorks.DisplayDimension) As Boolean
    On Error GoTo CannotRemove
    Dim ann As SldWorks.Annotation
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Exit Function
    drawing.ClearSelection2 True
    If ApiBoolean(ann.Select3(False, Nothing)) = False Then Exit Function
    If drawing.SelectionManager.GetSelectedObjectCount2(-1) <> 1 Then Exit Function
    RemoveNewDimension = ApiBoolean(drawing.Extension.DeleteSelection2(0))
    Exit Function
CannotRemove:
    LogLine "CLEANUP ERROR: " & Err.Number & ": " & Err.Description
End Function

Private Sub ClearSelectionSafely()
    On Error GoTo ClearFailed
    drawing.ClearSelection2 True
    Exit Sub
ClearFailed:
    LogLine "SELECTION CLEANUP ERROR: " & Err.Number & ": " & Err.Description
End Sub

Private Function ReadSelectedTubeGeometry(ByVal sm As SldWorks.SelectionMgr, _
        ByRef selectedView As SldWorks.View, ByRef selectedIndex As Long) As Object
    Dim count As Long, geometryCount As Long, i As Long, selectionType As Long
    Dim item As Object, candidate As Object, itemView As SldWorks.View
    count = sm.GetSelectedObjectCount2(-1)
    LogLine "SELECTION total=" & count
    If count = 0 Then Err.Raise vbObjectError + 3, , "Select one tube edge or silhouette in the drawing before running."
    For i = 1 To count
        selectionType = sm.GetSelectedObjectType3(i, -1)
        Set item = sm.GetSelectedObject6(i, -1)
        Set itemView = sm.GetSelectedObjectsDrawingView2(i, -1)
        If i <= 30 Then
            LogLine "SELECTION index=" & i & "; API type=" & selectionType & "; object=" & TypeName(item)
            If Not itemView Is Nothing Then LogLine "  drawing view=[" & itemView.GetName2 & "]"
        End If
        If selectionType = swSelEDGES Or selectionType = swSelSILHOUETTES Then
            geometryCount = geometryCount + 1
            If geometryCount = 1 Then
                selectedIndex = i
                Set selectedView = itemView
                If Not item Is Nothing Then
                    If selectionType = swSelEDGES And TypeOf item Is SldWorks.Edge Then Set candidate = item
                    If selectionType = swSelSILHOUETTES And TypeOf item Is SldWorks.SilhouetteEdge Then Set candidate = item
                End If
            End If
        End If
    Next i
    LogLine "SELECTION tube geometry count=" & geometryCount & "; selected index=" & selectedIndex & "; other selections=" & (count - geometryCount)
    If geometryCount <> 1 Then Err.Raise vbObjectError + 3, , "Expected ONE tube edge or silhouette; found=" & geometryCount
    If candidate Is Nothing Then Err.Raise vbObjectError + 4, , "Selected tube geometry object unavailable or unsupported."
    If selectedView Is Nothing Then Err.Raise vbObjectError + 5, , "Selected tube geometry has no drawing view."
    Set ReadSelectedTubeGeometry = candidate
End Function

Private Sub OpenLog()
    On Error GoTo NoFile
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    outputPath = fs.BuildPath(fs.GetSpecialFolder(2), "FrameTubeDimensionTrial-" & fs.GetTempName & ".txt")
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
            Err.Raise vbObjectError + 13, , "Configuration mismatch. Macro does not switch configurations."
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

Private Function StraightGeometry(ByVal edge As SldWorks.Edge, ByRef a As Variant, ByRef b As Variant, ByRef length As Double) As Boolean
    Dim curve As SldWorks.Curve, va As SldWorks.Vertex, vb As SldWorks.Vertex, line As Boolean
    Set curve = edge.GetCurve
    If curve Is Nothing Then Err.Raise vbObjectError + 43, , "Curve unavailable."
    line = ApiBoolean(curve.IsLine)
    If line = False Then Exit Function
    Set va = edge.GetStartVertex: Set vb = edge.GetEndVertex
    If va Is Nothing Or vb Is Nothing Then Err.Raise vbObjectError + 44, , "Straight edge vertices unavailable."
    a = va.GetPoint: b = vb.GetPoint
    length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2 + (b(2) - a(2)) ^ 2)
    If length <= 0.00000001 Then Exit Function
    StraightGeometry = True
End Function
