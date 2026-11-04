Attribute VB_Name = "FrameCutListDimensionsV2"
Option Explicit

' Bulk true lengths for cut-list bodies of ONE explicitly selected frame part.
' No other assembly components are dimensioned. No model saves/rebuilds,
' configuration switches or text/value overrides. Standard SldWorks/SwConst only.
Private Const PREFIX As String = "FrameCutListDimensionsV2 v2.2"
Private Const LENGTH_TOL_M As Double = 0.000001
Private Const POSITION_TOL_M As Double = 0.00001
Private Const CUTLIST_MM_TO_M As Double = 0.001
Private Const SHEET_MARGIN_M As Double = 0.001
Private Const MAX_EDGE_ATTEMPTS As Long = 8
Private Const PROBE_STEP_M As Double = 0.001
Private Const MAX_SOURCE_FACES As Long = 2000
Private Const MAX_SOURCE_CANDIDATES As Long = 32
Private Const SOURCE_PREP_BUDGET_S As Double = 30#
Private Type BodyJob
    Body As SldWorks.Body2
    FolderName As String
    Expected As Double
    State As String
    Reason As String
    Visible As Collection
    Source As Collection
    Route As String
    NormalCount As Long
    Scores As Collection
    CreatedDD As SldWorks.DisplayDimension
    TargetX As Double
    TargetY As Double
    HalfWidth As Double
    HalfHeight As Double
    AxisX As Double
    AxisY As Double
    AxisZ As Double
End Type
Private jobs() As BodyJob
Private jobCount As Long
Private app As SldWorks.SldWorks
Private drawing As SldWorks.ModelDoc2
Private view As SldWorks.View
Private frameModel As SldWorks.ModelDoc2
Private frameInstance As SldWorks.Component2
Private frameDrawingComponent As SldWorks.Component2
Private selectedBody As SldWorks.Body2
Private expectedLength As Double
Private visited As Collection, occupied As Collection
Private logAvailable As Boolean
Private outputPath As String
Private preferenceCaptured As Boolean, oldInputPreference As Boolean
Private sheetWidth As Double, sheetHeight As Double
Private sheetA As Variant, sheetB As Variant, sheetCentre As Variant
Private blocked As Boolean, queueComplete As Boolean
Private ignoredExisting As Long, unresolvedExisting As Long, unknownExisting As Long
Private resultSummary As String, resultWarning As Boolean
Private seedFace As SldWorks.Face2, supportFace As SldWorks.Face2
Private sourcePrepStarted As Double, sourcePrepTimedOut As Boolean
Private sourceCentre As Variant
Private currentJob As Long

Public Sub main()
    On Error GoTo Failed
    CloseLog
    Set drawing = Nothing: Set view = Nothing: Set frameModel = Nothing
    Set frameInstance = Nothing: Set frameDrawingComponent = Nothing: Set selectedBody = Nothing
    Set seedFace = Nothing: Set supportFace = Nothing
    Set visited = New Collection: Set occupied = New Collection
    sourcePrepStarted = 0#: sourcePrepTimedOut = False
    Erase jobs
    jobCount = 0: currentJob = 0: preferenceCaptured = False
    blocked = False: queueComplete = False
    ignoredExisting = 0: unresolvedExisting = 0: unknownExisting = 0
    resultSummary = "": resultWarning = False: outputPath = ""
    OpenLog
    LogLine "START " & CStr(Now)
    LogLine "Two phases; original bulk placement unchanged. Existing dimensions are preserved."
    Set app = Application.SldWorks
    Set drawing = app.ActiveDoc
    If drawing Is Nothing Then Err.Raise vbObjectError + 1, , "Open the drawing first."
    If drawing.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."
    Dim sm As SldWorks.SelectionMgr, picked As Object, selectedIndex As Long
    Set sm = drawing.SelectionManager
    Set picked = ReadSelectedTubeGeometry(sm, view, selectedIndex)
    If ApiBoolean(view.ProjectedDimensions) Then Err.Raise vbObjectError + 30, , _
        "Set selected view dimension type to True, not Projected. View dimension type was not changed."
    If TypeOf picked Is SldWorks.Edge Then
        Dim initialEdge As SldWorks.Edge
        Set initialEdge = picked
        ResolveContext initialEdge
    ElseIf TypeOf picked Is SldWorks.SilhouetteEdge Then
        Dim silhouette As SldWorks.SilhouetteEdge, silhouetteView As SldWorks.View
        Set silhouette = picked
        Set silhouetteView = silhouette.GetView
        If silhouetteView Is Nothing Then Err.Raise vbObjectError + 90, , "Selected silhouette has no view."
        If StrComp(silhouetteView.GetName2, view.GetName2, vbBinaryCompare) <> 0 Then _
            Err.Raise vbObjectError + 91, , "Silhouette view does not match selection view."
        Set seedFace = silhouette.GetFace
        If seedFace Is Nothing Then Err.Raise vbObjectError + 92, , "Selected silhouette has no supporting face."
        ResolveSilhouetteContext seedFace, sm, selectedIndex
    Else
        Err.Raise vbObjectError + 93, , "Select one frame edge or silhouette."
    End If
    BuildQueue
    queueComplete = True
    LogLine "QUEUE complete; physical cut-list bodies=" & jobCount
    If jobCount = 0 Then Err.Raise vbObjectError + 31, , "Selected part has no accessible cut-list bodies."
    Dim drawingDoc As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Set drawingDoc = drawing: Set sheet = drawingDoc.GetCurrentSheet
    sheet.GetSize sheetWidth, sheetHeight
    If sheetWidth <= 0 Or sheetHeight <= 0 Then Err.Raise vbObjectError + 32, , "Invalid sheet size."
    BuildVisibleCandidates
    Dim i As Long
    For i = 1 To jobCount
        jobs(i).NormalCount = jobs(i).Visible.Count
        If jobs(i).NormalCount > 0 Then jobs(i).Route = "normal visible edge"
    Next i
    InspectExistingDimensions
    If blocked = False And NeedSupportPhase Then DiscoverSupportFaces
    LogLine "PHASE 1: normal visible-edge candidates, original view style"
    For i = 1 To jobCount
        If blocked Then Exit For
        If jobs(i).State = "pending" And jobs(i).NormalCount > 0 Then
            EnsureInputPreference
            ProcessBody i
        End If
    Next i
    If blocked = False And NeedSupportPhase Then
        LogLine "PHASE 2: precollected source-face boundaries; view style unchanged"
        For i = 1 To jobCount
            If blocked Then Exit For
            If jobs(i).State = "pending" And jobs(i).NormalCount = 0 Then
                jobs(i).Route = "source-face boundary"
                If jobs(i).Visible.Count > 0 Then
                    EnsureInputPreference
                    ProcessBody i
                Else
                    jobs(i).State = "skipped"
                    jobs(i).Reason = "No usable full-LENGTH plane/cylinder boundary in this source cut-list body; visibility is not proven."
                End If
            End If
        Next i
    End If
    FinishAndValidate
    ClearSelectionSafely
    FinalReport
    CloseLog
    Dim status As String, icon As VbMsgBoxStyle
    status = "completed": icon = vbInformation
    If resultWarning Then status = "completed with unresolved bodies": icon = vbExclamation
    If blocked Then status = "STOPPED - inspect uncertain outcome/settings": icon = vbExclamation
    MsgBox PREFIX & ": " & status & vbCrLf & resultSummary & vbCrLf & _
        "No document was saved. Inspect dimensions before saving." & vbCrLf & "Log: " & outputPath, icon
    Exit Sub
Failed:
    Dim message As String
    message = CStr(Err.Number) & ": " & Err.Description
    blocked = True
    LogLine "STOP " & message
    FinishAndValidate
    ClearSelectionSafely
    FinalReport
    CloseLog
    MsgBox PREFIX & " stopped: " & message & vbCrLf & resultSummary & vbCrLf & _
        "Earlier valid new dimensions are retained. Inspect drawing and log before rerunning. Log: " & outputPath, vbExclamation
End Sub

Private Sub BuildQueue()
    Dim feature As SldWorks.Feature
    Set feature = frameModel.FirstFeature
    Do While Not feature Is Nothing
        VisitFeature feature
        Set feature = feature.GetNextFeature
    Loop
End Sub

Private Sub VisitFeature(ByVal feature As SldWorks.Feature)
    If FirstVisit(feature) = False Then Exit Sub
    If StrComp(feature.GetTypeName2, "CutListFolder", vbTextCompare) = 0 Then QueueFolder feature
    Dim child As SldWorks.Feature
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        VisitFeature child
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Sub QueueFolder(ByVal feature As SldWorks.Feature)
    Dim folder As SldWorks.BodyFolder, bodies As Variant, body As SldWorks.Body2
    Dim length As Double, reason As String, lengthOK As Boolean, i As Long, index As Long
    Set folder = feature.GetSpecificFeature2
    If folder Is Nothing Then Err.Raise vbObjectError + 34, , "Cut-list folder unavailable: " & feature.Name
    bodies = folder.GetBodies
    If HasItems(bodies) = False Then Exit Sub
    lengthOK = TryFolderLength(feature, length, reason)
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        If Not body Is Nothing Then
            index = FindJobForBody(body)
            If index > 0 Then
                jobs(index).State = "failed"
                jobs(index).Reason = "Duplicate cut-list membership: " & jobs(index).FolderName & "; " & feature.Name
            Else
                jobCount = jobCount + 1
                If jobCount = 1 Then
                    ReDim jobs(1 To 1)
                Else
                    ReDim Preserve jobs(1 To jobCount)
                End If
                Set jobs(jobCount).Body = body
                jobs(jobCount).FolderName = feature.Name
                jobs(jobCount).Expected = length
                jobs(jobCount).State = "pending"
                Set jobs(jobCount).Visible = New Collection
                Set jobs(jobCount).Source = New Collection
                Set jobs(jobCount).Scores = New Collection
                If lengthOK = False Then
                    jobs(jobCount).State = "skipped": jobs(jobCount).Reason = reason
                Else
                    ValidateLengthDirection jobCount
                End If
            End If
        End If
    Next i
End Sub

Private Function TryFolderLength(ByVal feature As SldWorks.Feature, ByRef length As Double, ByRef reason As String) As Boolean
    On Error GoTo InvalidLength
    Dim manager As SldWorks.CustomPropertyManager, names As Variant, i As Long, matches As Long
    Dim raw As String, resolved As String, wasResolved As Boolean, linked As Boolean, status As Long
    Set manager = feature.CustomPropertyManager
    names = manager.GetNames
    If HasItems(names) = False Then Err.Raise vbObjectError + 35, , "Cut-list properties unavailable."
    For i = LBound(names) To UBound(names)
        raw = "": resolved = "": wasResolved = False: linked = False
        status = manager.Get6(CStr(names(i)), True, raw, resolved, wasResolved, linked)
        If StrComp(PropertyToken(raw), "LENGTH", vbTextCompare) = 0 Then
            LogLine "LENGTH cutlist=[" & feature.Name & "]; property=[" & CStr(names(i)) & "]; resolved=[" & resolved & _
                "]; status=" & status & "; wasResolved=" & wasResolved & "; linked=" & linked
            If ApiBoolean(wasResolved) = False Then Err.Raise vbObjectError + 36, , "Cached LENGTH not resolved."
            If ApiBoolean(linked) = False Then Err.Raise vbObjectError + 37, , "LENGTH is not linked."
            matches = matches + 1
            length = ParsePlainPositiveNumber(resolved) * CUTLIST_MM_TO_M
        End If
    Next i
    If matches <> 1 Then Err.Raise vbObjectError + 38, , "Expected one linked LENGTH property; found=" & matches
    TryFolderLength = True
    Exit Function
InvalidLength:
    reason = "Unsupported cut-list LENGTH: " & Err.Description
    LogLine "LENGTH ERROR cutlist=[" & feature.Name & "]: " & reason
End Function

Private Sub ValidateLengthDirection(ByVal index As Long)
    On Error GoTo Unsupported
    Dim edges As Variant, edge As SldWorks.Edge, a As Variant, b As Variant, length As Double
    Dim i As Long, matches As Long, dx As Double, dy As Double, dz As Double, dot As Double
    edges = jobs(index).Body.GetEdges
    If HasItems(edges) = False Then Err.Raise vbObjectError + 39, , "Body has no accessible edges."
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        If StraightGeometry(edge, a, b, length) Then
            If Abs(length - jobs(index).Expected) <= LENGTH_TOL_M Then
                dx = (b(0) - a(0)) / length: dy = (b(1) - a(1)) / length: dz = (b(2) - a(2)) / length
                If matches = 0 Then
                    jobs(index).AxisX = dx: jobs(index).AxisY = dy: jobs(index).AxisZ = dz
                Else
                    dot = dx * jobs(index).AxisX + dy * jobs(index).AxisY + dz * jobs(index).AxisZ
                    If Abs(dot) < 0.999999 Then Err.Raise vbObjectError + 40, , "Ambiguous length direction: full-LENGTH edges are not parallel."
                End If
                matches = matches + 1
            End If
        End If
    Next i
    If matches = 0 Then Err.Raise vbObjectError + 41, , "Cut-list length is not represented by a straight source edge."
    LogLine "BODY preflight=[" & jobs(index).Body.Name & "]; length mm=" & jobs(index).Expected * 1000# & "; parallel full-length edges=" & matches
    Exit Sub
Unsupported:
    jobs(index).State = "skipped"
    jobs(index).Reason = "Unsupported source geometry: " & Err.Description
End Sub

Private Function FindJobForBody(ByVal body As SldWorks.Body2) As Long
    ' Body names need not be unique: patterned bodies can retain equal names.
    ' Only native identity selects a job. Name consistency is checked AFTER
    ' identity equality, never used to merge or reject two different bodies.
    Dim i As Long, same As Long
    For i = 1 To jobCount
        same = app.IsSame(body, jobs(i).Body)
        If same = swObjectSame Then
            If StrComp(body.Name, jobs(i).Body.Name, vbBinaryCompare) <> 0 Then _
                Err.Raise vbObjectError + 42, , "Identical source body returned inconsistent names: " & body.Name
            FindJobForBody = i
            Exit Function
        End If
    Next i
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

Private Sub BuildVisibleCandidates()
    Dim edges As Variant, item As Object, i As Long, found As Long, rejected As Long
    LogLine "BEGIN normal candidates GetVisibleEntities2 Edge1; target frame component only"
    edges = view.GetVisibleEntities2(frameDrawingComponent, swViewEntityType_Edge)
    LogLine "DONE normal candidates GetVisibleEntities2 Edge1"
    If HasItems(edges) = False Then
        LogLine "VISIBLE target frame component has no accessible visible edges."
        Exit Sub
    End If
    For i = LBound(edges) To UBound(edges)
        Set item = edges(i)
        If TypeOf item Is SldWorks.Edge Then
            If AddVisibleCandidate(item) Then found = found + 1 Else rejected = rejected + 1
        End If
    Next i
    LogLine "VISIBLE full-length candidate edges=" & found & "; other/unusable edges=" & rejected
End Sub

Private Function AddVisibleCandidate(ByVal visible As SldWorks.Edge) As Boolean
    On Error GoTo NotUsable
    Dim curve As SldWorks.Curve, converted As Object, source As SldWorks.Edge, body As SldWorks.Body2
    Dim a As Variant, b As Variant, length As Double, index As Long
    Set curve = visible.GetCurve
    If curve Is Nothing Then Exit Function
    If ApiBoolean(curve.IsLine) = False Then Exit Function
    Set converted = frameModel.Extension.GetCorrespondingEntity2(visible)
    If converted Is Nothing Then Exit Function
    If Not TypeOf converted Is SldWorks.Edge Then Exit Function
    Set source = converted
    If StraightGeometry(source, a, b, length) = False Then Exit Function
    Set body = source.GetBody
    If body Is Nothing Then Exit Function
    index = FindJobForBody(body)
    If index = 0 Then Exit Function
    If jobs(index).State <> "pending" Then Exit Function
    If Abs(length - jobs(index).Expected) > LENGTH_TOL_M Then Exit Function
    jobs(index).Visible.Add visible
    jobs(index).Source.Add source
    AddVisibleCandidate = True
    Exit Function
NotUsable:
    LogLine "VISIBLE candidate rejected: " & Err.Number & ": " & Err.Description
End Function

Private Sub ProcessBody(ByVal index As Long)
    currentJob = index
    If jobs(index).Visible.Count = 0 Then
        jobs(index).State = "skipped"
        jobs(index).Reason = "No verified full-LENGTH candidate for this source cut-list body."
        Exit Sub
    End If
    Set selectedBody = jobs(index).Body
    expectedLength = jobs(index).Expected
    Dim candidate As Long, reason As String
    For candidate = 1 To jobs(index).Visible.Count
        If candidate > MAX_EDGE_ATTEMPTS Or blocked Then Exit For
        LogLine "BODY attempt: [" & selectedBody.Name & "]; edge=" & candidate & "; length mm=" & expectedLength * 1000#
        If TryCreateDimension(jobs(index).Visible(candidate), jobs(index).Source(candidate), reason) Then
            jobs(index).State = "created"
            jobs(index).Reason = "True length and final annotation position verified."
            Exit Sub
        End If
    Next candidate
    jobs(index).State = "failed"
    jobs(index).Reason = reason
End Sub

Private Function TryCreateDimension(ByVal visible As SldWorks.Edge, ByVal source As SldWorks.Edge, ByRef reason As String) As Boolean
    On Error GoTo Failed
    Dim dd As SldWorks.DisplayDimension, attempted As Boolean
    Dim a As Variant, b As Variant, length As Double, box As Variant, centre As Variant
    If StraightGeometry(source, a, b, length) = False Then Err.Raise vbObjectError + 45, , "Source candidate no longer straight."
    If Abs(length - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 46, , "Source candidate length changed."
    sheetA = ProjectPoint(a): sheetB = ProjectPoint(b)
    box = selectedBody.GetBodyBox
    If HasItems(box) = False Then Err.Raise vbObjectError + 47, , "Body box unavailable."
    centre = Array((box(0) + box(3)) / 2#, (box(1) + box(4)) / 2#, (box(2) + box(5)) / 2#)
    sheetCentre = ProjectPoint(centre)
    Dim height As Double, halfWidth As Double, x As Double, y As Double
    height = 0.004: halfWidth = EstimateHalfWidth(expectedLength, height)
    If FindPosition(halfWidth, height, x, y) = False Then Err.Raise vbObjectError + 48, , "No free local text slot for candidate."
    drawing.ClearSelection2 True
    Dim selectionOK As Boolean
    LogLine "BEGIN view.SelectEntity; body=[" & selectedBody.Name & "]; route=" & jobs(currentJob).Route
    selectionOK = ApiBoolean(view.SelectEntity(visible, False))
    LogLine "DONE view.SelectEntity; body=[" & selectedBody.Name & "]; returned=" & selectionOK
    If selectionOK = False Then Err.Raise vbObjectError + 49, , "Cannot select candidate real edge in target view."
    VerifySingleSelection drawing.SelectionManager
    attempted = True
    LogLine "BEGIN AddDimension2; body=[" & selectedBody.Name & "]; expected mm=" & expectedLength * 1000#
    Set dd = drawing.AddDimension2(x, y, 0#)
    LogLine "DONE AddDimension2; body=[" & selectedBody.Name & "]; returned handle=" & (Not dd Is Nothing)
    If dd Is Nothing Then Err.Raise vbObjectError + 50, , "AddDimension2 returned Nothing; creation outcome uncertain."
    Dim value As SldWorks.Dimension, ann As SldWorks.Annotation, fmt As SldWorks.TextFormat, p As Variant, actual As Double
    Set value = dd.GetDimension2(0)
    If value Is Nothing Then Err.Raise vbObjectError + 51, , "Dimension readback unavailable."
    actual = value.SystemValue
    If Abs(actual - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 52, , "Created value is not true cut-list length."
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Err.Raise vbObjectError + 53, , "Annotation unavailable."
    Set fmt = ann.GetTextFormat(0)
    If Not fmt Is Nothing Then
        If fmt.CharHeight > height Then height = fmt.CharHeight
    End If
    halfWidth = EstimateHalfWidth(expectedLength, height)
    If FindPosition(halfWidth, height, x, y) = False Then Err.Raise vbObjectError + 54, , "No free slot for actual font."
    p = MoveAnnotationToTarget(ann, x, y)
    RequirePosition p
    If Abs(p(0) - x) > POSITION_TOL_M Or Abs(p(1) - y) > POSITION_TOL_M Then Err.Raise vbObjectError + 55, , "Final position did not converge."
    If PositionFits(CDbl(p(0)), CDbl(p(1)), halfWidth, height) = False Then Err.Raise vbObjectError + 56, , "Final position is outside estimated placement limits."
    actual = value.SystemValue
    If Abs(actual - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 57, , "Dimension value changed after positioning."
    If ApiBoolean(dd.GetOverride) Then Err.Raise vbObjectError + 96, , "Unexpected dimension override."
    If ApiBoolean(ann.IsDangling) Then Err.Raise vbObjectError + 97, , "New dimension is dangling."
    Set jobs(currentJob).CreatedDD = dd
    jobs(currentJob).TargetX = x: jobs(currentJob).TargetY = y
    jobs(currentJob).HalfWidth = halfWidth: jobs(currentJob).HalfHeight = height
    drawing.ClearSelection2 True
    LogLine "CREATED body=[" & selectedBody.Name & "]; value mm=" & actual * 1000# & "; text m=(" & p(0) & ", " & p(1) & ")"
    occupied.Add Array(CDbl(p(0)), CDbl(p(1)), halfWidth, height)
    TryCreateDimension = True
    Exit Function
Failed:
    reason = CStr(Err.Number) & ": " & Err.Description
    Set jobs(currentJob).CreatedDD = Nothing
    LogLine "ATTEMPT rejected: " & reason
    If Not dd Is Nothing Then
        If RemoveNewDimension(dd) = False Then
            blocked = True
            reason = reason & "; CLEANUP FAILED; inspect current new dimension manually."
        End If
    ElseIf attempted Then
        blocked = True
        reason = reason & "; CREATION UNCERTAIN; inspect drawing manually."
    End If
    ClearSelectionSafely
End Function

Private Sub FinalReport()
    Dim i As Long, created As Long, already As Long, skipped As Long, failed As Long, unprocessed As Long
    For i = 1 To jobCount
        If jobs(i).State = "pending" Then
            jobs(i).State = "unprocessed": jobs(i).Reason = "Stopped before processing this body."
        End If
        Select Case jobs(i).State
            Case "created": created = created + 1
            Case "already": already = already + 1
            Case "skipped": skipped = skipped + 1
            Case "failed": failed = failed + 1
            Case Else: unprocessed = unprocessed + 1
        End Select
        LogLine "RESULT body=[" & jobs(i).Body.Name & "]; cutlist=[" & jobs(i).FolderName & "]; expected mm=" & jobs(i).Expected * 1000# & _
            "; state=" & jobs(i).State & "; route=" & jobs(i).Route & "; reason=" & jobs(i).Reason
    Next i
    resultSummary = "total=" & jobCount & "; created=" & created & "; already=" & already & "; skipped=" & skipped & _
        "; failed=" & failed & "; unprocessed=" & unprocessed & "; blocked=" & blocked
    resultWarning = (skipped > 0 Or failed > 0 Or unprocessed > 0 Or queueComplete = False)
    LogLine "END " & resultSummary & "; sum=" & (created + already + skipped + failed + unprocessed) & "; queue complete=" & queueComplete
    LogLine "EXISTING other/ignored=" & ignoredExisting & "; unresolved=" & unresolvedExisting & "; unknown-body barrier count=" & unknownExisting
    LogLine "Placement uses estimated text rectangles; inspect geometry, extension lines and neighbouring views. No document saved."
End Sub

Private Sub InspectExistingDimensions()
    Dim dd As SldWorks.DisplayDimension, nextDimension As SldWorks.DisplayDimension
    Set dd = view.GetFirstDisplayDimension5
    Do While Not dd Is Nothing
        Set nextDimension = dd.GetNext5
        InspectExistingDimension dd
        Set dd = nextDimension
    Loop
    If unknownExisting > 0 Then
        blocked = True
        LogLine "STOP BEFORE CREATION: existing size(s) may be a target LENGTH but their frame/body attachments are unknown; inspect existing annotations."
    End If
End Sub

Private Function AnyExpectedLength(ByVal physical As Double) As Boolean
    Dim i As Long
    For i = 1 To jobCount
        If jobs(i).Expected > 0 Then
            If Abs(physical - jobs(i).Expected) <= LENGTH_TOL_M Then
                AnyExpectedLength = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub InspectExistingDimension(ByVal dd As SldWorks.DisplayDimension)
    On Error GoTo Unresolved
    Dim ann As SldWorks.Annotation, value As SldWorks.Dimension, attachments As Variant, p As Variant
    Dim physical As Double, valueKnown As Boolean, mayBeTargetLength As Boolean
    Set value = dd.GetDimension2(0)
    If value Is Nothing Then GoTo Unresolved
    physical = value.SystemValue
    valueKnown = True
    mayBeTargetLength = AnyExpectedLength(physical)
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then GoTo Unresolved
    p = ann.GetPosition
    If HasItems(p) Then occupied.Add Array(CDbl(p(0)), CDbl(p(1)), 0.015, 0.006)
    If mayBeTargetLength = False Then
        ignoredExisting = ignoredExisting + 1
        Exit Sub
    End If
    attachments = ann.GetAttachedEntities3
    If HasItems(attachments) = False Then GoTo Unresolved
    Dim count As Long, first As Long, i As Long, index As Long, bodyIndex As Long
    Dim source1 As Object, source2 As Object, item As Object, localEntity As Object
    Dim ownerKnown As Boolean, ownerOther As Boolean, allOwners As Boolean, overridden As Boolean
    first = LBound(attachments): count = UBound(attachments) - first + 1
    allOwners = True
    For i = first To UBound(attachments)
        Set item = attachments(i)
        Set localEntity = Nothing
        ownerKnown = False: ownerOther = False
        index = AttachmentSource(item, localEntity, ownerKnown, ownerOther)
        If ownerOther Then
            ignoredExisting = ignoredExisting + 1
            Exit Sub
        End If
        If ownerKnown = False Then allOwners = False
        If index > 0 Then
            If bodyIndex = 0 Then
                bodyIndex = index
            ElseIf bodyIndex <> index Then
                ignoredExisting = ignoredExisting + 1
                Exit Sub
            End If
        Else
            allOwners = False
        End If
        If i = first Then Set source1 = localEntity
        If i = first + 1 Then Set source2 = localEntity
    Next i
    If bodyIndex = 0 Then GoTo Unresolved
    If Abs(value.SystemValue - jobs(bodyIndex).Expected) > LENGTH_TOL_M Then
        ignoredExisting = ignoredExisting + 1
        Exit Sub
    End If
    If ApiBoolean(ann.IsDangling) Then
        ReviewExisting bodyIndex, "Existing dimension is dangling; it is not confirmed coverage."
        Exit Sub
    End If
    overridden = ApiBoolean(dd.GetOverride)
    If overridden Then
        ReviewExisting bodyIndex, "Existing dimension has a displayed-value override."
        Exit Sub
    End If
    If allOwners = False Then
        ReviewExisting bodyIndex, "Existing dimension source body known but frame instance/attachments unconfirmed."
        Exit Sub
    End If
    If count = 1 Then
        If Not source1 Is Nothing Then
            If TypeOf source1 Is SldWorks.Edge Then
                Dim edge As SldWorks.Edge, a As Variant, b As Variant, length As Double
                Set edge = source1
                If StraightGeometry(edge, a, b, length) Then
                    If Abs(length - jobs(bodyIndex).Expected) <= LENGTH_TOL_M Then
                        CoverExisting bodyIndex, "one attached frame-instance edge"
                        Exit Sub
                    End If
                End If
            End If
        End If
    ElseIf count = 2 Then
        If Not source1 Is Nothing And Not source2 Is Nothing Then
            If TypeOf source1 Is SldWorks.Vertex And TypeOf source2 Is SldWorks.Vertex Then
                Dim vertex1 As SldWorks.Vertex, vertex2 As SldWorks.Vertex
                Set vertex1 = source1: Set vertex2 = source2
                If VertexPairMatchesSourceEdge(bodyIndex, vertex1, vertex2) Then
                    CoverExisting bodyIndex, "two attached frame-instance vertices matching a real full-length edge"
                    Exit Sub
                End If
            End If
        End If
    End If
    ReviewExisting bodyIndex, "Existing dimension value matches but attachment geometry is not a confirmed full-length edge."
    Exit Sub
Unresolved:
    Dim detail As String
    detail = Err.Description
    If valueKnown And (mayBeTargetLength = False) Then
        ignoredExisting = ignoredExisting + 1
        LogLine "EXISTING ignored: physical value differs from every target LENGTH."
    ElseIf bodyIndex > 0 Then
        ReviewExisting bodyIndex, "Existing dimension API failed after body identification: " & detail
    Else
        unresolvedExisting = unresolvedExisting + 1
        unknownExisting = unknownExisting + 1
        LogLine "EXISTING UNCONFIRMED: possible target LENGTH has no trustworthy body attachment; creation barrier. " & detail
    End If
End Sub

Private Function AttachmentSource(ByVal item As Object, ByRef localEntity As Object, _
        ByRef ownerKnown As Boolean, ByRef ownerOther As Boolean) As Long
    On Error GoTo Unconfirmed
    ownerKnown = False: ownerOther = False
    Set localEntity = Nothing
    If item Is Nothing Then Exit Function
    ownerKnown = AttachmentOwner(item, ownerOther)
    If ownerOther Then Exit Function
    If TypeOf item Is SldWorks.Edge Then
        Dim attachedEdge As SldWorks.Edge, attachedCurve As SldWorks.Curve
        Dim attachedLine As Boolean
        Set attachedEdge = item
        Set attachedCurve = attachedEdge.GetCurve
        If attachedCurve Is Nothing Then Exit Function
        attachedLine = ApiBoolean(attachedCurve.IsLine)
        If attachedLine = False Then
            LogLine "EXISTING curved edge attachment unsupported; reverse correspondence not called"
            Exit Function
        End If
    End If
    Set localEntity = frameModel.Extension.GetCorrespondingEntity2(item)
    If localEntity Is Nothing Then Exit Function
    Dim body As SldWorks.Body2
    If TypeOf localEntity Is SldWorks.Edge Then
        Dim edge As SldWorks.Edge
        Set edge = localEntity: Set body = edge.GetBody
    ElseIf TypeOf localEntity Is SldWorks.Vertex Then
        Dim vertex As SldWorks.Vertex, edges As Variant, i As Long, found As Long, nextIndex As Long
        Set vertex = localEntity
        edges = vertex.GetEdges
        If HasItems(edges) = False Then Exit Function
        For i = LBound(edges) To UBound(edges)
            Set edge = edges(i)
            Set body = edge.GetBody
            If body Is Nothing Then Exit Function
            nextIndex = FindJobForBody(body)
            If nextIndex = 0 Then Exit Function
            If found = 0 Then found = nextIndex
            If found <> nextIndex Then Exit Function
        Next i
        AttachmentSource = found
        Exit Function
    ElseIf TypeOf localEntity Is SldWorks.Face2 Then
        Dim face As SldWorks.Face2
        Set face = localEntity: Set body = face.GetBody
    End If
    If Not body Is Nothing Then AttachmentSource = FindJobForBody(body)
    Exit Function
Unconfirmed:
    LogLine "EXISTING attachment unconfirmed: " & Err.Number & ": " & Err.Description
End Function

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

Private Function VertexPairMatchesSourceEdge(ByVal index As Long, ByVal va As SldWorks.Vertex, ByVal vb As SldWorks.Vertex) As Boolean
    Dim a As Variant, b As Variant, p As Variant, q As Variant, edges As Variant
    Dim edge As SldWorks.Edge, length As Double, i As Long
    a = va.GetPoint: b = vb.GetPoint
    edges = jobs(index).Body.GetEdges
    If HasItems(edges) = False Then Exit Function
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        If StraightGeometry(edge, p, q, length) Then
            If Abs(length - jobs(index).Expected) <= LENGTH_TOL_M Then
                If (SamePoint(a, p) And SamePoint(b, q)) Or (SamePoint(a, q) And SamePoint(b, p)) Then
                    VertexPairMatchesSourceEdge = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function SamePoint(ByVal a As Variant, ByVal b As Variant) As Boolean
    SamePoint = (Abs(a(0) - b(0)) <= LENGTH_TOL_M And Abs(a(1) - b(1)) <= LENGTH_TOL_M And Abs(a(2) - b(2)) <= LENGTH_TOL_M)
End Function

Private Sub CoverExisting(ByVal index As Long, ByVal proof As String)
    If jobs(index).State = "pending" Or Left$(jobs(index).Reason, 9) = "Existing " Then
        jobs(index).State = "already"
        jobs(index).Reason = "Confirmed existing dimension: " & proof
        LogLine "ALREADY body=[" & jobs(index).Body.Name & "]; proof=" & proof
    End If
End Sub

Private Sub ReviewExisting(ByVal index As Long, ByVal reason As String)
    unresolvedExisting = unresolvedExisting + 1
    LogLine "EXISTING REVIEW body=[" & jobs(index).Body.Name & "]: " & reason
    If jobs(index).State = "pending" Then
        jobs(index).State = "skipped"
        jobs(index).Reason = reason
    End If
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
    Dim selected As Boolean
    LogLine "BEGIN cleanup current new annotation Select3"
    selected = ApiBoolean(ann.Select3(False, Nothing))
    LogLine "DONE cleanup current new annotation Select3; returned=" & selected
    If selected = False Then Exit Function
    If drawing.SelectionManager.GetSelectedObjectCount2(-1) <> 1 Then Exit Function
    LogLine "BEGIN cleanup current new dimension DeleteSelection2"
    RemoveNewDimension = ApiBoolean(drawing.Extension.DeleteSelection2(0))
    LogLine "DONE cleanup current new dimension DeleteSelection2; returned=" & RemoveNewDimension
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

Private Function ReadSelectedFrameEdge(ByVal sm As SldWorks.SelectionMgr, _
        ByRef selectedView As SldWorks.View) As SldWorks.Edge
    Dim count As Long, edges As Long, i As Long, selectionType As Long
    Dim item As Object, candidate As SldWorks.Edge, itemView As SldWorks.View
    count = sm.GetSelectedObjectCount2(-1)
    LogLine "SELECTION total=" & count
    If count = 0 Then Err.Raise vbObjectError + 3, , _
        "Selection is empty at macro start. Click a frame edge, then run via Tools > Macro > Run. " & _
        "Select the frame edge in this drawing before running the macro."
    For i = 1 To count
        selectionType = sm.GetSelectedObjectType3(i, -1)
        Set item = sm.GetSelectedObject6(i, -1)
        Set itemView = sm.GetSelectedObjectsDrawingView2(i, -1)
        If i <= 30 Then
            LogLine "SELECTION index=" & i & "; API type=" & selectionType & "; object=" & TypeName(item)
            If itemView Is Nothing Then
                LogLine "  drawing view=Nothing"
            Else
                LogLine "  drawing view=[" & itemView.GetName2 & "]"
            End If
        End If
        If selectionType = swSelEDGES Then
            edges = edges + 1
            If edges = 1 Then
                If Not item Is Nothing Then
                    If TypeOf item Is SldWorks.Edge Then Set candidate = item
                End If
                Set selectedView = itemView
            End If
        End If
    Next i
    LogLine "SELECTION edge count=" & edges & "; other selections=" & (count - edges)
    If edges <> 1 Then Err.Raise vbObjectError + 3, , _
        "Expected ONE edge; found edges=" & edges & "; total selections=" & count & _
        ". Press Esc in the drawing, click one frame edge, then run the macro."
    If candidate Is Nothing Then Err.Raise vbObjectError + 4, , "Selected edge object unavailable or not SldWorks.Edge."
    If selectedView Is Nothing Then Err.Raise vbObjectError + 5, , "Selected edge has no drawing view."
    Set ReadSelectedFrameEdge = candidate
End Function

Private Sub OpenLog()
    On Error GoTo NoFile
    Dim fs As Object, stream As Object
    logAvailable = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    outputPath = fs.BuildPath(fs.GetSpecialFolder(2), "FrameCutListDimensionsV2-" & fs.GetTempName & ".txt")
    Set stream = fs.CreateTextFile(outputPath, False, True)
    stream.Close
    logAvailable = True
    Exit Sub
NoFile:
    outputPath = "VBA Immediate window (Ctrl+G); log file unavailable"
    Debug.Print PREFIX & " log file error: " & Err.Description
End Sub

Private Sub LogLine(ByVal message As String)
    Debug.Print PREFIX & " " & message
    If logAvailable Then WriteLogLine message
End Sub

Private Sub WriteLogLine(ByVal message As String)
    On Error GoTo WriteFailed
    Dim fs As Object, stream As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    ' ForAppending=8, TristateTrue=-1: close each UTF-16 line before native calls.
    Set stream = fs.OpenTextFile(outputPath, 8, False, -1)
    stream.WriteLine PREFIX & " " & message
    stream.Close
    Exit Sub
WriteFailed:
    logAvailable = False
    Debug.Print PREFIX & " log file write failed: " & Err.Description
End Sub

Private Sub CloseLog()
    ' No stream remains open between LogLine calls.
    logAvailable = False
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


Private Function ReadSelectedTubeGeometry(ByVal sm As SldWorks.SelectionMgr, _
        ByRef selectedView As SldWorks.View, ByRef selectedIndex As Long) As Object
    Dim count As Long, geometryCount As Long, i As Long, selectionType As Long
    Dim item As Object, candidate As Object, itemView As SldWorks.View
    count = sm.GetSelectedObjectCount2(-1)
    LogLine "SELECTION total=" & count
    If count = 0 Then Err.Raise vbObjectError + 3, , "Select one frame edge or silhouette in the drawing before running."
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
    LogLine "SELECTION frame geometry count=" & geometryCount & "; selected index=" & selectedIndex & "; other selections=" & (count - geometryCount)
    If geometryCount <> 1 Then Err.Raise vbObjectError + 3, , "Expected ONE frame edge or silhouette; found=" & geometryCount
    If candidate Is Nothing Then Err.Raise vbObjectError + 4, , "Selected frame geometry object unavailable or unsupported."
    If selectedView Is Nothing Then Err.Raise vbObjectError + 5, , "Selected frame geometry has no drawing view."
    Set ReadSelectedTubeGeometry = candidate
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


Private Function MapSourceBoundaryToView(ByVal source As SldWorks.Edge) As SldWorks.Edge
    Dim modelEntity As Object, drawingEntity As Object
    If frameInstance Is Nothing Then
        Set modelEntity = source
    Else
        Set modelEntity = frameInstance.GetCorresponding(source)
    End If
    If modelEntity Is Nothing Then
        LogLine "SOURCE FACE boundary has no corresponding entity in the selected instance"
        Exit Function
    End If
    If Not TypeOf modelEntity Is SldWorks.Edge Then Exit Function
    Set drawingEntity = view.GetCorresponding(modelEntity)
    If Not drawingEntity Is Nothing Then
        If TypeOf drawingEntity Is SldWorks.Edge Then
            Set MapSourceBoundaryToView = drawingEntity
            Exit Function
        End If
        LogLine "SOURCE FACE view correspondence returned a non-edge; candidate rejected"
        Exit Function
    End If
    ' SelectEntity supports model-context entities in a specified view.
    ' VerifySingleSelection must still prove view, actual instance and body
    ' before the sole AddDimension2. A successful selection is not visibility.
    Set MapSourceBoundaryToView = modelEntity
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

Private Function EdgeHasSupportFace(ByVal edge As SldWorks.Edge) As Boolean
    Dim faces As Variant, face As SldWorks.Face2, i As Long
    faces = edge.GetTwoAdjacentFaces2
    If HasItems(faces) = False Then Exit Function
    For i = LBound(faces) To UBound(faces)
        Set face = faces(i)
        If Not face Is Nothing Then
            If app.IsSame(face, supportFace) = swObjectSame Then
                EdgeHasSupportFace = True
                Exit Function
            End If
        End If
    Next i
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


Private Sub EnsureInputPreference()
    If preferenceCaptured = False Then
        oldInputPreference = ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate))
        preferenceCaptured = True
        app.SetUserPreferenceToggle swInputDimValOnCreate, False
    End If
    If ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate)) Then _
        Err.Raise vbObjectError + 33, , "Cannot disable dimension value input."
End Sub

Private Function NeedSupportPhase() As Boolean
    Dim i As Long
    For i = 1 To jobCount
        If jobs(i).State = "pending" And jobs(i).NormalCount = 0 Then
            NeedSupportPhase = True
            Exit Function
        End If
    Next i
End Function

Private Sub DiscoverSupportFaces()
    ' SOURCE ONLY. No drawing-edge reverse mapping or visibility queries here.
    ' Every face is obtained directly from its queued source cut-list body.
    sourcePrepStarted = Timer
    sourcePrepTimedOut = False
    LogLine "SOURCE PREPARATION source cut-list geometry only; face cap=" & MAX_SOURCE_FACES & _
        "; retained candidates/body=" & MAX_SOURCE_CANDIDATES & "; budget seconds=" & SOURCE_PREP_BUDGET_S & "; visibility NOT proven"
    Dim i As Long
    For i = 1 To jobCount
        If jobs(i).State = "pending" And jobs(i).NormalCount = 0 Then
            If SourcePrepBudgetExceeded Then
                SkipUnpreparedSourceJobs i
                Exit Sub
            End If
            PrepareSourceBody i
            If sourcePrepTimedOut Then
                SkipUnpreparedSourceJobs i
                Exit Sub
            End If
        End If
    Next i
    LogLine "SOURCE PREPARATION completed within budget"
End Sub




Private Sub CollectFaceEdges(ByVal index As Long, ByVal face As SldWorks.Face2, ByVal faceIndex As Long)
    Dim edges As Variant, i As Long, source As SldWorks.Edge, accepted As Long
    edges = face.GetEdges
    If HasItems(edges) = False Then Exit Sub
    For i = LBound(edges) To UBound(edges)
        If SourcePrepBudgetExceeded Then Exit Sub
        Set source = edges(i)
        If AddFaceEdge(index, source) Then accepted = accepted + 1
    Next i
    If SourcePrepBudgetExceeded Then Exit Sub
    LogLine "SOURCE cylinder face body=[" & jobs(index).Body.Name & "]; face index=" & faceIndex & _
        "; real edges=" & (UBound(edges) - LBound(edges) + 1) & "; accepted boundaries=" & accepted & "; retained=" & jobs(index).Source.Count
End Sub

Private Function AddFaceEdge(ByVal index As Long, ByVal source As SldWorks.Edge) As Boolean
    On Error GoTo Rejected
    Dim body As SldWorks.Body2, a As Variant, b As Variant, length As Double
    Dim mapped As SldWorks.Edge, i As Long, prior As SldWorks.Edge, score As Double, insertAt As Long
    Dim box As Variant, pa As Variant, pb As Variant, dx As Double, dy As Double
    Set body = source.GetBody
    If body Is Nothing Then Exit Function
    If FindJobForBody(body) <> index Then Exit Function
    For i = 1 To jobs(index).Source.Count
        Set prior = jobs(index).Source(i)
        If app.IsSame(prior, source) = swObjectSame Then Exit Function
    Next i
    If StraightGeometry(source, a, b, length) = False Then Exit Function
    If Abs(length - jobs(index).Expected) > LENGTH_TOL_M Then Exit Function
    If PlaneCylinderBoundary(source) = False Then Exit Function
    If EdgeHasSupportFace(source) = False Then Exit Function
    pa = ProjectPoint(a): pb = ProjectPoint(b)
    dx = pb(0) - pa(0): dy = pb(1) - pa(1)
    If Sqr(dx * dx + dy * dy) < 0.00001 Then Exit Function
    ' Never share another body's centre. This is only a radial ranking heuristic.
    box = jobs(index).Body.GetBodyBox
    If HasItems(box) = False Then Exit Function
    sourceCentre = Array((box(0) + box(3)) / 2#, (box(1) + box(4)) / 2#, (box(2) + box(5)) / 2#)
    score = RadialScore(a, b, length)
    Set mapped = MapSourceBoundaryToView(source)
    If mapped Is Nothing Then Exit Function
    For i = 1 To jobs(index).Scores.Count
        If score > CDbl(jobs(index).Scores(i)) Then
            insertAt = i
            Exit For
        End If
    Next i
    If insertAt > 0 Then
        jobs(index).Source.Add source, , insertAt
        jobs(index).Visible.Add mapped, , insertAt
        jobs(index).Scores.Add score, , insertAt
    Else
        jobs(index).Source.Add source
        jobs(index).Visible.Add mapped
        jobs(index).Scores.Add score
    End If
    ' Collections remain aligned after descending-score insertion and trimming.
    If jobs(index).Source.Count > MAX_SOURCE_CANDIDATES Then
        jobs(index).Source.Remove jobs(index).Source.Count
        jobs(index).Visible.Remove jobs(index).Visible.Count
        jobs(index).Scores.Remove jobs(index).Scores.Count
    End If
    AddFaceEdge = True
    Exit Function
Rejected:
    LogLine "SOURCE FACE edge rejected: " & Err.Number & ": " & Err.Description
End Function

Private Sub FinishAndValidate()
    Dim i As Long
    For i = 1 To jobCount
        If Not jobs(i).CreatedDD Is Nothing Then ValidateRetainedDimension i
    Next i
    If RestorePreference = False Then
        blocked = True
        LogLine "STOP: dimension value input setting could not be restored."
    End If
End Sub

Private Sub ValidateRetainedDimension(ByVal index As Long)
    On Error GoTo Invalid
    Dim dd As SldWorks.DisplayDimension, ann As SldWorks.Annotation, value As SldWorks.Dimension
    Dim p As Variant, physical As Double, dangling As Boolean, overridden As Boolean
    Set dd = jobs(index).CreatedDD
    Set value = dd.GetDimension2(0)
    If value Is Nothing Then Err.Raise vbObjectError + 110, , "New dimension readback unavailable during final validation."
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Err.Raise vbObjectError + 111, , "New annotation unavailable during final validation."
    p = ann.GetPosition
    RequirePosition p
    physical = value.SystemValue
    dangling = ApiBoolean(ann.IsDangling): overridden = ApiBoolean(dd.GetOverride)
    LogLine "FINAL VALIDATION body=[" & jobs(index).Body.Name & "]; actual mm=" & physical * 1000# & _
        "; expected mm=" & jobs(index).Expected * 1000# & "; dangling=" & dangling & "; override=" & overridden & _
        "; XY m=(" & p(0) & ", " & p(1) & ")"
    If Abs(physical - jobs(index).Expected) > LENGTH_TOL_M Then Err.Raise vbObjectError + 112, , "True LENGTH changed during final validation."
    If dangling Or overridden Then Err.Raise vbObjectError + 113, , "New dimension is dangling or overridden during final validation."
    If Abs(p(0) - jobs(index).TargetX) > POSITION_TOL_M Or Abs(p(1) - jobs(index).TargetY) > POSITION_TOL_M Then _
        Err.Raise vbObjectError + 114, , "New text position changed during final validation."
    ' Do not collision-test against our OWN occupied rectangle after reservation.
    If p(0) - jobs(index).HalfWidth < SHEET_MARGIN_M Or p(0) + jobs(index).HalfWidth > sheetWidth - SHEET_MARGIN_M Or _
        p(1) - jobs(index).HalfHeight < SHEET_MARGIN_M Or p(1) + jobs(index).HalfHeight > sheetHeight - SHEET_MARGIN_M Then _
        Err.Raise vbObjectError + 115, , "New text estimate is outside sheet during final validation."
    jobs(index).Reason = "True cut-list LENGTH, XY and non-dangling annotation verified; view style unchanged."
    Exit Sub
Invalid:
    Dim reason As String
    reason = CStr(Err.Number) & ": " & Err.Description
    jobs(index).State = "failed": jobs(index).Reason = reason
    LogLine "FINAL VALIDATION invalid new dimension: " & reason
    If RemoveNewDimension(jobs(index).CreatedDD) = False Then
        blocked = True
        jobs(index).Reason = reason & "; CLEANUP FAILED; inspect current new dimension manually."
    Else
        Set jobs(index).CreatedDD = Nothing
    End If
End Sub


Private Sub PrepareSourceBody(ByVal index As Long)
    On Error GoTo Unsupported
    Dim faces As Variant, face As SldWorks.Face2, owner As SldWorks.Body2
    Dim surface As SldWorks.Surface, seen As Collection, previous As SldWorks.Face2
    Dim i As Long, j As Long, count As Long, duplicate As Boolean, cylinders As Long
    Set seen = New Collection
    LogLine "BEGIN source Body.GetFaces; body=[" & jobs(index).Body.Name & "]"
    faces = jobs(index).Body.GetFaces
    If HasItems(faces) Then count = UBound(faces) - LBound(faces) + 1
    LogLine "DONE source Body.GetFaces; body=[" & jobs(index).Body.Name & "]; faces=" & count
    If count = 0 Then Err.Raise vbObjectError + 120, , "Source body has no accessible faces."
    If count > MAX_SOURCE_FACES Then Err.Raise vbObjectError + 121, , "Source body exceeds the face cap; entire body skipped."
    For i = LBound(faces) To UBound(faces)
        If SourcePrepBudgetExceeded Then Exit Sub
        Set face = faces(i)
        If Not face Is Nothing Then
            duplicate = False
            For j = 1 To seen.Count
                If SourcePrepBudgetExceeded Then Exit Sub
                Set previous = seen(j)
                If app.IsSame(previous, face) = swObjectSame Then
                    duplicate = True
                    Exit For
                End If
            Next j
            If duplicate = False Then
                seen.Add face
                Set owner = face.GetBody
                If owner Is Nothing Then Err.Raise vbObjectError + 122, , "Source face body unavailable."
                If FindJobForBody(owner) <> index Then Err.Raise vbObjectError + 123, , "Source face is not owned by its queued body."
                Set surface = face.GetSurface
                If Not surface Is Nothing Then
                    If ApiBoolean(surface.IsCylinder) Then
                        cylinders = cylinders + 1
                        Set supportFace = face
                        CollectFaceEdges index, face, i
                        If sourcePrepTimedOut Then Exit Sub
                    End If
                End If
            End If
        End If
    Next i
    If SourcePrepBudgetExceeded Then Exit Sub
    LogLine "SOURCE body prepared=[" & jobs(index).Body.Name & "]; unique faces=" & seen.Count & _
        "; cylinder faces=" & cylinders & "; retained ranked candidates=" & jobs(index).Source.Count & "; visibility NOT proven"
    Exit Sub
Unsupported:
    jobs(index).State = "skipped"
    jobs(index).Reason = "Source-only tube preparation unsupported: " & Err.Description
    ClearSourceCandidates index
    LogLine "SOURCE body skipped=[" & jobs(index).Body.Name & "]; " & jobs(index).Reason
End Sub

Private Function SourcePrepBudgetExceeded() As Boolean
    Dim elapsed As Double
    elapsed = Timer - sourcePrepStarted
    If elapsed < 0# Then elapsed = elapsed + 86400#
    If elapsed >= SOURCE_PREP_BUDGET_S Then
        If sourcePrepTimedOut = False Then LogLine "SOURCE PREPARATION budget reached; elapsed seconds=" & elapsed
        sourcePrepTimedOut = True
    End If
    SourcePrepBudgetExceeded = sourcePrepTimedOut
End Function

Private Sub SkipUnpreparedSourceJobs(ByVal first As Long)
    Dim i As Long
    For i = first To jobCount
        If jobs(i).State = "pending" And jobs(i).NormalCount = 0 Then
            ClearSourceCandidates i
            jobs(i).State = "skipped"
            jobs(i).Reason = "Source preparation time budget reached; partial candidates discarded, no legacy discovery fallback."
            LogLine "SOURCE budget skip body=[" & jobs(i).Body.Name & "]"
        End If
    Next i
End Sub

Private Sub ClearSourceCandidates(ByVal index As Long)
    Set jobs(index).Source = New Collection
    Set jobs(index).Visible = New Collection
    Set jobs(index).Scores = New Collection
End Sub
