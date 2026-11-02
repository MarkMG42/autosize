Attribute VB_Name = "FrameCutListDimensionTrial"
Option Explicit

' Creates ONE true dimension for the ONE selected full-length frame edge.
' Standalone module: standard SldWorks and SwConst references only.
' No bulk processing, numeric text overrides, model rebuilds or document saves.
Private Const PREFIX As String = "FrameCutListDimensionTrial v1.1"
Private Const LENGTH_TOL_M As Double = 0.000001
Private Const POSITION_TOL_M As Double = 0.00001
Private Const CUTLIST_MM_TO_M As Double = 0.001
Private Const SHEET_MARGIN_M As Double = 0.001
Private app As SldWorks.SldWorks
Private drawing As SldWorks.ModelDoc2
Private view As SldWorks.View
Private frameModel As SldWorks.ModelDoc2
Private frameInstance As SldWorks.Component2
Private selectedBody As SldWorks.Body2
Private visited As Collection, occupied As Collection
Private output As Object
Private outputPath As String
Private matchedFolder As SldWorks.Feature
Private folderMatches As Long
Private expectedLength As Double
Private preferenceCaptured As Boolean, oldInputPreference As Boolean
Private sheetWidth As Double, sheetHeight As Double
Private sheetA As Variant, sheetB As Variant, sheetCentre As Variant

Public Sub main()
    On Error GoTo Failed
    Dim dd As SldWorks.DisplayDimension, attempted As Boolean, selectionTouched As Boolean
    CloseLog
    Set drawing = Nothing: Set view = Nothing: Set frameModel = Nothing
    Set frameInstance = Nothing: Set selectedBody = Nothing: Set matchedFolder = Nothing
    Set visited = New Collection: Set occupied = New Collection
    folderMatches = 0: expectedLength = 0#: preferenceCaptured = False
    outputPath = ""
    OpenLog
    LogLine "START " & CStr(Now)
    LogLine "ONE dimension trial; rerun may create a duplicate. No model is saved."
    Set app = Application.SldWorks
    Set drawing = app.ActiveDoc
    If drawing Is Nothing Then Err.Raise vbObjectError + 1, , "Open the drawing first."
    If drawing.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 2, , "Active document must be a drawing."
    Dim sm As SldWorks.SelectionMgr, selectedEdge As SldWorks.Edge
    Set sm = drawing.SelectionManager
    Set selectedEdge = ReadSelectedFrameEdge(sm, view)
    If ApiBoolean(view.ProjectedDimensions) Then Err.Raise vbObjectError + 30, , _
        "Set this view's dimension type to True (not Projected), then rerun. View mode was not changed."
    ResolveContext selectedEdge
    Dim converted As Object, sourceEdge As SldWorks.Edge
    Set converted = frameModel.Extension.GetCorrespondingEntity2(selectedEdge)
    If converted Is Nothing Then Err.Raise vbObjectError + 31, , "Selected edge has no source-part correspondence."
    If Not TypeOf converted Is SldWorks.Edge Then Err.Raise vbObjectError + 32, , "Source correspondence is not an Edge."
    Set sourceEdge = converted
    Dim curve As SldWorks.Curve, va As SldWorks.Vertex, vb As SldWorks.Vertex
    Set curve = sourceEdge.GetCurve
    If curve Is Nothing Then Err.Raise vbObjectError + 33, , "Source curve unavailable."
    Dim rawLine As Variant, isLineCached As Boolean, curveIdentity As Long
    ' Some COM Boolean values can carry noncanonical True bits. Read ONCE;
    ' log the raw numeric representation and compare with False, never Not.
    rawLine = curve.IsLine
    curveIdentity = curve.Identity
    LogLine "SOURCE EDGE raw IsLine=" & CStr(rawLine) & "; TypeName=" & TypeName(rawLine) & _
        "; VarType=" & VarType(rawLine) & "; numeric=" & CLng(rawLine)
    isLineCached = ApiBoolean(rawLine)
    LogLine "SOURCE EDGE curve Identity=" & curveIdentity & "; IsLine cached=" & isLineCached & _
        "; canonical numeric=" & CLng(isLineCached)
    If isLineCached = False Then Err.Raise vbObjectError + 34, , "Selected source edge is not straight (cached result)."
    Set va = sourceEdge.GetStartVertex: Set vb = sourceEdge.GetEndVertex
    If va Is Nothing Or vb Is Nothing Then Err.Raise vbObjectError + 35, , "Source edge vertices unavailable."
    Dim a As Variant, b As Variant, edgeLength As Double
    a = va.GetPoint: b = vb.GetPoint
    edgeLength = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2 + (b(2) - a(2)) ^ 2)
    Set selectedBody = sourceEdge.GetBody
    If selectedBody Is Nothing Then Err.Raise vbObjectError + 36, , "Source body unavailable."
    LogLine "SELECTED SOURCE BODY=[" & selectedBody.Name & "]; self IsSame=" & app.IsSame(selectedBody, selectedBody)
    If app.IsSame(selectedBody, selectedBody) <> swObjectSame Then _
        Err.Raise vbObjectError + 37, , "Source body identity self-check failed."
    FindCutListMembership
    If folderMatches <> 1 Then Err.Raise vbObjectError + 38, , "Expected one cut-list body membership; found=" & folderMatches
    ReadCutListLength matchedFolder
    LogLine "CHECK selected edge m=" & CStr(edgeLength) & "; cut-list length m=" & CStr(expectedLength)
    If Abs(edgeLength - expectedLength) > LENGTH_TOL_M Then Err.Raise vbObjectError + 39, , _
        "Choose a full longitudinal edge of this body. Selected edge mm=" & CStr(edgeLength * 1000#) & _
        "; cut-list LENGTH mm=" & CStr(expectedLength * 1000#) & ". No dimension was created."
    If edgeLength <= LENGTH_TOL_M Then Err.Raise vbObjectError + 40, , "Degenerate edge."
    sheetA = ProjectPoint(a): sheetB = ProjectPoint(b)
    Dim box As Variant, centre As Variant
    box = selectedBody.GetBodyBox
    If Not HasItems(box) Then Err.Raise vbObjectError + 41, , "Source body bounding box unavailable."
    centre = Array((box(0) + box(3)) / 2#, (box(1) + box(4)) / 2#, (box(2) + box(5)) / 2#)
    sheetCentre = ProjectPoint(centre)
    LogLine "SHEET p0 m=(" & sheetA(0) & ", " & sheetA(1) & "); p1 m=(" & sheetB(0) & ", " & sheetB(1) & ")"
    Dim drawingDoc As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Set drawingDoc = drawing
    Set sheet = drawingDoc.GetCurrentSheet
    sheet.GetSize sheetWidth, sheetHeight
    If sheetWidth <= 0 Or sheetHeight <= 0 Then Err.Raise vbObjectError + 42, , "Invalid sheet size."
    SeedOccupied
    Dim x As Double, y As Double, height As Double, halfWidth As Double
    height = 0.004: halfWidth = EstimateHalfWidth(expectedLength, height)
    If Not FindPosition(halfWidth, height, x, y) Then Err.Raise vbObjectError + 43, , "No free near-edge text slot inside sheet margins."
    oldInputPreference = ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate))
    preferenceCaptured = True
    app.SetUserPreferenceToggle swInputDimValOnCreate, False
    If ApiBoolean(app.GetUserPreferenceToggle(swInputDimValOnCreate)) Then _
        Err.Raise vbObjectError + 44, , "Cannot disable dimension value input."
    selectionTouched = True
    drawing.ClearSelection2 True
    If ApiBoolean(view.SelectEntity(selectedEdge, False)) = False Then Err.Raise vbObjectError + 45, , "Cannot select original edge in its drawing view."
    VerifySingleSelection sm
    attempted = True
    Set dd = drawing.AddDimension2(x, y, 0#)
    If dd Is Nothing Then Err.Raise vbObjectError + 46, , "AddDimension2 returned Nothing. Creation outcome uncertain."
    Dim value As SldWorks.Dimension, ann As SldWorks.Annotation
    Set value = dd.GetDimension2(0)
    If value Is Nothing Then Err.Raise vbObjectError + 47, , "Dimension value readback unavailable."
    LogLine "DIMENSION readback m=" & CStr(value.SystemValue) & "; expected m=" & CStr(expectedLength)
    If Abs(value.SystemValue - expectedLength) > LENGTH_TOL_M Then _
        Err.Raise vbObjectError + 48, , "Created dimension is not the cut-list true length."
    Set ann = dd.GetAnnotation
    If ann Is Nothing Then Err.Raise vbObjectError + 49, , "New dimension annotation unavailable."
    Dim fmt As SldWorks.TextFormat
    Set fmt = ann.GetTextFormat(0)
    If Not fmt Is Nothing Then
        If fmt.CharHeight > height Then height = fmt.CharHeight
    End If
    halfWidth = EstimateHalfWidth(expectedLength, height)
    If Not FindPosition(halfWidth, height, x, y) Then Err.Raise vbObjectError + 50, , "No near-edge text slot for actual font height."
    If ApiBoolean(ann.SetPosition2(x, y, 0#)) = False Then Err.Raise vbObjectError + 51, , "Cannot position new dimension."
    Dim p As Variant
    p = ann.GetPosition
    If Not HasItems(p) Then Err.Raise vbObjectError + 52, , "New annotation position readback unavailable."
    LogLine "TEXT requested m=(" & x & ", " & y & "); actual m=(" & p(0) & ", " & p(1) & ")"
    If Abs(p(0) - x) > POSITION_TOL_M Or Abs(p(1) - y) > POSITION_TOL_M Then _
        Err.Raise vbObjectError + 53, , "SolidWorks did not retain requested annotation position."
    If Not PositionFits(CDbl(p(0)), CDbl(p(1)), halfWidth, height) Then _
        Err.Raise vbObjectError + 54, , "New annotation readback is outside estimated placement limits."
    LogLine "FINAL DIMENSION readback m=" & CStr(value.SystemValue) & "; expected m=" & CStr(expectedLength)
    If Abs(value.SystemValue - expectedLength) > LENGTH_TOL_M Then _
        Err.Raise vbObjectError + 72, , "Dimension changed after positioning; true cut-list length no longer matches."
    drawing.ClearSelection2 True
    If Not RestorePreference Then Err.Raise vbObjectError + 55, , "Dimension input preference could not be restored."
    LogLine "END created=1; cutlist=[" & matchedFolder.Name & "]; body=[" & selectedBody.Name & "]; length mm=" & expectedLength * 1000#
    LogLine "Text bounds are estimates; inspect extension lines, text and geometry before saving."
    CloseLog
    MsgBox PREFIX & ": created ONE true dimension, mm=" & CStr(expectedLength * 1000#) & vbCrLf & _
        "Inspect its attachment and position. No document was saved." & vbCrLf & "Log: " & outputPath, vbInformation
    Exit Sub
Failed:
    Dim message As String, cleanupOK As Boolean, restored As Boolean
    message = CStr(Err.Number) & ": " & Err.Description
    LogLine "STOP " & message
    If Not dd Is Nothing Then
        cleanupOK = RemoveNewDimension(dd)
        LogLine "CLEANUP current new dimension removed=" & cleanupOK
        If Not cleanupOK Then LogLine "UNCERTAIN OUTCOME: inspect drawing manually; current dimension could not be removed."
    ElseIf attempted Then
        LogLine "UNCERTAIN OUTCOME: dimension creation attempted without a returned handle. Inspect drawing manually."
    Else
        LogLine "No dimension creation was attempted."
    End If
    restored = RestorePreference
    If selectionTouched Then ClearSelectionSafely
    LogLine "END stopped; preference restored=" & restored
    CloseLog
    MsgBox PREFIX & " stopped: " & message & vbCrLf & "Inspect the log/drawing before rerunning." & vbCrLf & "Log: " & outputPath, vbExclamation
End Sub

Private Function ApiBoolean(ByVal raw As Variant) As Boolean
    ' A comparison produces a canonical VBA Boolean for either COM True
    ' representation (1 or -1), rather than applying bitwise Not to raw bits.
    ApiBoolean = (CBool(raw) <> False)
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
        If drawingComp Is Nothing Then Err.Raise vbObjectError + 8, , "Selected drawing component unavailable."
        Set frameInstance = ResolveFrameInstance(reference, drawingComp)
        Set frameModel = frameInstance.GetModelDoc2
        If frameModel Is Nothing Then Err.Raise vbObjectError + 9, , "Resolve/open selected frame part first."
        CheckConfiguration frameModel, frameInstance.ReferencedConfiguration
        LogLine "FRAME INSTANCE=[" & frameInstance.Name2 & "]; config=[" & frameInstance.ReferencedConfiguration & "]"
    ElseIf reference.GetType = swDocPART Then
        Set frameModel = reference
    Else
        Err.Raise vbObjectError + 10, , "Only assembly and part references supported."
    End If
    If frameModel.GetType <> swDocPART Then Err.Raise vbObjectError + 11, , "Selected frame must be a PART."
    LogLine "FRAME MODEL=[" & frameModel.GetTitle & "]; path=[" & frameModel.GetPathName & "]"
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

Private Sub CheckFolderMembership(ByVal feature As SldWorks.Feature)
    Dim folder As SldWorks.BodyFolder, bodies As Variant, body As SldWorks.Body2, i As Long
    Dim same As Long, sameName As Boolean
    Set folder = feature.GetSpecificFeature2
    If folder Is Nothing Then Err.Raise vbObjectError + 56, , "Cut-list body folder unavailable."
    bodies = folder.GetBodies
    If Not HasItems(bodies) Then Exit Sub
    For i = LBound(bodies) To UBound(bodies)
        Set body = bodies(i)
        If Not body Is Nothing Then
            same = app.IsSame(selectedBody, body)
            sameName = (StrComp(selectedBody.Name, body.Name, vbBinaryCompare) = 0)
            If same = swObjectSame Or sameName Then
                LogLine "BODY membership cutlist=[" & feature.Name & "]; selected=[" & selectedBody.Name & "]; folder body=[" & body.Name & "]; IsSame=" & same
                If same <> swObjectSame Or Not sameName Then _
                    Err.Raise vbObjectError + 57, , "Body identity/name consistency check failed."
                folderMatches = folderMatches + 1
                Set matchedFolder = feature
            End If
        End If
    Next i
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
    LogLine "LENGTH input unit required by this trial=mm; scaled to metres; geometry cross-check is mandatory."
End Sub

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

Private Function FindPosition(ByVal halfWidth As Double, ByVal halfHeight As Double, ByRef x As Double, ByRef y As Double) As Boolean
    Dim dx As Double, dy As Double, projected As Double, nx As Double, ny As Double
    Dim mx As Double, my As Double, dot As Double, side As Long, level As Long, offset As Double
    dx = sheetB(0) - sheetA(0): dy = sheetB(1) - sheetA(1)
    projected = Sqr(dx * dx + dy * dy)
    If projected < 0.00001 Then Exit Function
    mx = (sheetA(0) + sheetB(0)) / 2#: my = (sheetA(1) + sheetB(1)) / 2#
    nx = -dy / projected: ny = dx / projected
    dot = (mx - sheetCentre(0)) * nx + (my - sheetCentre(1)) * ny
    If dot < 0 Then nx = -nx: ny = -ny
    For level = 0 To 3
        offset = 0.006 + level * 0.002
        For side = 1 To -1 Step -2
            x = mx + side * nx * offset: y = my + side * ny * offset
            If PositionFits(x, y, halfWidth, halfHeight) Then
                FindPosition = True
                Exit Function
            End If
        Next side
    Next level
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

Private Function ReadSelectedFrameEdge(ByVal sm As SldWorks.SelectionMgr, _
        ByRef selectedView As SldWorks.View) As SldWorks.Edge
    Dim count As Long, edges As Long, i As Long, selectionType As Long
    Dim item As Object, candidate As SldWorks.Edge, itemView As SldWorks.View
    count = sm.GetSelectedObjectCount2(-1)
    LogLine "SELECTION total=" & count
    If count = 0 Then Err.Raise vbObjectError + 3, , _
        "Selection is empty at macro start. Click a frame edge, then run via Tools > Macro > Run. " & _
        "Select the frame edge in this drawing before running the trial."
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
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    outputPath = fs.BuildPath(fs.GetSpecialFolder(2), "FrameCutListDimensionTrial-" & fs.GetTempName & ".txt")
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
            Err.Raise vbObjectError + 13, , "Configuration mismatch. Trial does not switch configurations."
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
