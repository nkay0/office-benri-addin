Option Explicit

Private resizeRibbon As IRibbonUI
Private aspectLock As Boolean
Private ResizeAnchor As Long

Public Sub resizeRibbonOnLoad(ribbon As IRibbonUI)
    ResizeAnchor = 5
    aspectLock = False
    Set resizeRibbon = ribbon
    resizeRibbon.Invalidate
End Sub

Public Sub ResizeAspect_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = aspectLock
End Sub

Public Sub ResizeAspectChange(control As IRibbonControl, pressed As Boolean)
    aspectLock = pressed
End Sub

Public Sub ResizeAnchor_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (ResizeAnchor = CInt(Right(control.Id, 1)))
End Sub

Public Sub ResizeAnchorChange(control As IRibbonControl, pressed As Boolean)
    ResizeAnchor = CInt(Right(control.Id, 1))
    resizeRibbon.Invalidate
End Sub


'Stretch shrink to align
Public Sub StretchLeft()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim TargetRight As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Set TargetShape = SelectedShapes(1)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.Width = TargetPoints("Left") + TargetPoints("Width")
        TargetShape.Left = TargetShape.Left - GetVisualPoints(TargetShape)("Left")
    End If

    KeyPoint = GetVisualPoints(SelectedShapes(MaxNum))("Left")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetRight = TargetPoints("Left") + TargetPoints("Width")
        If TargetRight < KeyPoint Then
            TargetShape.Width = TargetShape.Width + KeyPoint - TargetRight
            TargetShape.Left = TargetPoints("Left") + TargetShape.Left - GetVisualPoints(TargetShape)("Left")
        Else
            TargetShape.Width = TargetShape.Width + TargetPoints("Left") - KeyPoint
            TargetShape.Left = KeyPoint + TargetShape.Left - GetVisualPoints(TargetShape)("Left")
        End If
    Next i
End Sub

Public Sub StretchRight()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim TargetRight As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Set TargetShape = SelectedShapes(1)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.Width = ActivePresentation.PageSetup.SlideWidth - TargetPoints("Left")
        TargetShape.Left = TargetPoints("Left") + TargetShape.Left - GetVisualPoints(TargetShape)("Left")
    End If

    SelectedShapes.LockAspectRatio = aspectLock
    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Left") + KeyPoints("Width")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetRight = TargetPoints("Left") + TargetPoints("Width")
        If TargetPoints("Left") < KeyPoint Then
            TargetShape.Width = TargetShape.Width + KeyPoint - TargetRight
            TargetShape.Left = TargetPoints("Left") + TargetShape.Left - GetVisualPoints(TargetShape)("Left")
        Else
            TargetShape.Width = TargetShape.Width + TargetPoints("Left") - KeyPoint
            TargetShape.Left = KeyPoint + TargetShape.Left - GetVisualPoints(TargetShape)("Left")
        End If
    Next i
End Sub

Public Sub StretchTop()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim TargetBottom As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Set TargetShape = SelectedShapes(1)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.Height = TargetPoints("Top") + TargetPoints("Height")
        TargetShape.Top = TargetShape.Top - GetVisualPoints(TargetShape)("Top")
    End If

    KeyPoint = GetVisualPoints(SelectedShapes(MaxNum))("Top")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetBottom = TargetPoints("Top") + TargetPoints("Height")
        If TargetBottom < KeyPoint Then
            TargetShape.Height = TargetShape.Height + KeyPoint - TargetBottom
            TargetShape.Top = TargetPoints("Top") + TargetShape.Top - GetVisualPoints(TargetShape)("Top")
        Else
            TargetShape.Height = TargetShape.Height + TargetPoints("Top") - KeyPoint
            TargetShape.Top = KeyPoint + TargetShape.Top - GetVisualPoints(TargetShape)("Top")
        End If
    Next i
End Sub

Public Sub StretchBottom()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim TargetBottom As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Set TargetShape = SelectedShapes(1)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.Height = ActivePresentation.PageSetup.SlideHeight - TargetPoints("Top")
        TargetShape.Top = TargetPoints("Top") + TargetShape.Top - GetVisualPoints(TargetShape)("Top")
    End If

    SelectedShapes.LockAspectRatio = aspectLock
    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Top") + KeyPoints("Height")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetBottom = TargetPoints("Top") + TargetPoints("Height")
        If TargetPoints("Top") < KeyPoint Then
            TargetShape.Height = TargetShape.Height + KeyPoint - TargetBottom
            TargetShape.Top = TargetPoints("Top") + TargetShape.Top - GetVisualPoints(TargetShape)("Top")
        Else
            TargetShape.Height = TargetShape.Height + TargetPoints("Top") - KeyPoint
            TargetShape.Top = KeyPoint + TargetShape.Top - GetVisualPoints(TargetShape)("Top")
        End If
    Next i
End Sub


'Equalize
Public Sub ResizeToSameWidth()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyWidth As Double
    Dim TargetShape As Shape
    Dim OriginalPosition() As Double, NewPosition() As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    SelectedShapes.LockAspectRatio = aspectLock
    KeyWidth = GetVisualPoints(SelectedShapes(MaxNum))("Width")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Width = KeyWidth
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub

Public Sub ResizeToSameHeight()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyHeight As Double
    Dim TargetShape As Shape
    Dim OriginalPosition() As Double, NewPosition() As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    SelectedShapes.LockAspectRatio = aspectLock
    KeyHeight = GetVisualPoints(SelectedShapes(MaxNum))("Height")
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Height = KeyHeight
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub

Public Sub ResizeToSameHeightAndWidth()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim TargetShape As Shape
    Dim OriginalPosition() As Double, NewPosition() As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    SelectedShapes.LockAspectRatio = msoFalse
    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Height = KeyPoints("Height")
        TargetShape.Width = KeyPoints("Width")
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub


'ResizeByRatio
Public Sub ResizeWidthByRatio()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim TargetShape As Shape
    Dim RatioStr As String
    Dim Ratio As Double
    Dim KeyWidth As Double
    Dim OriginalPosition() As Double, NewPosition() As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    MaxNum = SelectedShapes.Count

AskRatio:
    RatioStr = InputBox("比（0以外の数値）を入力", Default:="1")
    If IsNumeric(RatioStr) Then
        Ratio = CDbl(RatioStr)
        If 0 < Ratio Then
            GoTo Main
        ElseIf Ratio < 0 Then
            Ratio = 1 / -Ratio
            GoTo Main
        End If
    ElseIf StrPtr(RatioStr) = 0 Then
        Exit Sub
    End If
    GoTo AskRatio

Main:
    KeyWidth = GetVisualPoints(SelectedShapes(MaxNum))("Width")
    If MaxNum = 1 Then
        MaxNum = 2
    End If

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Width = KeyWidth * Ratio
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub

Public Sub ResizeHeightByRatio()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim TargetShape As Shape
    Dim RatioStr As String
    Dim Ratio As Double
    Dim KeyHeight As Double
    Dim OriginalPosition() As Double, NewPosition() As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    MaxNum = SelectedShapes.Count

AskRatio:
    RatioStr = InputBox("比（0以外の数値）を入力", Default:="1")
    If IsNumeric(RatioStr) Then
        Ratio = CDbl(RatioStr)
        If 0 < Ratio Then
            GoTo Main
        ElseIf Ratio < 0 Then
            Ratio = 1 / -Ratio
            GoTo Main
        End If
    ElseIf StrPtr(RatioStr) = 0 Then
        Exit Sub
    End If
    GoTo AskRatio

Main:
    KeyHeight = GetVisualPoints(SelectedShapes(MaxNum))("Height")
    If MaxNum = 1 Then
        MaxNum = 2
    End If

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Height = KeyHeight * Ratio
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub


'Match
Public Sub MatchHeightToWidth()
    Dim SelectedShapes As ShapeRange
    Dim TargetShape As Shape
    Dim i As Long
    Dim OriginalPosition() As Double, NewPosition() As Double

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = msoFalse
    For i = 1 To SelectedShapes.Count
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Height = TargetShape.Width
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementTop OriginalPosition(2) - NewPosition(2)
    Next i
End Sub

Public Sub MatchWidthToHeight()
    Dim SelectedShapes As ShapeRange
    Dim TargetShape As Shape
    Dim i As Long
    Dim OriginalPosition() As Double, NewPosition() As Double

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = msoFalse
    For i = 1 To SelectedShapes.Count
        Set TargetShape = SelectedShapes(i)
        OriginalPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.Width = TargetShape.Height
        NewPosition = GetResizeAnchorPoint(TargetShape)
        TargetShape.IncrementLeft OriginalPosition(1) - NewPosition(1)
    Next i
End Sub


'FitToSlide
Public Sub FitToWidth()
    Dim SelectedShapes As ShapeRange
    Dim TargetShape As Shape
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    For i = 1 To SelectedShapes.Count
        Set TargetShape = SelectedShapes(i)
        TargetShape.Width = ActivePresentation.PageSetup.SlideWidth
        TargetShape.Left = 0
    Next i
End Sub

Public Sub FitToHeight()
    Dim SelectedShapes As ShapeRange
    Dim TargetShape As Shape
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = aspectLock
    For i = 1 To SelectedShapes.Count
        Set TargetShape = SelectedShapes(i)
        TargetShape.Height = ActivePresentation.PageSetup.SlideHeight
        TargetShape.Top = 0
    Next i
End Sub

Public Sub FitToFill()
    Dim SelectedShapes As ShapeRange
    Dim TargetShape As Shape
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    SelectedShapes.LockAspectRatio = msoFalse
    For i = 1 To SelectedShapes.Count
        Set TargetShape = SelectedShapes(i)
        TargetShape.Width = ActivePresentation.PageSetup.SlideWidth
        TargetShape.Height = ActivePresentation.PageSetup.SlideHeight
        TargetShape.Left = 0
        TargetShape.Top = 0
    Next i
End Sub


'Function
Private Function GetVisualPoints(ByRef ShapeA As Shape) As Collection
    Dim shapeB As Shape
    Dim node As ShapeNode
    Dim i As Long
    Dim nodePoints() As Double
    Dim VisualPoints As Collection

    If ShapeA.Rotation = 0 Or ShapeA.Rotation = 180 Then
        Set VisualPoints = New Collection
        VisualPoints.Add ShapeA.Left, "Left"
        VisualPoints.Add ShapeA.Top, "Top"
        VisualPoints.Add ShapeA.Width, "Width"
        VisualPoints.Add ShapeA.Height, "Height"
    ElseIf ShapeA.Type = msoAutoShape Or ShapeA.Type = msoFreeform Then
        Set shapeB = ShapeA.Duplicate(1)
        shapeB.Left = ShapeA.Left
        shapeB.Top = ShapeA.Top
        If ShapeA.Type = msoAutoShape Then
            shapeB.Nodes.Insert 1, msoSegmentLine, msoEditingAuto, 0, 0
            shapeB.Nodes.Delete 2
        End If

        ReDim nodePoints(1 To shapeB.Nodes.Count, 0 To 1)
        For i = 1 To shapeB.Nodes.Count
            Set node = shapeB.Nodes(i)
            nodePoints(i, 0) = node.Points(1, 1)
            nodePoints(i, 1) = node.Points(1, 2)
        Next i

        shapeB.Rotation = 0
        For i = 1 To shapeB.Nodes.Count
            shapeB.Nodes.SetPosition i, nodePoints(i, 0), nodePoints(i, 1)
        Next i

        Set VisualPoints = New Collection
        VisualPoints.Add shapeB.Left, "Left"
        VisualPoints.Add shapeB.Top, "Top"
        VisualPoints.Add shapeB.Width, "Width"
        VisualPoints.Add shapeB.Height, "Height"
        shapeB.Delete
    Else
        Set VisualPoints = GetVisualPointsByRotation(ShapeA)
    End If
    Set GetVisualPoints = VisualPoints
End Function

Private Function GetVisualPointsByRotation(ByRef ShapeA As Shape) As Collection
    Dim degRotation As Double, radRotation As Double
    Dim VisualWidth As Double, VisualHeight As Double
    Dim VisualLeft As Double, VisualTop As Double
    Dim PI As Double
    Dim VisualPoints As Collection

    PI = 4 * Atn(1)
    degRotation = ShapeA.Rotation
    radRotation = degRotation * PI / 180

    If degRotation = 0 Or degRotation = 180 Then
        VisualWidth = ShapeA.Width
        VisualHeight = ShapeA.Height
    ElseIf degRotation = 90 Or degRotation = 270 Then
        VisualWidth = ShapeA.Height
        VisualHeight = ShapeA.Width
    ElseIf degRotation > 0 And degRotation < 90 Then
        VisualWidth = ShapeA.Height * Sin(radRotation) + ShapeA.Width * Cos(radRotation)
        VisualHeight = ShapeA.Height * Cos(radRotation) + ShapeA.Width * Sin(radRotation)
    ElseIf degRotation > 90 And degRotation < 180 Then
        VisualWidth = ShapeA.Height * Sin(radRotation) - ShapeA.Width * Cos(radRotation)
        VisualHeight = -ShapeA.Height * Cos(radRotation) + ShapeA.Width * Sin(radRotation)
    ElseIf degRotation > 180 And degRotation < 270 Then
        VisualWidth = -ShapeA.Height * Sin(radRotation) - ShapeA.Width * Cos(radRotation)
        VisualHeight = -ShapeA.Height * Cos(radRotation) - ShapeA.Width * Sin(radRotation)
    ElseIf degRotation > 270 And degRotation < 360 Then
        VisualWidth = -ShapeA.Height * Sin(radRotation) + ShapeA.Width * Cos(radRotation)
        VisualHeight = ShapeA.Height * Cos(radRotation) - ShapeA.Width * Sin(radRotation)
    End If
    VisualLeft = ShapeA.Left + ShapeA.Width / 2 - VisualWidth / 2
    VisualTop = ShapeA.Top + ShapeA.Height / 2 - VisualHeight / 2

    Set VisualPoints = New Collection
    VisualPoints.Add VisualLeft, "Left"
    VisualPoints.Add VisualTop, "Top"
    VisualPoints.Add VisualWidth, "Width"
    VisualPoints.Add VisualHeight, "Height"
    Set GetVisualPointsByRotation = VisualPoints
End Function

Function GetResizeAnchorPoint(ByRef ShapeA As Shape)
    Dim ShapePoints As Collection
    Dim XY(2) As Double

    Set ShapePoints = GetVisualPoints(ShapeA)
    Select Case ResizeAnchor
        Case 1
            XY(1) = ShapePoints("Left")
            XY(2) = ShapePoints("Top")
        Case 2
            XY(1) = ShapePoints("Left") + ShapePoints("Width") / 2
            XY(2) = ShapePoints("Top")
        Case 3
            XY(1) = ShapePoints("Left") + ShapePoints("Width")
            XY(2) = ShapePoints("Top") + ShapePoints("Height")
        Case 4
            XY(1) = ShapePoints("Left")
            XY(2) = ShapePoints("Top") + ShapePoints("Height") / 2
        Case 5
            XY(1) = ShapePoints("Left") + ShapePoints("Width") / 2
            XY(2) = ShapePoints("Top") + ShapePoints("Height") / 2
        Case 6
            XY(1) = ShapePoints("Left") + ShapePoints("Width")
            XY(2) = ShapePoints("Top") + ShapePoints("Height") / 2
        Case 7
            XY(1) = ShapePoints("Left")
            XY(2) = ShapePoints("Top") + ShapePoints("Height")
        Case 8
            XY(1) = ShapePoints("Left") + ShapePoints("Width") / 2
            XY(2) = ShapePoints("Top") + ShapePoints("Height")
        Case 9
            XY(1) = ShapePoints("Left") + ShapePoints("Width")
            XY(2) = ShapePoints("Top") + ShapePoints("Height")
    End Select
    GetResizeAnchorPoint = XY
End Function
