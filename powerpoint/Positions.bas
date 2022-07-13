Option Explicit

Private positionsRibbon As IRibbonUI
Private DistributeByCenter As Boolean
Private SwapReference As Integer

Public Sub positionsRibbonOnLoad(ribbon As IRibbonUI)
    SwapReference = 5
    DistributeByCenter = False
    Set positionsRibbon = ribbon
    positionsRibbon.Invalidate
End Sub

Public Sub DistributeByCenter_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = DistributeByCenter
End Sub

Public Sub DistributeByCenterChange(control As IRibbonControl, pressed As Boolean)
    DistributeByCenter = pressed
End Sub

Public Sub SwapRef_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (SwapReference = CInt(Right(control.Id, 1)))
End Sub

Public Sub SwapRefChange(control As IRibbonControl, pressed As Boolean)
    SwapReference = CInt(Right(control.Id, 1))
    positionsRibbon.Invalidate
End Sub


'Align
Public Sub AlignLeft()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignLefts, msoTrue
        Exit Sub
    End If

    KeyPoint = GetVisualPoints(SelectedShapes(MaxNum))("Left")

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        TargetShape.IncrementLeft KeyPoint - GetVisualPoints(TargetShape)("Left")
    Next i
End Sub

Public Sub AlignVertical()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignCenters, msoTrue
        Exit Sub
    End If

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Left") + KeyPoints("Width") / 2

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementLeft KeyPoint - TargetPoints("Left") - TargetPoints("Width") / 2
    Next i
End Sub

Public Sub AlignRight()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignRights, msoTrue
        Exit Sub
    End If

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Left") + KeyPoints("Width")

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementLeft KeyPoint - TargetPoints("Left") - TargetPoints("Width")
    Next i
End Sub

Public Sub AlignTop()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignTops, msoTrue
        Exit Sub
    End If

    KeyPoint = GetVisualPoints(SelectedShapes(MaxNum))("Top")

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        TargetShape.IncrementTop KeyPoint - GetVisualPoints(TargetShape)("Top")
    Next i
End Sub

Public Sub AlignHorizontal()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignMiddles, msoTrue
        Exit Sub
    End If

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Top") + KeyPoints("Height") / 2

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementTop KeyPoint - TargetPoints("Top") - TargetPoints("Height") / 2
    Next i
End Sub

Public Sub AlignBottom()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPoint As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignBottoms, msoTrue
        Exit Sub
    End If

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPoint = KeyPoints("Top") + KeyPoints("Height")

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementTop KeyPoint - TargetPoints("Top") - TargetPoints("Height")
    Next i
End Sub

Public Sub AlignCenter()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPointX As Double, KeyPointY As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes.Align msoAlignCenters, msoTrue
        SelectedShapes.Align msoAlignMiddles, msoTrue
        Exit Sub
    End If

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPointX = KeyPoints("Left") + KeyPoints("Width") / 2
    KeyPointY = KeyPoints("Top") + KeyPoints("Height") / 2

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementLeft KeyPointX - TargetPoints("Left") - TargetPoints("Width") / 2
        TargetShape.IncrementTop KeyPointY - TargetPoints("Top") - TargetPoints("Height") / 2
    Next i
End Sub

Public Sub AlignRadial()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim PivotShape As Shape
    Dim PivotPoints As Collection
    Dim TargetShape As Shape
    Dim TargetPointX As Single, TargetPointY As Single
    Dim Distance As Double, currentDistance As Double, proportion As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 3 Then
        Exit Sub
    End If

    Set PivotShape = SelectedShapes(1)
    Set PivotPoints = GetVisualPoints(PivotShape)
    Distance = DistanceBetweenTwoShapes(PivotShape, SelectedShapes(MaxNum))

    For i = 2 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        currentDistance = DistanceBetweenTwoShapes(PivotShape, TargetShape)
        proportion = (currentDistance - Distance) / currentDistance
        TargetPointX = TargetShape.Left + TargetShape.Width / 2
        TargetPointY = TargetShape.Top + TargetShape.Height / 2

        TargetShape.IncrementLeft (PivotPoints("Left") + PivotPoints("Width") / 2 - TargetPointX) * proportion
        TargetShape.IncrementTop (PivotPoints("Top") + PivotPoints("Height") / 2 - TargetPointY) * proportion
    Next i
End Sub


'Adjoin
Public Sub AdjoinHorizontal()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyShape As Shape
    Dim KeyPoints As Collection
    Dim KeyPointL As Double, KeyPointR As Double
    Dim SortedShapes() As Shape
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long, KeyShapeIndex As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    Set KeyShape = SelectedShapes(MaxNum)
    Set KeyPoints = GetVisualPoints(KeyShape)
    KeyPointL = KeyPoints("Left")
    KeyPointR = KeyPoints("Left") + KeyPoints("Width")

    ReDim SortedShapes(MaxNum)
    For i = 1 To MaxNum
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    SortedShapes = SortShapesByLeft(SortedShapes)

    KeyShapeIndex = MaxNum
    For i = 1 To MaxNum
        If KeyShape Is SortedShapes(i) Then
            KeyShapeIndex = i
        ElseIf KeyShapeIndex < i Then
            Set TargetShape = SortedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementLeft KeyPointR - TargetPoints("Left")
            KeyPointR = KeyPointR + TargetPoints("Width")
        End If
    Next i

    For i = KeyShapeIndex - 1 To 1 Step -1
        Set TargetShape = SortedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementLeft KeyPointL - TargetPoints("Left") - TargetPoints("Width")
        KeyPointL = KeyPointL - TargetPoints("Width")
    Next i
End Sub

Public Sub AdjoinAlignHorizontal()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyShape As Shape
    Dim KeyPoints As Collection
    Dim KeyPointH As Double, KeyPointL As Double, KeyPointR As Double
    Dim SortedShapes() As Shape
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long, KeyShapeIndex As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    Set KeyShape = SelectedShapes(MaxNum)
    Set KeyPoints = GetVisualPoints(KeyShape)
    KeyPointL = KeyPoints("Left")
    KeyPointR = KeyPoints("Left") + KeyPoints("Width")
    KeyPointH = KeyPoints("Top") + KeyPoints("Height") / 2

    ReDim SortedShapes(MaxNum)
    For i = 1 To MaxNum
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    SortedShapes = SortShapesByLeft(SortedShapes)

    KeyShapeIndex = MaxNum
    For i = 1 To MaxNum
        If KeyShape Is SortedShapes(i) Then
            KeyShapeIndex = i
        ElseIf KeyShapeIndex < i Then
            Set TargetShape = SortedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementLeft KeyPointR - TargetPoints("Left")
            KeyPointR = KeyPointR + TargetPoints("Width")
            TargetShape.IncrementTop KeyPointH - TargetPoints("Top") - TargetPoints("Height") / 2
        End If
    Next i

    For i = KeyShapeIndex - 1 To 1 Step -1
        Set TargetShape = SortedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementLeft KeyPointL - TargetPoints("Left") - TargetPoints("Width")
        KeyPointL = KeyPointL - TargetPoints("Width")
        TargetShape.IncrementTop KeyPointH - TargetPoints("Top") - TargetPoints("Height") / 2
    Next i
End Sub

Public Sub AdjoinVertical()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyShape As Shape
    Dim KeyPoints As Collection
    Dim KeyPointT As Double, KeyPointB As Double
    Dim SortedShapes() As Shape
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long, KeyShapeIndex As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    Set KeyShape = SelectedShapes(MaxNum)
    Set KeyPoints = GetVisualPoints(KeyShape)
    KeyPointT = KeyPoints("Top")
    KeyPointB = KeyPoints("Top") + KeyPoints("Height")

    ReDim SortedShapes(MaxNum)
    For i = 1 To MaxNum
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    SortedShapes = SortShapesByTop(SortedShapes)

    KeyShapeIndex = MaxNum
    For i = 1 To MaxNum
        If KeyShape Is SortedShapes(i) Then
            KeyShapeIndex = i
        ElseIf KeyShapeIndex < i Then
            Set TargetShape = SortedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementTop KeyPointB - TargetPoints("Top")
            KeyPointB = KeyPointB + TargetPoints("Height")
        End If
    Next i

    For i = KeyShapeIndex - 1 To 1 Step -1
        Set TargetShape = SortedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementTop KeyPointT - TargetPoints("Top") - TargetPoints("Height")
        KeyPointT = KeyPointT - TargetPoints("Height")
    Next i
End Sub

Public Sub AdjoinAlignVertical()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyShape As Shape
    Dim KeyPoints As Collection
    Dim KeyPointV As Double, KeyPointT As Double, KeyPointB As Double
    Dim SortedShapes() As Shape
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim i As Long, KeyShapeIndex As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        Exit Sub
    End If

    Set KeyShape = SelectedShapes(MaxNum)
    Set KeyPoints = GetVisualPoints(KeyShape)
    KeyPointT = KeyPoints("Top")
    KeyPointB = KeyPoints("Top") + KeyPoints("Height")
    KeyPointV = KeyPoints("Left") + KeyPoints("Width") / 2

    ReDim SortedShapes(MaxNum)
    For i = 1 To MaxNum
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    SortedShapes = SortShapesByTop(SortedShapes)

    KeyShapeIndex = MaxNum
    For i = 1 To MaxNum
        If KeyShape Is SortedShapes(i) Then
            KeyShapeIndex = i
        ElseIf KeyShapeIndex < i Then
            Set TargetShape = SortedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementTop KeyPointB - TargetPoints("Top")
            KeyPointB = KeyPointB + TargetPoints("Height")
            TargetShape.IncrementLeft KeyPointV - TargetPoints("Left") - TargetPoints("Width") / 2
        End If
    Next i

    For i = KeyShapeIndex - 1 To 1 Step -1
        Set TargetShape = SortedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetShape.IncrementTop KeyPointT - TargetPoints("Top") - TargetPoints("Height")
        KeyPointT = KeyPointT - TargetPoints("Height")
        TargetShape.IncrementLeft KeyPointV - TargetPoints("Left") - TargetPoints("Width") / 2
    Next i
End Sub


'Distribute
Public Sub DistributeHorizontal()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPointsL As Collection, KeyPointsR As Collection
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim SpaceWidth As Double, TempKeyPoint As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 3 Then
        Exit Sub
    End If

    If SelectedShapes(1).Left < SelectedShapes(2).Left Then
        Set KeyPointsL = GetVisualPoints(SelectedShapes(1))
        Set KeyPointsR = GetVisualPoints(SelectedShapes(MaxNum))
    Else
        Set KeyPointsL = GetVisualPoints(SelectedShapes(MaxNum))
        Set KeyPointsR = GetVisualPoints(SelectedShapes(1))
    End If

    If DistributeByCenter Then
        TempKeyPoint = KeyPointsL("Left") + KeyPointsL("Width") / 2
        SpaceWidth = (KeyPointsR("Left") + KeyPointsR("Width") / 2 - TempKeyPoint) / (MaxNum - 1)
        For i = 2 To MaxNum - 1
            Set TargetShape = SelectedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TempKeyPoint = TempKeyPoint + SpaceWidth
            TargetShape.IncrementLeft TempKeyPoint - TargetPoints("Left") - TargetPoints("Width") / 2
        Next i
    Else
        SpaceWidth = KeyPointsR("Left") - (KeyPointsL("Left") + KeyPointsL("Width"))
        For i = 2 To MaxNum - 1
            SpaceWidth = SpaceWidth - GetVisualPoints(SelectedShapes(i))("Width")
        Next i
        SpaceWidth = SpaceWidth / (MaxNum - 1)

        TempKeyPoint = KeyPointsL("Left") + KeyPointsL("Width") + SpaceWidth
        For i = 2 To MaxNum - 1
            Set TargetShape = SelectedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementLeft TempKeyPoint - TargetPoints("Left")
            TempKeyPoint = TempKeyPoint + TargetPoints("Width") + SpaceWidth
        Next i
    End If
End Sub

Public Sub DistributeVertical()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPointsT As Collection, KeyPointsB As Collection
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim SpaceHeight As Double, TempKeyPoint As Double
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 3 Then
        Exit Sub
    End If

    If SelectedShapes(1).Top < SelectedShapes(2).Top Then
        Set KeyPointsT = GetVisualPoints(SelectedShapes(1))
        Set KeyPointsB = GetVisualPoints(SelectedShapes(MaxNum))
    Else
        Set KeyPointsT = GetVisualPoints(SelectedShapes(MaxNum))
        Set KeyPointsB = GetVisualPoints(SelectedShapes(1))
    End If

    If DistributeByCenter Then
        TempKeyPoint = KeyPointsT("Top") + KeyPointsT("Height") / 2
        SpaceHeight = (KeyPointsB("Top") + KeyPointsB("Height") / 2 - TempKeyPoint) / (MaxNum - 1)
        For i = 2 To MaxNum - 1
            Set TargetShape = SelectedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TempKeyPoint = TempKeyPoint + SpaceHeight
            TargetShape.IncrementTop TempKeyPoint - TargetPoints("Top") - TargetPoints("Height") / 2
        Next i
    Else
        SpaceHeight = KeyPointsB("Top") - (KeyPointsT("Top") + KeyPointsT("Height"))
        For i = 2 To MaxNum - 1
            SpaceHeight = SpaceHeight - GetVisualPoints(SelectedShapes(i))("Height")
        Next i
        SpaceHeight = SpaceHeight / (MaxNum - 1)

        TempKeyPoint = KeyPointsT("Top") + KeyPointsT("Height") + SpaceHeight
        For i = 2 To MaxNum - 1
            Set TargetShape = SelectedShapes(i)
            Set TargetPoints = GetVisualPoints(TargetShape)
            TargetShape.IncrementTop TempKeyPoint - TargetPoints("Top")
            TempKeyPoint = TempKeyPoint + TargetPoints("Height") + SpaceHeight
        Next i
    End If
End Sub

Public Sub DistributeCenter()
    Call DistributeHorizontal
    Call DistributeVertical
End Sub

'Swap
Public Sub Swap()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim TargetShape As Shape, NextShape As Shape
    Dim FirstPosition() As Single
    Dim CurrentPosition() As Single
    Dim NextPosition() As Single
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

    FirstPosition = GetSwapReferencePoint(SelectedShapes(1))
    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        CurrentPosition = GetSwapReferencePoint(TargetShape)
        Set NextShape = SelectedShapes(i + 1)
        NextPosition = GetSwapReferencePoint(NextShape)

        TargetShape.IncrementLeft NextPosition(1) - CurrentPosition(1)
        TargetShape.IncrementTop NextPosition(2) - CurrentPosition(2)
        Call SwapZOrder(TargetShape, NextShape)
    Next i

    Set TargetShape = SelectedShapes(SelectedShapes.Count)
    CurrentPosition = GetSwapReferencePoint(TargetShape)

    TargetShape.IncrementLeft FirstPosition(1) - CurrentPosition(1)
    TargetShape.IncrementTop FirstPosition(2) - CurrentPosition(2)
End Sub

'Rotation
Public Sub ResetRotation()
    Dim SelectedShapes As ShapeRange
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    For i = 1 To SelectedShapes.Count
        SelectedShapes(i).Rotation = 0
        If SelectedShapes(i).Connector Then
            If SelectedShapes(i).Height > SelectedShapes(i).Width Then
                SelectedShapes(i).Width = 0
            Else
                SelectedShapes(i).Height = 0
            End If
        End If
    Next i
End Sub

Public Sub CopyRotation()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    MaxNum = SelectedShapes.Count
    If MaxNum < 2 Then
        SelectedShapes(1).Rotation = 0
        Exit Sub
    End If

    For i = 1 To SelectedShapes.Count - 1
        SelectedShapes(i).Rotation = SelectedShapes(MaxNum).Rotation
    Next i
End Sub

Public Sub ConcentrateOrient()
    Dim SelectedShapes As ShapeRange
    Dim MaxNum As Long
    Dim KeyPoints As Collection
    Dim KeyPointX As Double, KeyPointY As Double
    Dim TargetShape As Shape
    Dim TargetPoints As Collection
    Dim TargetPointX As Double, TargetPointY As Double
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

    Set KeyPoints = GetVisualPoints(SelectedShapes(MaxNum))
    KeyPointX = KeyPoints("Left") + KeyPoints("Width") / 2
    KeyPointY = KeyPoints("Top") + KeyPoints("Height") / 2

    For i = 1 To MaxNum - 1
        Set TargetShape = SelectedShapes(i)
        Set TargetPoints = GetVisualPoints(TargetShape)
        TargetPointX = TargetPoints("Left") + TargetPoints("Width") / 2
        TargetPointY = TargetPoints("Top") + TargetPoints("Height") / 2
        TargetShape.Rotation = AngleBetweenTwoPoints(KeyPointX, KeyPointY, TargetPointX, TargetPointY)
    Next i
End Sub

'ZOrder
Public Sub MoveFront()
    Dim SelectedShapes As ShapeRange
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    For i = SelectedShapes.Count To 1 Step -1
        SelectedShapes(i).ZOrder msoBringToFront
    Next i
End Sub

Public Sub MoveBack()
    Dim SelectedShapes As ShapeRange
    Dim i As Long

    If ActiveWindow.Selection.Type = ppSelectionNone _
    Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Exit Sub
    End If

    Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    For i = SelectedShapes.Count To 1 Step -1
        SelectedShapes(i).ZOrder msoSendToBack
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

Private Function DistanceBetweenTwoShapes(ByRef ShapeA As Shape, _
                                  ByRef shapeB As Shape) As Double
    Dim shapeAX As Single, shapeAY As Single
    Dim shapeBX As Single, shapeBY As Single

    shapeAX = ShapeA.Left + ShapeA.Width / 2
    shapeAY = ShapeA.Top + ShapeA.Height / 2
    shapeBX = shapeB.Left + shapeB.Width / 2
    shapeBY = shapeB.Top + shapeB.Height / 2

    DistanceBetweenTwoShapes = Sqr((shapeAX - shapeBX) ^ 2 + (shapeAY - shapeBY) ^ 2)
End Function

Private Function AngleBetweenTwoPoints(ByRef KeyPointX As Double, _
                               ByRef KeyPointY As Double, _
                               ByRef TargetPointX As Double, _
                               ByRef TargetPointY As Double) As Double
    Dim X As Double, Y As Double
    Dim angle As Double

    X = TargetPointX - KeyPointX
    Y = TargetPointY - KeyPointY
    If X = 0 And Y <= 0 Then
        angle = 180
    ElseIf X = 0 And Y > 0 Then
        angle = 0
    ElseIf X > 0 And Y = 0 Then
        angle = 270
    ElseIf X < 0 And Y = 0 Then
        angle = 90
    Else
        angle = Atn(Y / X) * 45 / Atn(1)
        If X < 0 Then
            angle = angle + 90
        Else
            angle = angle + 270
        End If
    End If
    AngleBetweenTwoPoints = angle
End Function

Private Function SortShapesByLeft(ByRef SelectedShapes() As Shape)
    Dim i As Long
    Dim j As Long
    Dim TempShape As Shape
    Dim SortedShapes() As Shape
    ReDim SortedShapes(UBound(SelectedShapes))

    For i = 1 To UBound(SortedShapes)
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    For i = 2 To UBound(SortedShapes)
        Set TempShape = SortedShapes(i)
        If SortedShapes(i - 1).Left > TempShape.Left Then
            j = i
            Do While j > 1
                If SortedShapes(j - 1).Left <= TempShape.Left Then
                    Exit Do
                End If
                Set SortedShapes(j) = SortedShapes(j - 1)
                j = j - 1
            Loop
            Set SortedShapes(j) = TempShape
        End If
    Next i
    SortShapesByLeft = SortedShapes
End Function

Private Function SortShapesByTop(ByRef SelectedShapes() As Shape)
    Dim i As Long
    Dim j As Long
    Dim TempShape As Shape
    Dim SortedShapes() As Shape
    ReDim SortedShapes(UBound(SelectedShapes))

    For i = 1 To UBound(SortedShapes)
        Set SortedShapes(i) = SelectedShapes(i)
    Next i
    For i = 2 To UBound(SortedShapes)
        Set TempShape = SortedShapes(i)
        If SortedShapes(i - 1).Top > TempShape.Top Then
            j = i
            Do While j > 1
                If SortedShapes(j - 1).Top <= TempShape.Top Then
                    Exit Do
                End If
                Set SortedShapes(j) = SortedShapes(j - 1)
                j = j - 1
            Loop
            Set SortedShapes(j) = TempShape
        End If
    Next i
    SortShapesByTop = SortedShapes
End Function

Private Function MoveZTo(ByRef ShapeA As Shape, ByVal TargetZOrder As Long)
    Dim CurrentValue As Long

    Do While ShapeA.ZOrderPosition < TargetZOrder
        CurrentValue = ShapeA.ZOrderPosition
        ShapeA.ZOrder (msoBringForward)
        If ShapeA.ZOrderPosition = CurrentValue Then
            Exit Do
        End If
    Loop

    Do While ShapeA.ZOrderPosition > TargetZOrder
        CurrentValue = ShapeA.ZOrderPosition
        ShapeA.ZOrder (msoSendBackward)
        If ShapeA.ZOrderPosition = CurrentValue Then
            Exit Do
        End If
    Loop
End Function

Private Function SwapZOrder(ByRef ShapeA As Shape, ByRef shapeB As Shape)
    Dim HigherShape As Shape, LowerShape As Shape
    Dim HigherZOrder As Long, LowerZOrder As Long

    If ShapeA.ZOrderPosition > shapeB.ZOrderPosition Then
        Set HigherShape = ShapeA
        Set LowerShape = shapeB
    Else
        Set HigherShape = shapeB
        Set LowerShape = ShapeA
    End If
    HigherZOrder = HigherShape.ZOrderPosition
    LowerZOrder = LowerShape.ZOrderPosition

    If LowerShape.Type = msoGroup Then
        HigherZOrder = HigherZOrder - LowerShape.GroupItems.Count
    End If
    If HigherShape.Type = msoGroup Then
        HigherZOrder = HigherZOrder + HigherShape.GroupItems.Count
    End If

    Call MoveZTo(LowerShape, HigherZOrder)
    Call MoveZTo(HigherShape, LowerZOrder)
End Function

Private Function GetSwapReferencePoint(ByRef ShapeA As Shape)
    Dim ShapePoints As Collection
    Dim XY(2) As Single

    Set ShapePoints = GetVisualPoints(ShapeA)
    Select Case SwapReference
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
    GetSwapReferencePoint = XY
End Function
