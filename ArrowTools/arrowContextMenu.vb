Imports System.Windows.Forms

Class arrowContextMenu
    Inherits ESRI.ArcGIS.Desktop.AddIns.MultiItem

    Private screenPosition As System.Drawing.Point

    Public Sub New()

    End Sub

    Protected Overloads Overrides Sub OnClick(ByVal item As Item)
        Try
            Select Case item.Caption
                Case SHORTER
                    arrowUtilities.keyCommands(Windows.Forms.Keys.Left, True)
                Case LONGER
                    arrowUtilities.keyCommands(Windows.Forms.Keys.Right, True)
                Case FLIP
                    arrowUtilities.keyCommands(Windows.Forms.Keys.F, False)
                Case UNLOCK
                    arrowUtilities.keyCommands(Windows.Forms.Keys.U, False)
                Case SWITCH
                    arrowUtilities.keyCommands(Windows.Forms.Keys.S, False)
                Case CANCEL
                    arrowUtilities.clearAll()
                Case SCALE10
                    _arrowScale = 0.1
                Case SCALE20
                    _arrowScale = 0.2
                Case SCALE30
                    _arrowScale = 0.3
                Case SCALE40
                    _arrowScale = 0.4
                Case SCALE50
                    _arrowScale = 0.5
                Case SCALE100
                    _arrowScale = 1
                Case SCALE200
                    _arrowScale = 2
                Case SCALE400
                    _arrowScale = 4
                Case SCALE800
                    _arrowScale = 8
                Case SCALE1000
                    _arrowScale = 10
                Case SCALE2000
                    _arrowScale = 20
                Case NARROWER
                    _zigzagWidth = _zigzagWidth - 1
                    If _zigzagWidth < 1 Then _zigzagWidth = 1
                Case WIDER
                    _zigzagWidth = _zigzagWidth + 1
                Case TOPOINT
                    _zigzagPosition = _zigzagPosition - 2.5
                    If _zigzagPosition < 1 Then _zigzagPosition = 1
                Case TOEND
                    _zigzagPosition = _zigzagPosition + 2.5
                    If _zigzagPosition > 19 Then _zigzagPosition = 19
                Case CURVELESS
                    _zigzagCurve = _zigzagCurve - 1
                Case CURVEMORE
                    _zigzagCurve = _zigzagCurve + 1
                Case STYLE_STRAIGHT
                    changeArrowStyle()
                    _thisArrow.category = arrowCategories.SingleArrow
                    _thisArrow.style = arrowStyles.Straight
                    SaveSetting("OR_DOR_dimensionArrows", "default", "category", CStr(arrowCategories.SingleArrow))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "style", CStr(arrowStyles.Straight))
                Case STYLE_LEADER
                    changeArrowStyle()
                    _thisArrow.category = arrowCategories.SingleArrow
                    _thisArrow.style = arrowStyles.Leader
                    SaveSetting("OR_DOR_dimensionArrows", "default", "category", CStr(arrowCategories.SingleArrow))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "style", CStr(arrowStyles.Leader))
                Case STYLE_ZIGZAG
                    changeArrowStyle()
                    _thisArrow.category = arrowCategories.SingleArrow
                    _thisArrow.style = arrowStyles.Zigzag
                    SaveSetting("OR_DOR_dimensionArrows", "default", "category", CStr(arrowCategories.SingleArrow))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "style", CStr(arrowStyles.Zigzag))
                Case STYLE_FREEFORM
                    changeArrowStyle()
                    _thisArrow.category = arrowCategories.SingleArrow
                    _thisArrow.style = arrowStyles.Freeform
                    SaveSetting("OR_DOR_dimensionArrows", "default", "category", CStr(arrowCategories.SingleArrow))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "style", CStr(arrowStyles.Freeform))
                    ReDim _freeformPoints(0)
                    _freeformPoints(0) = _firstPoint
                Case STYLE_ROADTIC
                    changeArrowStyle()
                    _thisArrow.category = arrowCategories.SingleArrow
                    _thisArrow.style = arrowStyles.RoadTic
                    SaveSetting("OR_DOR_dimensionArrows", "default", "category", CStr(arrowCategories.SingleArrow))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "style", CStr(arrowStyles.RoadTic))
                Case SAVEDEFAULT
                    SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagWidth", CStr(_zigzagWidth))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagCurve", CStr(_zigzagCurve))
                    SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagPosition", CStr(_zigzagPosition))
                Case FINISH
                    placeFreeformArrow()
                Case SELECTION
                    _selectNewArrows = Not _selectNewArrows 'toggle whether new arrows remain selected
                    SaveSetting("OR_DOR_dimensionArrows", "default", "selectNewArrows", CStr(_selectNewArrows))
                Case HELP
                    showHelp()
            End Select

            If _pointNumber <> 1 Then
                Windows.Forms.Cursor.Position = screenPosition
                showLineFeedback(_lastPoint)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Context menu")
        End Try
    End Sub

    Private Sub changeArrowStyle()
        Dim savepoint As ESRI.ArcGIS.Geometry.IPoint
        savePoint = _firstPoint
        clearAll()
        setLineFeedback(savePoint)
        _pointNumber = 2
    End Sub

    Protected Overloads Overrides Sub OnPopup(ByVal items As ItemCollection)
        Dim map As ESRI.ArcGIS.Carto.IMap = My.ArcMap.Document.FocusMap
        Dim isNewGroup As Boolean = False
        For i As Integer = 0 To _menuitems.Count - 1
            If _menuitems(i) = SEPARATOR Then
                isNewGroup = True
            Else
                Dim item As New Item()
                item.Caption = _menuitems(i)
                item.BeginGroup = isNewGroup
                items.Add(item)
                If item.Caption = SELECTION Then
                    If _selectNewArrows Then
                        item.Checked = True
                    Else
                        item.Checked = False
                    End If
                End If
                isNewGroup = False
            End If
        Next
        screenPosition = Windows.Forms.Cursor.Position
    End Sub
End Class