Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Desktop.AddIns.MultiItem

Public Class RoadTic
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Public Sub New()
        Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
        editorUID.Value = "esriEditor.editor"

        _editor = DirectCast(My.ArcMap.Application.FindExtensionByCLSID(editorUID), IEditor3)
        _editEvents = CType(CType(_editor, IEditEvents_Event), Editor)
        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)
    End Sub

    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        Try
            If checkFeatureTemplate() = False Then
                clearAll()
                Exit Sub
            End If

            _selectNewArrows = CBool(GetSetting("OR_DOR_dimensionArrows", "default", "selectNewArrows", CStr(True)))
            _thisArrow.category = arrowCategories.RoadTic
            clearAll()
            _pointNumber = 1
        Catch ex As Exception
            MessageBox.Show("OnActivate - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = (_editor.EditState = esriEditState.esriStateEditing)
    End Sub

    Protected Overrides Function OnDeactivate() As Boolean
        clearAll()
        Return True
    End Function

    Protected Overrides Sub OnMouseDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        Dim snap As ISnapEnvironment = DirectCast(_editor, ISnapEnvironment)
        Dim clickPoint As New ESRI.ArcGIS.Geometry.Point
        clickPoint = CType(_mxDoc.CurrentLocation, ESRI.ArcGIS.Geometry.Point)
        snap.SnapPoint(clickPoint)

        If arg.Button = Windows.Forms.MouseButtons.Left Then
            If _pointNumber = 1 Then
                setLineFeedback(_mxDoc.CurrentLocation)
                _pointNumber = 2
            Else
                placeArrows(_mxDoc.CurrentLocation)
                _pointNumber = 1
            End If
        ElseIf arg.Button = Windows.Forms.MouseButtons.Right Then
            CreateContextMenu()
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.KeyEventArgs)
        MyBase.OnKeyDown(arg)
        keyCommands(arg.KeyCode, arg.Shift)
    End Sub

    Protected Overrides Sub OnMouseMove(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseMove(arg)
        checkSnap()
        If _pointNumber = 2 Then
            _angleIsSet = False
            showLineFeedback(_mxDoc.CurrentLocation)
        End If
    End Sub

    Private Sub CreateContextMenu()
        Try
            Dim commandBars As ICommandBars = My.ArcMap.Application.Document.CommandBars
            Dim commandBar As ICommandBar = commandBars.Create("TemporaryContextMenu",
                ESRI.ArcGIS.SystemUI.esriCmdBarType.esriCmdBarTypeShortcutMenu)

            Dim optionalIndex As System.Object = Type.Missing
            Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UID

            uid.Value = "ORMAP_ArrowTools_arrowContextMenu"
            uid.SubType = 0
            commandBar.Add(uid, optionalIndex)

            With _menuitems
                .Clear()
                If _pointNumber > 1 Then
                    .Add(CANCEL)
                    .Add(SEPARATOR)
                    .Add(SCALE10)
                    .Add(SCALE20)
                    .Add(SCALE30)
                    .Add(SCALE40)
                    .Add(SCALE50)
                    .Add(SCALE100)
                    .Add(SCALE200)
                    .Add(SCALE400)
                    .Add(SCALE800)
                    .Add(SCALE1000)
                    .Add(SCALE2000)
                    .Add(SEPARATOR)
                    .Add(SELECTION)
                    .Add(HELP)
                Else
                    .Add(SELECTION)
                    .Add(HELP)
                End If
            End With
            commandBar.Popup()
        Catch ex As Exception
            MessageBox.Show("createContextMenu - " & ex.ToString)
        End Try
    End Sub

End Class
