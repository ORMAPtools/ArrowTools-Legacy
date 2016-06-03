Imports System.Windows.Forms

Public Class settingsMenu
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

    Public Sub New()
        Me.Add("Arrow type assignments")
        Me.Add("Arrow definitions")
    End Sub

    Protected Overrides Sub OnSelChange(cookie As Integer)
        MyBase.OnSelChange(cookie)
        Try
            If Me.Value = "Arrow type assignments" Then
                Dim theWindow As ESRI.ArcGIS.Framework.IDockableWindow
                Dim theUID As New ESRI.ArcGIS.esriSystem.UID
                theUID.Value = "ORMAP_ArrowTools_SettingsForm"
                theWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(theUID)
                theWindow.Show(True)
            End If
        Catch ex As Exception
            MessageBox.Show("onSelChange - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        MyBase.OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
End Class
