Public Class settingsWindow
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim theWindow As ESRI.ArcGIS.Framework.IDockableWindow
        Dim theUID As New ESRI.ArcGIS.esriSystem.UID
        theUID.Value = "ORMAP_ArrowTools_SettingsForm"
        theWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(theUID)
        theWindow.Show(Not theWindow.IsVisible)
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
