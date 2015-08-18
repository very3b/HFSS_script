' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 14.0.1
' 10:34:17 AM  Jan 02, 2015
' ----------------------------------------------
Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.SetActiveProject("Project22")
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "PATCH", "Tool Parts:=",  _
  "CylinderPD"), Array("NAME:SubtractParameters", "KeepOriginals:=", false)
oDesign.Undo
oEditor.Unite Array("NAME:Selections", "Selections:=", "CylinderLED,PATCHb"), Array("NAME:UniteParameters", "KeepOriginals:=",  _
  false)
oDesign.Undo
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "PATCHb", "Tool Parts:=",  _
  "CylinderLED"), Array("NAME:SubtractParameters", "KeepOriginals:=", false)
