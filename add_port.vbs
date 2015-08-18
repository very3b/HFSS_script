'--------------VBS to generate MSline on FPC board-------------------------------------
'Created by JJWW ....last modified date 59:21 30/12/2014.
'--------------VBS to generate MSline on FPC board-------------------------------------
'---------------12:56PM 31/12/2014------------------------------
'Base on added-on script, the function is used to add 
' 1, Geometric parameters definition of tissue
' 2, Material properties into material library
' 3, Geometry box draw into mother drawing
'-----------------------------------------------------------------------



Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()
'--------------Define parameters-------------------------


'------------------------------------------------------------
'----------------------Fat (fatH)------1mm----------------------
'----------------------Blood (bloodH)----3mm----------------------
'----------------------Fat (fatH)------1mm----------------------
'------------------------------------------------------------


'--------------define materials----------------------------------------------------------------------------------------
'eps_fat =3.82			 rho=1.65


'--------------define materials----------------------------------------------------------------------------------------
 
''''''''''''''''''''''''''''''''''Draw tissue''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Set oEditor = oDesign.SetActiveEditor("3D Modeler")




oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "subx/2-4.6mm", "YStart:=", "subY-1mm-patchY-6.6mm", "ZStart:=", "5mm+CuH+SubH", "Width:=", "CuH+SubH-FlexH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "Rectangle1", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)



  
