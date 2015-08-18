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




oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "subx/2-5.8mm", "YPosition:=", "-1mm" , "ZPosition:=", "SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=",  _
  "6mm", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "fee1", "Flags:=",  _
  "", "Color:=", "(1 256 1)", "Transparency:=", 0.94, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", true)



  
