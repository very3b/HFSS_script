


'--------------VBS to generate Patch Antenna board-------------------------------------
'1, Airbox modified to support tissue increase from 1mm to 10mm
'	tissue thickness is around 5 mm
'2, "SolveInside:=", changed to false to reducing simulation time
'3, Modify to Patch Antenna
'--------------VBS to generate Patch Antenna on FPC board-------------------------------------
'--30/12/2014
'Create simulation template for MSline 
'-1,Parameters Definition
'0, Material adding
'1,Patch
'2,Substrate
'3,Port2 setting up
'4,Airbox
'5,Solution/Sweep
'--------------VBS to generate Patch Antenna on FPC board-------------------------------------

Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
oAnsoftApp.SetNumberOfProcessors(20)
Set oDesktop = oAnsoftApp.GetAppDesktop()

oDesktop.RestoreWindow
oDesktop.NewProject
Set oProject = oDesktop.GetActiveProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")

'--------------Define paramters
dim circle_or_rect, solution_freq, patchX, patchY, subH, subX, subY, feedX, feedY, coax_inner_rad, coax_outer_rad, feed_length, units, israd,patch_edge
dim airH,airX,airY,CuH,tissueH,fatH,bloodH

bloodH=3
fatH=1

tissueH=fatH*2 + bloodH
'-----------------------------------------

'circle_or_rect = CInt( args(0))  '1 is rectangular, 0 is circular

dim light_speed
light_speed = 299792458


solution_freq = 24

  dim freq_hz
  freq_hz=solution_freq*1e9

  dim WL 
  WL = light_speed/(freq_hz)
''''''''''''''
'Product Code	Dielectric thickness(Mil)	Copper
'AP 7156E** 2.0 9 
'AP 7125E** 2.0 12 
'AP 8535R   3.0 18 
'AP 8545R   4.0 18 
CuH =0.018
subH = 0.0254*2 'mm 0.0254 is mil to millimeter conversion

subX = 15 'mm
subY = 28 'mm

 patchX = 12 'mm
 patchY = 10 'mm

feedX =  0.0254*3.5 'mm 
feedY = 15 'mm
feed_length = 10 'mm
patch_edge = 5 'mm
units = "mm"
airH = 10 
airX = 30
airY = 60

   

if patchX > subX then
  subX = patchX*1.1
end if
if patchY > subY then
  subY = patchY*1.1
end if


oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", _
Array("NAME:--Patch Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:patchX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchX & units), _
Array("NAME:patchY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patchY & units), _
Array("NAME:patch_edge", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", patch_edge & units), _
Array("NAME:--Substrate Dimensions", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:subH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subH & units ), _
Array("NAME:subX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subX & units), _
Array("NAME:subY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", subY & units), _
Array("NAME:--Copper", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:CuH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", CuH & units), _
Array("NAME:--AirBox", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:airX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", airX & units), _
Array("NAME:airY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", airY & units), _
Array("NAME:airH", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", airH & units),_
Array("NAME:--Feed", "PropType:=", "SeparatorProp", "UserDef:=", true, "Value:=", ""), _
Array("NAME:feedX", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedX & units), _
Array("NAME:feedY", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feedY & units), _
Array("NAME:feed_length", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", feed_length & units))))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create substrate material

Material_Name ="AP Laminate"
Material_Name = Material_Name & "FPC"
Permittivity = 3.4
TandD = 0.004

Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:" & Material_Name, "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=", Permittivity , "dielectric_loss_tangent:=",TandD)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Substrate and bottom metalization
'''
''''''''''''''''''''''MET1'''''''''''''''''''e=subH
''''''''''''''''''''''SUB'''''''''''''''''''e=0
''''''''''''''''''''''METG'''''''''''''''''e=0-CuH
''''''''''''''''''''''ABS 5mm''''''''''''''e=0-CuH-5mm




Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), _
  Array("NAME:Attributes", "Name:=", "sub", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)

  
  oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), _
  Array("NAME:Attributes", "Name:=", "METG", "Flags:=",  _
  "", "Color:=", "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)

'----------------------------------------FACE B-----------------------------------------------
  oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "0mm+SubH+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), _
  Array("NAME:Attributes", "Name:=", "subb", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)

  
  oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "0mm+SubH*2+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), _
  Array("NAME:Attributes", "Name:=", "METGb", "Flags:=",  _
  "", "Color:=", "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)
'----------------------------------------FACE B-----------------------------------------------


  
'----------------------------------------FACE Vertical-----------------------------------------------
 oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "-cuH", "XSize:=", "subX", "YSize:=",  _
  "-CuH", "ZSize:=", "subH*2+cuH*3+5mm"), _
  Array("NAME:Attributes", "Name:=", "METV", "Flags:=",  _
  "", "Color:=", "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)
  
  oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-subX/2", "YPosition:=", "-1mm" , "ZPosition:=", "subH", "XSize:=", "subX", "YSize:=",  _
  "subH", "ZSize:=", "5mm+cuH*2"), _
  Array("NAME:Attributes", "Name:=", "subv", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", Material_Name, "SolveInside:=", true)
  
  
 '----------------------------------------FACE Vertical-----------------------------------------------
 
  
  
  
'  Set oModule = oDesign.GetModule("BoundarySetup")  

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws Patch Feed

Set oEditor = oDesign.SetActiveEditor("3D Modeler")


 oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-feedX/2", "YPosition:=", "0" , "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=",  _
  "feedY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)
 
 oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-patchX/2", "YPosition:=", "feedY" , "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), _
  Array("NAME:Attributes", "Name:=", "PATCH", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)

'----------------------------------------FACE B-----------------------------------------------
 oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-feedX/2", "YPosition:=", "0" , "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=",  _
  "feedY", "ZSize:=", "CuH"), _
  Array("NAME:Attributes", "Name:=", "MET1b", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)
 
 oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-patchX/2", "YPosition:=", "feedY" , "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), _
  Array("NAME:Attributes", "Name:=", "PATCHb", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "copper", "SolveInside:=", false)
'----------------------------------------FACE B-----------------------------------------------

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draws MSline




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Draw port 1
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-feedX/2", "YStart:=", "0", "ZStart:=",  _
  "0mm", "Width:=", "subH", "Height:=", "feedX", "WhichAxis:=", "Y"),_
  Array("NAME:Attributes", "Name:=", "port1", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
''Draw Port 2'''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
'' Port does not use variable from VBScript, should define in variable manager
' Should be very careful
oEditor.CreateRectangle Array("NAME:RectangleParameters", "CoordinateSystemID:=",  _
  -1, "IsCovered:=", true, "XStart:=", "-feedX/2", "YStart:=", "0", "ZStart:=",  _
  "0mm+SubH*2+CuH+5mm+CuH", "Width:=", "-subH", "Height:=", "feedX", "WhichAxis:=", "Y"),_
  Array("NAME:Attributes", "Name:=", "port2", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=",  _
  0.3, "PartCoordinateSystem:=", "Global", "MaterialName:=",  _
  "pec", "SolveInside:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")  

oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "0mm", "0mm"), "End:=", Array( _
  "0mm", "0mm", subH & units)), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")

  dim port2e1, port2e2
  port2e1=SubH+CuH*2+tissueH
  port2e2=SubH*2+ CuH*2+tissueH


oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "0mm", port2e1 & uints), "End:=", Array( _
  "0mm", "0mm", port2e2 & units)), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AirBox Setup
  
  
  Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "CoordinateSystemID:=", -1, "XPosition:=",  _
  "-airX/2", "YPosition:=", "-airY/4" , "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=",  _
  "airY", "ZSize:=", "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=",  _
  "", "Color:=", "(132 132 256)", "Transparency:=", 0.9, "PartCoordinateSystem:=",  _
  "Global", "MaterialName:=", "air", "SolveInside:=", true)

set oModule=oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedHField:=", false, "IsEnforcedEField:=", false, "IsFssReference:=",  _
  false, "IsForPML:=", false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=",  _
  true)
  
  
  
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Solution Setup 
 


 



Set oModule = oDesign.GetModule("AnalysisSetup")

oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", solution_freq&"GHZ", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "BasisOrder:=", 1, "UseIterativeSolver:=",  _
  true, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseConvOutputVariable:=", false, "IsEnabled:=",  _
  true, "ExternalMesh:=", false, "UseMaxTetIncrease:=", false, "MaxTetIncrease:=",  _
  100000, "PortAccuracy:=", 2, "UseABCOnPort:=", false, "SetPortMinMaxTri:=",  _
  false)
  
dim start_freq
dim stop_freq
start_freq = 0.8*solution_freq
stop_freq = 1.2*solution_freq

 oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", start_freq&"GHz", "StopValue:=", stop_freq&"GHz", "Count:=", 100, "Type:=",  _
  "Interpolating", "SaveFields:=", false, "InterpTolerance:=", 0.5, "InterpMaxSolns:=",  _
  50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=", false, "InterpUseS:=",  _
  true, "InterpUseT:=", false, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseFullBasis:=", true)   




 
