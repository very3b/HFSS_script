' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 14.0.1
' 10:06:41 AM  Jan 02, 2015
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
Set oProject = oDesktop.SetActiveProject("Project17")
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateCylinder Array("NAME:CylinderParameters", "XCenter:=", "-1mm", "YCenter:=",  _
  "14mm", "ZCenter:=", "0mm", "Radius:=", "5.8309518948453mm", "Height:=", "2mm", "WhichAxis:=",  _
  "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=", "Cylinder1", "Flags:=",  _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Geometry3DAttributeTab", Array("NAME:PropServers",  _
  "Cylinder1"), Array("NAME:ChangedProps", Array("NAME:Material", "Value:=", "" & Chr(34) & "copper" & Chr(34) & ""))))
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "28mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--PD-LED", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:cypdX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cypdY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cypdH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cypdR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cyledY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cyledH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "subb", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH*2+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), Array("NAME:Attributes", "Name:=", "METGb", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "-cuH", "XSize:=", "subX", "YSize:=", "-CuH", "ZSize:=",  _
  "subH*2+cuH*3+5mm"), Array("NAME:Attributes", "Name:=", "METV", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "subH", "XSize:=", "subX", "YSize:=", "subH", "ZSize:=",  _
  "5mm+cuH*2"), Array("NAME:Attributes", "Name:=", "subv", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=", "patchY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "PATCH", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1b", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "PATCHb", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "28mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--PD-LED", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:cypdX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cypdY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cypdH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cypdR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cyledY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cyledH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "subb", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH*2+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), Array("NAME:Attributes", "Name:=", "METGb", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "-cuH", "XSize:=", "subX", "YSize:=", "-CuH", "ZSize:=",  _
  "subH*2+cuH*3+5mm"), Array("NAME:Attributes", "Name:=", "METV", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "subH", "XSize:=", "subX", "YSize:=", "subH", "ZSize:=",  _
  "5mm+cuH*2"), Array("NAME:Attributes", "Name:=", "subv", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=", "patchY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "PATCH", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1b", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "PATCHb", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateCylinder Array("NAME:CylinderParameters", "XCenter:=", "cypdX", "YCenter:=",  _
  "cypdY", "ZCenter:=", "cuH", "Radius:=", "cypdR", "Height:=", "-cypdH", "WhichAxis:=",  _
  "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=", "CylinderPD", "Flags:=",  _
  "", "Color:=", "(132 1 193)", "Transparency:=", 0.4, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm+SubH*2+CuH+5mm+CuH", "Width:=",  _
  "-subH", "Height:=", "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port2", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "4.42650744617717e-18mm",  _
  "-8.1650628641107e-18mm"), "End:=", Array("6.93889390390723e-18mm",  _
  "1.25958732376599e-17mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("-6.93889390390723e-18mm",  _
  "-3.74285834530558e-18mm", "5.0868mm"), "End:=", Array("0mm",  _
  "4.42650744617717e-18mm", "5.1376mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/4", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
  "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=", "", "Color:=",  _
  "(132 133 0)", "Transparency:=", 0.9, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "air" & Chr(34) & "", "SolveInside:=",  _
  true)
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedField:=", false, "IsFssReference:=", false, "IsForPML:=",  _
  false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=", true)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "UseIterativeSolver:=", true, "IterativeResidual:=",  _
  0.0001, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "EnableSolverDomains:=", false, "SaveRadFieldsOnly:=",  _
  false, "SaveAnyFields:=", true, "NoAdditionalRefinementOnImport:=", false)
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", "19.2GHz", "StopValue:=", "28.8GHz", "Count:=",  _
  100, "Type:=", "Interpolating", "SaveFields:=", false, "InterpTolerance:=",  _
  0.5, "InterpMaxSolns:=", 50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=",  _
  false, "InterpUseS:=", true, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseDerivativeConvergence:=", false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=",  _
  true, "EnforcePassivity:=", false)
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "28mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--PD-LED", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:cypdX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cypdY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:cypdH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cypdR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cyledY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cyledH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "subb", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH*2+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), Array("NAME:Attributes", "Name:=", "METGb", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "-cuH", "XSize:=", "subX", "YSize:=", "-CuH", "ZSize:=",  _
  "subH*2+cuH*3+5mm"), Array("NAME:Attributes", "Name:=", "METV", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "subH", "XSize:=", "subX", "YSize:=", "subH", "ZSize:=",  _
  "5mm+cuH*2"), Array("NAME:Attributes", "Name:=", "subv", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=", "patchY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "PATCH", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1b", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "PATCHb", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateCylinder Array("NAME:CylinderParameters", "XCenter:=", "0", "YCenter:=",  _
  "cypdY", "ZCenter:=", "cuH", "Radius:=", "cypdR", "Height:=", "-cypdH", "WhichAxis:=",  _
  "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=", "CylinderPD", "Flags:=",  _
  "", "Color:=", "(132 1 193)", "Transparency:=", 0.4, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm+SubH*2+CuH+5mm+CuH", "Width:=",  _
  "-subH", "Height:=", "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port2", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "4.42650744617717e-18mm",  _
  "-8.1650628641107e-18mm"), "End:=", Array("6.93889390390723e-18mm",  _
  "1.25958732376599e-17mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("-6.93889390390723e-18mm",  _
  "-3.74285834530558e-18mm", "5.0868mm"), "End:=", Array("0mm",  _
  "4.42650744617717e-18mm", "5.1376mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/4", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
  "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=", "", "Color:=",  _
  "(132 133 0)", "Transparency:=", 0.9, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "air" & Chr(34) & "", "SolveInside:=",  _
  true)
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedField:=", false, "IsFssReference:=", false, "IsForPML:=",  _
  false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=", true)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "UseIterativeSolver:=", true, "IterativeResidual:=",  _
  0.0001, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "EnableSolverDomains:=", false, "SaveRadFieldsOnly:=",  _
  false, "SaveAnyFields:=", true, "NoAdditionalRefinementOnImport:=", false)
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", "19.2GHz", "StopValue:=", "28.8GHz", "Count:=",  _
  100, "Type:=", "Interpolating", "SaveFields:=", false, "InterpTolerance:=",  _
  0.5, "InterpMaxSolns:=", 50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=",  _
  false, "InterpUseS:=", true, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseDerivativeConvergence:=", false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=",  _
  true, "EnforcePassivity:=", false)
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "28mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--PD-LED", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:cypdX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cypdY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:cypdH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"), Array("NAME:cypdR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:cyledX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "12mm"), Array("NAME:cyledY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:cyledH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"), Array("NAME:cyledR", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "subH"), Array("NAME:Attributes", "Name:=", "subb", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "0mm+SubH*2+CuH+5mm+CuH", "XSize:=", "subX", "YSize:=",  _
  "subY", "ZSize:=", "-CuH"), Array("NAME:Attributes", "Name:=", "METGb", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "-cuH", "XSize:=", "subX", "YSize:=", "-CuH", "ZSize:=",  _
  "subH*2+cuH*3+5mm"), Array("NAME:Attributes", "Name:=", "METV", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-1mm", "ZPosition:=", "subH", "XSize:=", "subX", "YSize:=", "subH", "ZSize:=",  _
  "5mm+cuH*2"), Array("NAME:Attributes", "Name:=", "subv", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=", "patchY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "PATCH", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1b", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "0mm+SubH+CuH+5mm", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "PATCHb", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateCylinder Array("NAME:CylinderParameters", "XCenter:=", "0", "YCenter:=",  _
  "cypdY", "ZCenter:=", "cuH", "Radius:=", "cypdR", "Height:=", "-cypdH", "WhichAxis:=",  _
  "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=", "CylinderPD", "Flags:=",  _
  "", "Color:=", "(132 1 193)", "Transparency:=", 0.4, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateCylinder Array("NAME:CylinderParameters", "XCenter:=", "0", "YCenter:=",  _
  "cyledY", "ZCenter:=", "cuH+subH+5mm", "Radius:=", "cyledR", "Height:=",  _
  "cyledH", "WhichAxis:=", "Z", "NumSides:=", "0"), Array("NAME:Attributes", "Name:=",  _
  "CylinderPD_1", "Flags:=", "", "Color:=", "(132 1 193)", "Transparency:=", 0.4, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "vacuum" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm+SubH*2+CuH+5mm+CuH", "Width:=",  _
  "-subH", "Height:=", "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port2", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "4.42650744617717e-18mm",  _
  "-8.1650628641107e-18mm"), "End:=", Array("6.93889390390723e-18mm",  _
  "1.25958732376599e-17mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("-6.93889390390723e-18mm",  _
  "-3.74285834530558e-18mm", "5.0868mm"), "End:=", Array("0mm",  _
  "4.42650744617717e-18mm", "5.1376mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/4", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
  "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=", "", "Color:=",  _
  "(132 133 0)", "Transparency:=", 0.9, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "air" & Chr(34) & "", "SolveInside:=",  _
  true)
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedField:=", false, "IsFssReference:=", false, "IsForPML:=",  _
  false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=", true)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "UseIterativeSolver:=", true, "IterativeResidual:=",  _
  0.0001, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "EnableSolverDomains:=", false, "SaveRadFieldsOnly:=",  _
  false, "SaveAnyFields:=", true, "NoAdditionalRefinementOnImport:=", false)
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", "19.2GHz", "StopValue:=", "28.8GHz", "Count:=",  _
  100, "Type:=", "Interpolating", "SaveFields:=", false, "InterpTolerance:=",  _
  0.5, "InterpMaxSolns:=", 50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=",  _
  false, "InterpUseS:=", true, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseDerivativeConvergence:=", false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=",  _
  true, "EnforcePassivity:=", false)
