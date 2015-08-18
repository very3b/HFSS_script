' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 14.0.1
' 12:16:38 PM  Dec 31, 2014
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
Set oProject = oDesktop.SetActiveProject("Project_MSLine_script")
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:fat", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "conductivity:=",  _
  "1.65")
oDefinitionManager.RemoveMaterial "fat", true, "", "Project"
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_MSLine", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "11mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "35mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "75mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-subY/2", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-subY/2", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "-feedY/2", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "-feedY/2", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "feedY/2", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port2", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "-10mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "-10mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "10mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "10mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/2", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
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
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Human Part", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:earX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:earY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:bloddH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "3mm"), Array("NAME:fatH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"))))
oDefinitionManager.AddMaterial Array("NAME:Human_Blood", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "30.2", "conductivity:=", "40.1")
oDefinitionManager.AddMaterial Array("NAME:Human_Fat", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "30.2", "conductivity:=", "3.82")
oDesign.Undo
oDesign.Undo
oDesign.Undo
Set oModule = oDesign.GetModule("BoundarySetup")
oDesign.Undo
oDesign.Undo
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Human Part", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:earX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:earY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:bloddH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "3mm"), Array("NAME:fatH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"))))
oDefinitionManager.AddMaterial Array("NAME:Human_Blood0", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "30.2", "conductivity:=", "40.1")
oDefinitionManager.AddMaterial Array("NAME:Human_Fat0", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.82", "conductivity:=", "3.82")
oDesign.Undo
oDesign.Undo
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Human Part", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:earX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:earY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:bloddH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "3mm"), Array("NAME:fatH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "1mm"))))
oDefinitionManager.AddMaterial Array("NAME:Human_Blood1", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "30.2", "conductivity:=", "40.1")
oDefinitionManager.AddMaterial Array("NAME:Human_Fat1", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.82", "conductivity:=", "3.82")
