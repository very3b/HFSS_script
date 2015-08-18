' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 2014.0.2
' 19:31:10  Dec 30, 2014
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
Set oProject = oDesktop.SetActiveProject("Project36")
Set oDesign = oProject.SetActiveDesign("FPC_Patch_Antenna")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedHField:=", false, "IsEnforcedEField:=", false, "IsFssReference:=",  _
  false, "IsForPML:=", false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=",  _
  true)
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_Patch_Antenna", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_Patch_Antenna")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "75mm"), Array("NAME:--AirBox", "PropType:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-subX/2", "YStart:=", "-subY/2", "ZStart:=", "0mm", "Width:=", "subX", "Height:=",  _
  "subY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=", "Ground", "Flags:=",  _
  "", "Color:=", "(255 128 65)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-patchX/2", "YStart:=", "subY/2-patchY-patch_edge", "ZStart:=", "subH", "Width:=",  _
  "patchX", "Height:=", "patchY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignPerfectE Array("NAME:perfE", "Objects:=", Array("patch"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "subY/2-patchY-patch_edge-feedY", "ZStart:=", "subH", "Width:=",  _
  "feedX", "Height:=", "feedY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "feed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignPerfectE Array("NAME:perfE1", "Objects:=", Array("feed"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "subY/2-patchY-patch_edge-feedY", "ZStart:=", "0mm", "Width:=",  _
  "subH", "Height:=", "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "2.5mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "2.5mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "UseDomains:=", false, "UseIterativeSolver:=",  _
  true, "IterativeResidual:=", 0.0001, "SaveRadFieldsOnly:=", false, "SaveAnyFields:=",  _
  true, "NoAdditionalRefinementOnImport:=", false)
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", "12GHz", "StopValue:=", "36GHz", "Count:=", 200, "Type:=",  _
  "Interpolating", "SaveFields:=", false, "InterpTolerance:=", 0.5, "InterpMaxSolns:=",  _
  50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=", false, "InterpUseS:=",  _
  true, "InterpUsePortImped:=", false, "InterpUsePropConst:=", true, "UseDerivativeConvergence:=",  _
  false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=", true, "EnforcePassivity:=",  _
  true, "PassivityErrorTolerance:=", 0.0001)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/2", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
  "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=", "", "Color:=",  _
  "(132 133 0)", "Transparency:=", 0.9, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "air" & Chr(34) & "", "SolveInside:=",  _
  true)
oDesktop.DeleteProject "Project30"
oDesktop.DeleteProject "Project31"
oDesktop.DeleteProject "Project32"
oDesktop.DeleteProject "Project33"
oDesktop.DeleteProject "Project34"
oDesktop.DeleteProject "Project37"
oDesktop.DeleteProject "Project36"
oDesktop.DeleteProject "Project35"
Set oProject = oDesktop.NewProject
oProject.InsertDesign "HFSS", "FPC_Patch_Antenna", "DrivenModal", ""
Set oDesign = oProject.SetActiveDesign("FPC_Patch_Antenna")
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _
  "LocalVariables"), Array("NAME:NewProps", Array("NAME:--Patch Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:patchX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patchY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:patch_edge", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:--Substrate Dimensions", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:subH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0508mm"), Array("NAME:subX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "15mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "75mm"), Array("NAME:--AirBox", "PropType:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-subX/2", "YStart:=", "-subY/2", "ZStart:=", "0mm", "Width:=", "subX", "Height:=",  _
  "subY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=", "Ground", "Flags:=",  _
  "", "Color:=", "(255 128 65)", "Transparency:=", 0.8, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignPerfectE Array("NAME:Ground", "Objects:=", Array("Ground"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-patchX/2", "YStart:=", "subY/2-patchY-patch_edge", "ZStart:=", "subH", "Width:=",  _
  "patchX", "Height:=", "patchY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "patch", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignPerfectE Array("NAME:perfE", "Objects:=", Array("patch"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "subY/2-patchY-patch_edge-feedY", "ZStart:=", "subH", "Width:=",  _
  "feedX", "Height:=", "feedY", "WhichAxis:=", "Z"), Array("NAME:Attributes", "Name:=",  _
  "feed", "Flags:=", "", "Color:=", "(255 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignPerfectE Array("NAME:perfE1", "Objects:=", Array("feed"), "InfGroundPlane:=",  _
  false)
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "subY/2-patchY-patch_edge-feedY", "ZStart:=", "0mm", "Width:=",  _
  "subH", "Height:=", "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=",  _
  "port1", "Flags:=", "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
  false)
oModule.AssignLumpedPort Array("NAME:1", "Objects:=", Array("port1"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "2.5mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "2.5mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "UseDomains:=", false, "UseIterativeSolver:=",  _
  true, "IterativeResidual:=", 0.0001, "SaveRadFieldsOnly:=", false, "SaveAnyFields:=",  _
  true, "NoAdditionalRefinementOnImport:=", false)
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", "IsEnabled:=", true, "SetupType:=",  _
  "LinearCount", "StartValue:=", "12GHz", "StopValue:=", "36GHz", "Count:=", 200, "Type:=",  _
  "Interpolating", "SaveFields:=", false, "InterpTolerance:=", 0.5, "InterpMaxSolns:=",  _
  50, "InterpMinSolns:=", 0, "InterpMinSubranges:=", 1, "ExtrapToDC:=", false, "InterpUseS:=",  _
  true, "InterpUsePortImped:=", false, "InterpUsePropConst:=", true, "UseDerivativeConvergence:=",  _
  false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=", true, "EnforcePassivity:=",  _
  true, "PassivityErrorTolerance:=", 0.0001)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/2", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
  "airH"), Array("NAME:Attributes", "Name:=", "airbox", "Flags:=", "", "Color:=",  _
  "(132 133 0)", "Transparency:=", 0.9, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "air" & Chr(34) & "", "SolveInside:=",  _
  true)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignRadiation Array("NAME:Rad1", "Objects:=", Array("airbox"), "IsIncidentField:=",  _
  false, "IsEnforcedHField:=", false, "IsEnforcedEField:=", false, "IsFssReference:=",  _
  false, "IsForPML:=", false, "UseAdaptiveIE:=", false, "IncludeInPostproc:=",  _
  true)
oProject.SaveAs "/home/jiafu/Ansoft/Project30.hfss", true
oDesign.AnalyzeAll
