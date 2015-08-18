' ----------------------------------------------
' Script Recorded by Ansoft HFSS Version 14.0.1
' 3:03:30 PM  Dec 31, 2014
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
Set oProject = oDesktop.SetActiveProject("FPC_MSline_tissue")
Set oDesign = oProject.SetActiveDesign("FPC_MSLine")
Set oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/FPC_MSline_tissue_FPC_MSLine.s2p", Array("All"), true, 50,  _
  "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_" & Chr(39) & "&feedY&" & Chr(39) & ".s2p", Array( _
  "All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.EditFrequencySweep "Setup1", "Sweep1", Array("NAME:Sweep1", "IsEnabled:=",  _
  true, "SetupType:=", "LinearCount", "StartValue:=", "19.2GHz", "StopValue:=",  _
  "28.8GHz", "Count:=", 10, "Type:=", "Interpolating", "SaveFields:=", false, "SaveRadFields:=",  _
  false, "InterpTolerance:=", 0.5, "InterpMaxSolns:=", 50, "InterpMinSolns:=", 0, "InterpMinSubranges:=",  _
  1, "ExtrapToDC:=", false, "InterpUseS:=", true, "InterpUsePortImped:=", false, "InterpUsePropConst:=",  _
  true, "UseDerivativeConvergence:=", false, "InterpDerivTolerance:=", 0.2, "UseFullBasis:=",  _
  true, "EnforcePassivity:=", false)
oModule.EditSetup "Setup1", Array("NAME:Setup1", "Frequency:=", "24GHz", "PortsOnly:=",  _
  false, "MaxDeltaS:=", 0.02, "UseMatrixConv:=", false, "MaximumPasses:=", 15, "MinimumPasses:=",  _
  1, "MinimumConvergedPasses:=", 1, "PercentRefinement:=", 30, "IsEnabled:=",  _
  true, "BasisOrder:=", 1, "UseIterativeSolver:=", true, "IterativeResidual:=",  _
  0.0001, "DoLambdaRefine:=", true, "DoMaterialLambda:=", true, "SetLambdaTarget:=",  _
  false, "Target:=", 0.3333, "UseMaxTetIncrease:=", false, "PortAccuracy:=", 2, "UseABCOnPort:=",  _
  false, "SetPortMinMaxTri:=", false, "EnableSolverDomains:=", false, "SaveRadFieldsOnly:=",  _
  false, "SaveAnyFields:=", true, "NoAdditionalRefinementOnImport:=", false)
Set oModule = oDesign.GetModule("Optimetrics")
oModule.EditSetup "ParametricSetup1", Array("NAME:ParametricSetup1", Array("NAME:ProdOptiSetupData", "SaveFields:=",  _
  false, "CopyMesh:=", false), Array("NAME:StartingPoint"), "Sim. Setups:=", Array( _
  "Setup1"), Array("NAME:Sweeps", Array("NAME:SweepDefinition", "Variable:=", "feedY", "Data:=",  _
  "LIN 10mm 40mm 1mm", "OffsetF1:=", false, "Synchronize:=", 0)), Array("NAME:Sweep Operations"), Array("NAME:Goals"))
oProject.Save
oDesign.AnalyzeAll
Set oModule = oDesign.GetModule("Solutions")
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
oProject.Save
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "40mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:Sweep1"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_300.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30.s2p", Array("All"), true, 50, "S", -1, 0, 15
oProject.Save
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30.s2p", Array("All"), true, 50, "S", -1, 0, 15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "10mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_10mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "11mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_11mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "12mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_12mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "13mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_13mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "14mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_14mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "15mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_15mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "16mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_16mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "17mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_17mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "18mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_18mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "19mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_19mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "20mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_20mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "21mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_21mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "22mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_22mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "23mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_23mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "24mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_24mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "25mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_25mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "26mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_26mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "27mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_27mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "28mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_28mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "29mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_29mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oModule.ExportNetworkData  _
  "airH=" & Chr(39) & "10mm" & Chr(39) & " airX=" & Chr(39) & "20mm" & Chr(39) & "" & _ 
  " airY=" & Chr(39) & "75mm" & Chr(39) & " bloodH=" & Chr(39) & "3mm" & Chr(39) & "" & _ 
  " CuH=" & Chr(39) & "0.018mm" & Chr(39) & " earX=" & Chr(39) & "15mm" & Chr(39) & "" & _ 
  " earY=" & Chr(39) & "20mm" & Chr(39) & " fatH=" & Chr(39) & "1mm" & Chr(39) & "" & _ 
  " feed_length=" & Chr(39) & "20mm" & Chr(39) & " feedX=" & Chr(39) & "0.0889mm" & Chr(39) & "" & _ 
  " feedY=" & Chr(39) & "30mm" & Chr(39) & " patch_edge=" & Chr(39) & "5mm" & Chr(39) & "" & _ 
  " patchX=" & Chr(39) & "10mm" & Chr(39) & " patchY=" & Chr(39) & "10mm" & Chr(39) & "" & _ 
  " subH=" & Chr(39) & "0.0508mm" & Chr(39) & " subX=" & Chr(39) & "11mm" & Chr(39) & "" & _ 
  " subY=" & Chr(39) & "35mm" & Chr(39) & "", Array("Setup1:LastAdaptive"), 3,  _
  "/home/jiafu/Ansoft/test_feedY_30mm.s2p", Array("All"), true, 50, "S", -1, 0,  _
  15
oProject.Save
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "35mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "75mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
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
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "0mm", "YPosition:=",  _
  "-feedY/2", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "-patchY/2+feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "MET1_1", "Flags:=",  _
  "", "Color:=", "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
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
  "-airY/2", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "35mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "35mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "-patchY/2+feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=",  _
  "patchY", "ZSize:=", "CuH"), Array("NAME:Attributes", "Name:=", "MET1_1", "Flags:=",  _
  "", "Color:=", "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
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
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "-2.5mm",  _
  "-8.1650628641107e-18mm"), "End:=", Array("6.93889390390723e-18mm", "-2.5mm",  _
  "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=", 0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=",  _
  false, "ReporterFilter:=", Array(true), "FullResistance:=", "50ohm", "FullReactance:=",  _
  "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "2.5mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "2.5mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "30mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-feedX/2", "YPosition:=",  _
  "0", "ZPosition:=", "subH", "XSize:=", "feedX", "YSize:=", "feedY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-patchX/2", "YPosition:=",  _
  "feedY", "ZPosition:=", "subH", "XSize:=", "patchX", "YSize:=", "patchY", "ZSize:=",  _
  "CuH"), Array("NAME:Attributes", "Name:=", "MET1_1", "Flags:=", "", "Color:=",  _
  "(1 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
  false)
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
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "-2.5mm",  _
  "-8.1650628641107e-18mm"), "End:=", Array("6.93889390390723e-18mm", "-2.5mm",  _
  "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=", 0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=",  _
  false, "ReporterFilter:=", Array(true), "FullResistance:=", "50ohm", "FullReactance:=",  _
  "0ohm")
oModule.AssignLumpedPort Array("NAME:2", "Objects:=", Array("port2"), "RenormalizeAllTerminals:=",  _
  true, "DoDeembed:=", false, Array("NAME:Modes", Array("NAME:Mode1", "ModeNum:=", 1, "UseIntLine:=",  _
  true, Array("NAME:IntLine", "Start:=", Array("0mm", "2.5mm", "-8.1650628641107e-18mm"), "End:=", Array( _
  "6.93889390390723e-18mm", "2.5mm", "0.0508mm")), "CharImp:=", "Zpi", "AlignmentGroup:=",  _
  0, "RenormImp:=", "50ohm")), "ShowReporterFilter:=", false, "ReporterFilter:=", Array( _
  true), "FullResistance:=", "50ohm", "FullReactance:=", "0ohm")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "70mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
  "Global", "UDMId:=", "", "MaterialValue:=", "" & Chr(34) & "pec" & Chr(34) & "", "SolveInside:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "70mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-airX/2", "YPosition:=",  _
  "-airY/2", "ZPosition:=", "-airH/5", "XSize:=", "airX", "YSize:=", "airY", "ZSize:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "80mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
oDesktop.DeleteProject "Project39"
oDesktop.DeleteProject "Project37"
oDesktop.DeleteProject "Project38"
oDesktop.DeleteProject "Project34"
oDesktop.DeleteProject "Project36"
oDesktop.DeleteProject "Project33"
oDesktop.DeleteProject "Project32"
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "5mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:feedX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.0889mm"), Array("NAME:feedY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "25mm"), Array("NAME:feed_length", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"))))
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.AddMaterial Array("NAME:AP LaminateFPC", "CoordinateSystemType:=",  _
  "Cartesian", Array("NAME:AttachedData"), Array("NAME:ModifierData"), "permittivity:=",  _
  "3.4", "dielectric_loss_tangent:=", "0.004")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
  "VariableProp", "UserDef:=", true, "Value:=", "13.2mm"), Array("NAME:subY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "40mm"), Array("NAME:--Copper", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:CuH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "0.018mm"), Array("NAME:--AirBox", "PropType:=",  _
  "SeparatorProp", "UserDef:=", true, "Value:=", ""), Array("NAME:airX", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "20mm"), Array("NAME:airY", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "60mm"), Array("NAME:airH", "PropType:=",  _
  "VariableProp", "UserDef:=", true, "Value:=", "10mm"), Array("NAME:--Feed", "PropType:=",  _
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
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "subH"), Array("NAME:Attributes", "Name:=", "sub", "Flags:=", "", "Color:=",  _
  "(132 132 193)", "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "AP LaminateFPC" & Chr(34) & "", "SolveInside:=",  _
  true)
oEditor.CreateBox Array("NAME:BoxParameters", "XPosition:=", "-subX/2", "YPosition:=",  _
  "-5mm", "ZPosition:=", "0mm", "XSize:=", "subX", "YSize:=", "subY", "ZSize:=",  _
  "-CuH"), Array("NAME:Attributes", "Name:=", "METG", "Flags:=", "", "Color:=",  _
  "(255 252 1)", "Transparency:=", 0.3, "PartCoordinateSystem:=", "Global", "UDMId:=",  _
  "", "MaterialValue:=", "" & Chr(34) & "copper" & Chr(34) & "", "SolveInside:=",  _
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
oEditor.CreateRectangle Array("NAME:RectangleParameters", "IsCovered:=", true, "XStart:=",  _
  "-feedX/2", "YStart:=", "0", "ZStart:=", "0mm", "Width:=", "subH", "Height:=",  _
  "feedX", "WhichAxis:=", "Y"), Array("NAME:Attributes", "Name:=", "port1", "Flags:=",  _
  "", "Color:=", "(128 128 65)", "Transparency:=", 0.3, "PartCoordinateSystem:=",  _
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
