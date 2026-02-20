Attribute VB_Name = "Module3"
Sub Populate_AsseticNewAssets()

Cur_Sht = ActiveSheet.Name
Application.StatusBar = True
Application.StatusBar = "Please wait while Assetic New Asset worksheets are being populated!"
Application.ScreenUpdating = False

If Sht_New.Visible = xlSheetVisible Then


Assetic_NewAssets.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":K" & Target_Row + 1000).Select
'Selection.ClearContents
Selection.EntireRow.Delete

Assetic_NewComponent.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":P" & Target_Row + 1000).Select
'Selection.ClearContents
Selection.EntireRow.Delete

Assetic_NewNetworkMeasure.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":I" & Target_Row + 1000).Select
'Selection.ClearContents
Selection.EntireRow.Delete

Assetic_NewValuations.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":W" & Target_Row + 1000).Select
'Selection.ClearContents
Selection.EntireRow.Delete


Sht_Summary.Select

'Get T1 Number
'Get Project Description

PRCode = Range("PR_T1_Number").Cells(1, 1).Value
PRDesc = Range("PR_Project_Name").Cells(1, 1).Value

Sht_New.Select


'Determine Column Locations
  
  For k = 2 To 34
  
  
         
       If InStr(1, Cells(9, k).Value, "Asset Class", vbTextCompare) > 0 Then
        
       Class_Col = k
       
              
       End If



       If InStr(1, Cells(9, k).Value, "Asset Type", vbTextCompare) > 0 Then
        
       Type_Col = k
       
              
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset ID", vbTextCompare) > 0 Then
        
       ID_Col = k
       
              
       End If
       
       If InStr(1, Cells(9, k).Value, "Quantity", vbTextCompare) > 0 Then
        
       Quantity_Col = k
       
              
       End If
            
        
       If InStr(1, Cells(9, k).Value, "Unit Cost", vbTextCompare) > 0 Then
        
       Unit_Col = k
       
              
       End If
            
       If InStr(1, Cells(9, k).Value, "Total Cost", vbTextCompare) > 0 Then
        
       Total_Col = k
       
              
       End If
  
    
  
       If InStr(1, Cells(9, k).Value, "Asset SubClass", vbTextCompare) > 0 Then
        
       SubClass_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset SubType", vbTextCompare) > 0 Then
        
       SubType_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Component Name", vbTextCompare) > 0 Then
        
       Component_Col = k
                     
       End If
       
       
       If InStr(1, Cells(9, k).Value, "Asset Category", vbTextCompare) > 0 Then
        
       AssetCategory_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset Name", vbTextCompare) > 0 Then
        
       AssetName_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Component Type", vbTextCompare) > 0 Then
        
       ComponentType_Col = k
                     
       End If
              
       If InStr(1, Cells(9, k).Value, "Financial Class", vbTextCompare) > 0 Then
        
       FinClass_Col = k
                     
       End If
                    
       If InStr(1, Cells(9, k).Value, "Financial SubClass", vbTextCompare) > 0 Then
        
       FinSubClass_Col = k
                     
       End If
                    
       If InStr(1, Cells(9, k).Value, "Primary Material", vbTextCompare) > 0 Then
        
       Material_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset Network Measure Type", vbTextCompare) > 0 Then
        
       MeasureType_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Unit of Measurement", vbTextCompare) > 0 Then
        
       UoM_Col = k
                     
       End If
      
       If InStr(1, Cells(9, k).Value, "Useful Life", vbTextCompare) = 1 Then
        
       ULife_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Revaluation Date Built", vbTextCompare) > 0 Then
        
       RevDate_Col = k
                     
       End If
      
       If InStr(1, Cells(9, k).Value, "Valuation Record Type", vbTextCompare) > 0 Then
        
       ValRecType_Col = k
                     
       End If
  
       If InStr(1, Cells(9, k).Value, "Valuation Date", vbTextCompare) = 1 Then
        
       InstallDate_Col = k
                     
       End If
  
       'If InStr(1, Cells(9, k).Value, "Remaining Useful Life", vbTextCompare) > 0 Then
        
       'RULife_Col = k
                     
       'End If
  
       If InStr(1, Cells(9, k).Value, "WIP$ New & Upgrade", vbTextCompare) > 0 Then
        
       NewDollar_Col = k
                     
       End If
  
  
  Next k
  
  'Promot error if any columns are not found


  If Class_Col * Type_Col * ID_Col * Quantity_Col * Unit_Col * Total_Col * SubClass_Col * SubType_Col * Component_Col * AssetCategory_Col * AssetName_Col * ComponentType_Col * FinClass_Col * FinSubClass_Col * Material_Col * MeasureType_Col * UoM_Col * ULife_Col * RevDate_Col * ValRecType_Col * InstallDate_Col * NewDollar_Col = 0 Then
  
  errdesx = MsgBox("Incorrect template. There is an issue with new assets sheet - Please contact Governance & Performance!", vbError, "Project Asset Information")

  Exit Sub
  
  
  End If

        Assetic_NewAssets.Cells(1, 1).Value = "Asset Category"
        Assetic_NewAssets.Cells(1, 2).Value = "Asset ID"
        Assetic_NewAssets.Cells(1, 3).Value = "Asset Name"
        Assetic_NewAssets.Cells(1, 4).Value = "Asset Class"
        Assetic_NewAssets.Cells(1, 5).Value = "Asset Sub Class"
        Assetic_NewAssets.Cells(1, 6).Value = "Asset Type"
        Assetic_NewAssets.Cells(1, 7).Value = "Asset Sub Type"
        Assetic_NewAssets.Cells(1, 8).Value = "Maintenance Asset Sub Type"
        Assetic_NewAssets.Cells(1, 9).Value = "Maintenance Asset Type"
        Assetic_NewAssets.Cells(1, 10).Value = "Work Group"
        Assetic_NewAssets.Cells(1, 11).Value = "Criticality"
        Assetic_NewAssets.Cells(1, 12).Value = "Project Code"


        Assetic_NewComponent.Cells(1, 1).Value = "Asset Id"
        Assetic_NewComponent.Cells(1, 2).Value = "Component Name"
        Assetic_NewComponent.Cells(1, 3).Value = "Component Type"
        Assetic_NewComponent.Cells(1, 4).Value = "Financial Class"
        Assetic_NewComponent.Cells(1, 5).Value = "Financial Subclass"
        Assetic_NewComponent.Cells(1, 6).Value = "Primary Material"
        Assetic_NewComponent.Cells(1, 7).Value = "Network Measure Type"
        Assetic_NewComponent.Cells(1, 8).Value = "Unit"
        Assetic_NewComponent.Cells(1, 9).Value = "Weight"
        Assetic_NewComponent.Cells(1, 10).Value = "Threshold"
        Assetic_NewComponent.Cells(1, 11).Value = "Is Critical"
        Assetic_NewComponent.Cells(1, 12).Value = "External Identifier"
        Assetic_NewComponent.Cells(1, 13).Value = "Design Life"
        Assetic_NewComponent.Cells(1, 14).Value = "Reference Value"
        Assetic_NewComponent.Cells(1, 15).Value = "Reference Date"
        Assetic_NewComponent.Cells(1, 16).Value = "Revaluation Date Built"
        
        
        Assetic_NewNetworkMeasure.Cells(1, 1).Value = "Measurement"
        Assetic_NewNetworkMeasure.Cells(1, 2).Value = "Measurement Unit"
        Assetic_NewNetworkMeasure.Cells(1, 3).Value = "Asset Id"
        Assetic_NewNetworkMeasure.Cells(1, 4).Value = "Component Name"
        Assetic_NewNetworkMeasure.Cells(1, 5).Value = "Measurement Record Id"
        Assetic_NewNetworkMeasure.Cells(1, 6).Value = "Record Type"
        Assetic_NewNetworkMeasure.Cells(1, 7).Value = "Multiplier"
        Assetic_NewNetworkMeasure.Cells(1, 8).Value = "Comments"
        Assetic_NewNetworkMeasure.Cells(1, 9).Value = "Measurement Type"


        Assetic_NewValuations.Cells(1, 1).Value = "Valuation Record Id"
        Assetic_NewValuations.Cells(1, 2).Value = "Asset Id"
        Assetic_NewValuations.Cells(1, 3).Value = "Component Name"
        Assetic_NewValuations.Cells(1, 4).Value = "Valuation Component Type"
        Assetic_NewValuations.Cells(1, 5).Value = "Valuation Date"
        Assetic_NewValuations.Cells(1, 6).Value = "Valuation Record Type"
        Assetic_NewValuations.Cells(1, 7).Value = "Date Built"
        Assetic_NewValuations.Cells(1, 8).Value = "Valuation Pattern"
        Assetic_NewValuations.Cells(1, 9).Value = "Valuation Pattern Index"
        Assetic_NewValuations.Cells(1, 10).Value = "Depreciation Method"
        Assetic_NewValuations.Cells(1, 11).Value = "Depreciation Calculation Method"
        Assetic_NewValuations.Cells(1, 12).Value = "Replacement Cost"
        Assetic_NewValuations.Cells(1, 13).Value = "Useful Life"
        Assetic_NewValuations.Cells(1, 14).Value = "Remaining Useful Life"
        Assetic_NewValuations.Cells(1, 15).Value = "Unit Rate"
        Assetic_NewValuations.Cells(1, 16).Value = "Depreciation Rate"
        Assetic_NewValuations.Cells(1, 17).Value = "Depreciation Effective Date"
        Assetic_NewValuations.Cells(1, 18).Value = "Depreciated Replacement Cost"
        Assetic_NewValuations.Cells(1, 19).Value = "Residual Cost (%)"
        Assetic_NewValuations.Cells(1, 20).Value = "Is End Of Day"
        Assetic_NewValuations.Cells(1, 21).Value = "Project Code"
        Assetic_NewValuations.Cells(1, 22).Value = "Description"
        Assetic_NewValuations.Cells(1, 23).Value = "Comments"



'blankCounter
 Blank_Counter = 0

 For i = 10 To Cells.SpecialCells(xlCellTypeLastCell).Row
  
    If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) > 0 Then  'Ignore if there are no units or quantity
    
    'Populate Assetic New Asset
    
    
    Assetic_NewAssets.Cells(Target_Row, 1).Value = Cells(i, AssetCategory_Col).Value
    Assetic_NewAssets.Cells(Target_Row, 2).Value = Cells(i, ID_Col).Value
    Assetic_NewAssets.Cells(Target_Row, 3).Value = Cells(i, AssetName_Col).Value
    Assetic_NewAssets.Cells(Target_Row, 4).Value = Cells(i, Class_Col).Value
    Assetic_NewAssets.Cells(Target_Row, 5).Value = Cells(i, SubClass_Col).Value
    Assetic_NewAssets.Cells(Target_Row, 6).Value = Trim(Cells(i, Type_Col).Value)
    Assetic_NewAssets.Cells(Target_Row, 7).Value = Cells(i, SubType_Col).Value
    'Assetic_NewAssets.Cells(Target_Row, 8).Value = ""
    'Assetic_NewAssets.Cells(Target_Row, 9).Value = ""
    'Assetic_NewAssets.Cells(Target_Row, 10).Value = ""
    'Assetic_NewAssets.Cells(Target_Row, 11).Value = ""
    Assetic_NewAssets.Cells(Target_Row, 12).Value = PRCode
    
    
    'Populate Assetic New Component
    
    Assetic_NewComponent.Cells(Target_Row, 1).Value = Cells(i, ID_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 2).Value = Trim(Cells(i, Component_Col).Value)
    Assetic_NewComponent.Cells(Target_Row, 3).Value = Cells(i, ComponentType_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 4).Value = Cells(i, FinClass_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 5).Value = Cells(i, FinSubClass_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 6).Value = Cells(i, Material_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 7).Value = Cells(i, MeasureType_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 8).Value = Cells(i, UoM_Col).Value
    Assetic_NewComponent.Cells(Target_Row, 9).Value = 1
    'Assetic_NewComponent.Cells(Target_Row, 10).Value = ""
    'Assetic_NewComponent.Cells(Target_Row, 11).Value = ""
    'Assetic_NewComponent.Cells(Target_Row, 12).Value = ""
    Assetic_NewComponent.Cells(Target_Row, 13).Value = Cells(i, ULife_Col).Value
    'Assetic_NewComponent.Cells(Target_Row, 14).Value = ""
    'Assetic_NewComponent.Cells(Target_Row, 15).Value = ""
    Assetic_NewComponent.Cells(Target_Row, 16).Value = Cells(i, RevDate_Col).Value
    
    'Populate Assetic Network Measure
    
    Assetic_NewNetworkMeasure.Cells(Target_Row, 1).Value = Cells(i, Quantity_Col).Value
    Assetic_NewNetworkMeasure.Cells(Target_Row, 2).Value = Cells(i, UoM_Col).Value
    Assetic_NewNetworkMeasure.Cells(Target_Row, 3).Value = Cells(i, ID_Col).Value
    Assetic_NewNetworkMeasure.Cells(Target_Row, 4).Value = Trim(Cells(i, Component_Col).Value)
    Assetic_NewNetworkMeasure.Cells(Target_Row, 5).Value = ""
    Assetic_NewNetworkMeasure.Cells(Target_Row, 6).Value = "Addition"
    Assetic_NewNetworkMeasure.Cells(Target_Row, 7).Value = 1
    'Assetic_NewNetworkMeasure.Cells(Target_Row, 8).Value = ""
    Assetic_NewNetworkMeasure.Cells(Target_Row, 9).Value = Cells(i, MeasureType_Col).Value
    
    'Populate Assetic New Valuations
    
    'Assetic_NewValuations.Cells(Target_Row, 1).Value = ""
    
    FinSubclass = Cells(i, FinSubClass_Col).Value
    
    If FinSubclass = "Public Art, Statues and Monuments" Then
    Assetic_NewValuations.Cells(Target_Row, 8).Value = "None"
    Assetic_NewValuations.Cells(Target_Row, 9).Value = ""
    Assetic_NewValuations.Cells(Target_Row, 10).Value = "None"
    Else
    Assetic_NewValuations.Cells(Target_Row, 8).Value = "Standard Straight Line"
    Assetic_NewValuations.Cells(Target_Row, 9).Value = 0
    Assetic_NewValuations.Cells(Target_Row, 10).Value = "StraightLine"
    
    End If
    
    
    
    Assetic_NewValuations.Cells(Target_Row, 2).Value = Cells(i, ID_Col).Value
    Assetic_NewValuations.Cells(Target_Row, 3).Value = Trim(Cells(i, Component_Col).Value)
    Assetic_NewValuations.Cells(Target_Row, 4).Value = Cells(i, ComponentType_Col).Value
    Assetic_NewValuations.Cells(Target_Row, 5).Value = Cells(i, InstallDate_Col).Value
    Assetic_NewValuations.Cells(Target_Row, 6).Value = Cells(i, ValRecType_Col).Value
    Assetic_NewValuations.Cells(Target_Row, 7).Value = Cells(i, InstallDate_Col).Value
    'Assetic_NewValuations.Cells(Target_Row, 8).Value = "Standard Straight Line"
    'Assetic_NewValuations.Cells(Target_Row, 9).Value = 0
    'Assetic_NewValuations.Cells(Target_Row, 10).Value = "StraightLine"
    Assetic_NewValuations.Cells(Target_Row, 11).Value = "Retrospective"
    Assetic_NewValuations.Cells(Target_Row, 12).Value = Cells(i, NewDollar_Col).Value
    Assetic_NewValuations.Cells(Target_Row, 13).Value = Cells(i, ULife_Col).Value
    'Assetic_NewValuations.Cells(Target_Row, 14).Value = Cells(i, RULife_Col).Value
    
    If Cells(i, Quantity_Col).Value = 0 Or Not IsNumeric(Cells(i, Quantity_Col).Value) Or Not IsNumeric(Cells(i, NewDollar_Col).Value) Then
    
    Assetic_NewValuations.Cells(Target_Row, 15).Value = ""
    
    Else
    
    Assetic_NewValuations.Cells(Target_Row, 15).Value = Cells(i, NewDollar_Col).Value / Cells(i, Quantity_Col).Value
    
    End If
    
    'Assetic_NewValuations.Cells(Target_Row, 16).Value = ""
    'Assetic_NewValuations.Cells(Target_Row, 17).Value = ""
    'Assetic_NewValuations.Cells(Target_Row, 18).Value = ""
    'Assetic_NewValuations.Cells(Target_Row, 19).Value = ""
    Assetic_NewValuations.Cells(Target_Row, 20).Value = "No"
    Assetic_NewValuations.Cells(Target_Row, 21).Value = PRCode
    Assetic_NewValuations.Cells(Target_Row, 22).Value = PRDesc
    'Assetic_NewValuations.Cells(Target_Row, 23).Value = ""


    
    Target_Row = Target_Row + 1
    
    End If
    
    
    
      If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) = 0 Then
    
        Blank_Counter = Blank_Counter + 1
    
       If Blank_Counter > 10 Then
       
        Exit For
       
       End If
     
    End If
    
 
  Next i
  

Assetic_NewAssets.Name = PRCode & "_Assetic_NewAssets"
Assetic_NewComponent.Name = PRCode & "_Assetic_NewComponent"
Assetic_NewNetworkMeasure.Name = PRCode & "_Assetic_NewMeasure"
Assetic_NewValuations.Name = PRCode & "_Assetic_NewValuations"

End If

Application.StatusBar = False
'Application.ScreenUpdating = True

Sheets(Cur_Sht).Select

End Sub
