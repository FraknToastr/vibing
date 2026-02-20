Attribute VB_Name = "Module4"
Sub Populate_AsseticRenewedAssets()

Cur_Sht = ActiveSheet.Name
Application.StatusBar = True
Application.StatusBar = "Please wait while Assetic Renewed Asset worksheets are being populated!"
Application.ScreenUpdating = False

If Sht_Renew.Visible = xlSheetVisible Then


Assetic_CapExRenewals.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":AD" & Target_Row + 1000).Select
'Selection.ClearContents
Selection.EntireRow.Delete


Sht_Summary.Select

'Get T1 Number
'Get Project Description

PRCode = Range("PR_T1_Number").Cells(1, 1).Value
PRDesc = Range("PR_Project_Name").Cells(1, 1).Value

Sht_Renew.Select


'Determine Column Locations
  
  For k = 2 To 45
  
  
         
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
       
       If InStr(1, Cells(9, k).Value, "Component Name", vbTextCompare) = 1 Then
        
       Component_Col = k
                     
       End If
       
        If InStr(1, Cells(9, k).Value, "Valuation Component Name", vbTextCompare) = 1 Then
        
       ValComponent_Col = k
                     
       End If
       
       
       If InStr(1, Cells(9, k).Value, "Asset Category", vbTextCompare) > 0 Then
        
       AssetCategory_Col = k
                     
       End If
       
'
       
       If InStr(1, Cells(9, k).Value, "Component Type", vbTextCompare) > 0 Then
        
       ComponentType_Col = k
                     
       End If
              
       If InStr(1, Cells(9, k).Value, "Financial Class", vbTextCompare) > 0 Then
        
       FinClass_Col = k
                     
       End If
                    
       If InStr(1, Cells(9, k).Value, "Financial SubClass", vbTextCompare) > 0 Then
        
       FinSubClass_Col = k
                     
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
       
       If InStr(1, Cells(9, k).Value, "Date Built", vbTextCompare) > 0 Then
        
       RevDate_Col = k
                     
       End If
      
       If InStr(1, Cells(9, k).Value, "Valuation Record ID", vbTextCompare) > 0 Then
        
       ValRecID_Col = k
                     
       End If
  
       If InStr(1, Cells(9, k).Value, "Valuation Date", vbTextCompare) > 0 Then
        
       InstallDate_Col = k
                     
       End If

       If InStr(1, Cells(9, k).Value, "WIP$ New & Upgrade", vbTextCompare) > 0 Then
        
       NewDollar_Col = k
                     
       End If
  
       If InStr(1, Cells(9, k).Value, "WIP$ Renewal", vbTextCompare) > 0 Then
        
       ReNewDollar_Col = k
                     
       End If
  
        If InStr(1, Cells(9, k).Value, "Comments", vbTextCompare) = 1 Then
        
       Comments_Col = k
                     
       End If
   
         If InStr(1, Cells(9, k).Value, "End of Day", vbTextCompare) = 1 Then
        
       EoD_Col = k
                     
       End If
  
        'If InStr(1, Cells(9, k).Value, "Remaining Useful Life", vbTextCompare) > 0 Then
        
       'RULife_Col = k
                     
       'End If
  
         If InStr(1, Cells(9, k).Value, "Condition Rating", vbTextCompare) > 0 Then
        
       CondRating_Col = k
                     
       End If
  
    
       If InStr(1, Cells(9, k).Value, "Treatment Type", vbTextCompare) > 0 Then
        
       Treatment_Col = k
                     
       End If
  
  
        If InStr(1, Cells(9, k).Value, "% of Asset Renewed", vbTextCompare) > 0 Then
        
       RenewedPerCent_Col = k
                     
       End If
  
  
  Next k
  
  'Promot error if any columns are not found


  If Class_Col * Type_Col * ID_Col * Quantity_Col * Unit_Col * Total_Col * SubClass_Col * SubType_Col * Component_Col * AssetCategory_Col * ComponentType_Col * FinClass_Col * FinSubClass_Col * UoM_Col * ULife_Col * RevDate_Col * ValRecID_Col * InstallDate_Col * NewDollar_Col = 0 Then
  
  errdesx = MsgBox("Incorrect template. There is an issue with renewed assets sheet - Please contact Governance & Performance!", vbError, "Project Asset Information")

  Exit Sub
  
  
  End If

        Assetic_CapExRenewals.Cells(1, 1).Value = "Valuation Record Id"
        Assetic_CapExRenewals.Cells(1, 2).Value = "Asset Id"
        Assetic_CapExRenewals.Cells(1, 3).Value = "Valuation Component Name"
        Assetic_CapExRenewals.Cells(1, 4).Value = "Valuation Date"
        Assetic_CapExRenewals.Cells(1, 5).Value = "Description"
        Assetic_CapExRenewals.Cells(1, 6).Value = "Comments"
        Assetic_CapExRenewals.Cells(1, 7).Value = "Is End Of Day"
        Assetic_CapExRenewals.Cells(1, 8).Value = "Project Code"
        Assetic_CapExRenewals.Cells(1, 9).Value = "Upgrade CapEx"
        Assetic_CapExRenewals.Cells(1, 10).Value = "Upgrade Capitalize WIP"
        Assetic_CapExRenewals.Cells(1, 11).Value = "Upgrade Opex"
        Assetic_CapExRenewals.Cells(1, 12).Value = "Renewal CapEx"
        Assetic_CapExRenewals.Cells(1, 13).Value = "Renewal Capitalize WIP"
        Assetic_CapExRenewals.Cells(1, 14).Value = "Renewal Opex"
        Assetic_CapExRenewals.Cells(1, 15).Value = "Extension CapEx"
        Assetic_CapExRenewals.Cells(1, 16).Value = "Extension Capitalize WIP"
        Assetic_CapExRenewals.Cells(1, 17).Value = "Extension Opex"
        Assetic_CapExRenewals.Cells(1, 18).Value = "Disposal Percentage"
        Assetic_CapExRenewals.Cells(1, 19).Value = "Disposal Expense"
        Assetic_CapExRenewals.Cells(1, 20).Value = "Disposal Proceeds"
        Assetic_CapExRenewals.Cells(1, 21).Value = "WIP Amount"
        Assetic_CapExRenewals.Cells(1, 22).Value = "Residual Value %"
        Assetic_CapExRenewals.Cells(1, 23).Value = "Date Built"
        Assetic_CapExRenewals.Cells(1, 24).Value = "Useful Life"
        Assetic_CapExRenewals.Cells(1, 25).Value = "Valuation Pattern"
        Assetic_CapExRenewals.Cells(1, 26).Value = "Valuation Pattern Index"
        Assetic_CapExRenewals.Cells(1, 27).Value = "Remaining Useful Life"
        Assetic_CapExRenewals.Cells(1, 28).Value = "Calculation Method"
        Assetic_CapExRenewals.Cells(1, 29).Value = "Treatment Name"
        Assetic_CapExRenewals.Cells(1, 30).Value = "Treatment Type"



Blank_Counter = 0

 For i = 10 To Cells.SpecialCells(xlCellTypeLastCell).Row
  
    If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) > 0 Then  'Ignore if there are no units or quantity
    
    'Populate Assetic CapExRenewals
    FinSubclass = Cells(i, FinSubClass_Col).Value
    
    Assetic_CapExRenewals.Cells(Target_Row, 1).Value = Cells(i, ValRecID_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 2).Value = Cells(i, ID_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 3).Value = Cells(i, ValComponent_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 4).Value = Cells(i, InstallDate_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 5).Value = PRDesc
    Assetic_CapExRenewals.Cells(Target_Row, 6).Value = Cells(i, Comments_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 7).Value = Cells(i, EoD_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 8).Value = PRCode
    Assetic_CapExRenewals.Cells(Target_Row, 9).Value = IIf(Cells(i, NewDollar_Col).Value = 0, "", Cells(i, NewDollar_Col).Value)
    'Assetic_CapExRenewals.Cells(Target_Row, 10).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 11).Value = ""
    Assetic_CapExRenewals.Cells(Target_Row, 12).Value = IIf(Cells(i, ReNewDollar_Col).Value = 0, "", Cells(i, ReNewDollar_Col).Value)
    'Assetic_CapExRenewals.Cells(Target_Row, 13).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 14).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 15).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 16).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 17).Value = ""
    Assetic_CapExRenewals.Cells(Target_Row, 18).Value = Cells(i, RenewedPerCent_Col).Value * 100
    'Assetic_CapExRenewals.Cells(Target_Row, 19).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 20).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 21).Value = ""
    'Assetic_CapExRenewals.Cells(Target_Row, 22).Value = ""
    Assetic_CapExRenewals.Cells(Target_Row, 23).Value = Cells(i, RevDate_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 24).Value = Cells(i, ULife_Col).Value
    
    If FinSubclass = "Public Art, Statues and Monuments" Then
    
    Assetic_CapExRenewals.Cells(Target_Row, 25).Value = "None"
    
    Else
        
    Assetic_CapExRenewals.Cells(Target_Row, 25).Value = "Standard Straight Line"
    
    End If
    
    Assetic_CapExRenewals.Cells(Target_Row, 26).Value = IIf(Cells(i, RenewedPerCent_Col).Value = 1, 0, Cells(i, CondRating_Col).Value)
    'Assetic_CapExRenewals.Cells(Target_Row, 27).Value = Cells(i, RULife_Col).Value
    Assetic_CapExRenewals.Cells(Target_Row, 28).Value = IIf(Cells(i, RenewedPerCent_Col).Value = 1, "Retrospective", "Prospective")
    Assetic_CapExRenewals.Cells(Target_Row, 29).Value = Cells(i, Treatment_Col).Value & "-" & Trim(Cells(i, Component_Col).Value) & "-" & Cells(i, ID_Col).Value & "-" & PRCode
    Assetic_CapExRenewals.Cells(Target_Row, 30).Value = Cells(i, Treatment_Col).Value
    
    
 


    
    Target_Row = Target_Row + 1
    
    End If
    
    
      If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) = 0 Then
    
    Blank_Counter = Blank_Counter + 1
    
       If Blank_Counter > 10 Then
       
        Exit For
       
       End If
     
    End If
    
 
  Next i
  


End If




Application.StatusBar = Flase
'Application.ScreenUpdating = True

Sheets(Cur_Sht).Select
Assetic_CapExRenewals.Name = PRCode & "_Assetic_CapExRenewals"


End Sub

