Attribute VB_Name = "Module1"
'Project Asset Register
'Developed by Amer Acosta
'8-November-2018
'Upgraded by Amer Acosta on Jan 2020 to include integration with Assetic and changes to CoA Schema




Sub Populate_Handover_Cost()

Application.Calculation = xlAutomatic

SHt_HandoverCost.Select
Cur_Sht = ActiveSheet.Name
 SHt_HandoverCost.Unprotect ("ips")
  
Application.StatusBar = True
Application.StatusBar = "Please wait while Cost of Project Detials is being generated. This may take a few minutes!"
Application.ScreenUpdating = False
SHt_HandoverCost.Select

'determine target row by looking for the text Cost of Project Details and Category

Target_Row = 0

For i = 15 To Cells.SpecialCells(xlCellTypeLastCell).Row

  If InStr(1, Cells(i, 2).Value, "Cost of Project Details", vbTextCompare) > 0 Then
  
  
        For j = i To i + 7
        
        If InStr(1, Cells(j, 2).Value, "Category", vbTextCompare) > 0 Then
        
        Target_Row = j + 1
        
        Exit For
        
        End If
        
        Next j
  
  Exit For
  
  End If

Next i

'Prompt error if target row not found

If Target_Row = 0 Then

errdesx = MsgBox("Incorrect template. There is an issue with handover cost sheet - Please contact IPS!", vbError, "Project Asset Information")

Exit Sub


End If


'clear existingrows


        Range("A" & Target_Row & ":U" & Target_Row + 1000).Select
        Selection.ClearContents




'Check if project has new assets

If Sht_New.Visible = xlSheetVisible Then


 Sht_New.Select

 'Determine Asset columns Asset Class, Asset Type, Asset5 Id etc
  
  For k = 2 To 26
  
  
         
        
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
  
       If InStr(1, Cells(9, k).Value, "Allocate Project", vbTextCompare) > 0 Then
        
       Allocate_Col = k
       
              
       End If
  
       If InStr(1, Cells(9, k).Value, "Capitalise This", vbTextCompare) > 0 Then
        
       Capitalise_Col = k
       
              
       End If
       
       If InStr(1, Cells(9, k).Value, "Valuation Record ID", vbTextCompare) > 0 Then
        
       FAR_Col = k
       
              
       End If
  
       If InStr(1, Cells(9, k).Value, "Useful Life", vbTextCompare) = 1 Then
        
       Useful_Col = k
       
              
       End If
       
       'added subclass, subtype and component
       
       If InStr(1, Cells(9, k).Value, "Asset SubClass", vbTextCompare) > 0 Then
        
       SubClass_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset SubType", vbTextCompare) > 0 Then
        
       SubType_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Component Name", vbTextCompare) > 0 Then
        
       Component_Col = k
                     
       End If
       
       
       
  
  Next k
  
  'Prompt error if a column cannot be found
  
  If Capitalise_Col = 0 Or Allocate_Col = 0 Or Total_Col = 0 Or Unit_Col = 0 Or Quantity_Col = 0 Or ID_Col = 0 Or Type_Col = 0 Or Class_Col = 0 Then
  
  errdesx = MsgBox("Incorrect template. There is an issue with new assets sheet - Please contact Governance & Performance!", vbError, "Project Asset Information")

  Exit Sub
  
  
  End If
  
  
 
  'populate cost of project details report
  'set blank counter
  Blank_Counter = 0
  
  For i = 10 To Cells.SpecialCells(xlCellTypeLastCell).Row
  
    
    
    
  
    If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) > 0 Then  '+ Len(Cells(i, Total_Col).Value)
    
    Sheets(Cur_Sht).Cells(Target_Row, 2).Value = "New Asset"
    Sheets(Cur_Sht).Cells(Target_Row, 3).Value = Cells(i, Class_Col).Value & "-" & Cells(i, SubClass_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 4).Value = "=Iferror(VLOOKUP(""" & Cells(i, SubClass_Col).Value & """,Asset_Class!A:D,4,FALSE),"""")"
    Sheets(Cur_Sht).Cells(Target_Row, 5).Value = Format(0, "0.00%")
    Sheets(Cur_Sht).Cells(Target_Row, 6).Value = Cells(i, SubClass_Col).Value & "-" & Cells(i, Type_Col).Value & "-" & Cells(i, SubType_Col).Value & "-" & Cells(i, Component_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 7).Value = Cells(i, ID_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 8).Value = Cells(i, Quantity_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 9).Value = Format(Cells(i, Unit_Col).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 10).Value = Format(Cells(i, Total_Col).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 11).Value = Cells(i, Allocate_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 12).Value = "=IF(K" & Target_Row & "=""Yes""," & "J" & Target_Row & "/SUMIF(K:K,""Yes"",J:J)*PW_Total_Costs,0)"
    Sheets(Cur_Sht).Cells(Target_Row, 13).Value = "=J" & Target_Row & "+L" & Target_Row
    Sheets(Cur_Sht).Cells(Target_Row, 14).Value = "=If(B" & Target_Row & "=""Write-off"",0,Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_CY_Expenditure)*FI_Overhead_Percentage,0)+ Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_Prev_Overhead),0))"
    Sheets(Cur_Sht).Cells(Target_Row, 15).Value = Cells(i, Capitalise_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 16).Value = IIf(Cells(i, Capitalise_Col).Value = "No", "", "=M" & Target_Row & "+N" & Target_Row)
    Sheets(Cur_Sht).Cells(Target_Row, 17).Value = "=P" & Target_Row & "*(1-E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 18).Value = "=P" & Target_Row & "*(E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 19).Value = IIf(Cells(i, Capitalise_Col).Value = "No", "=M" & Target_Row & "+N" & Target_Row, "")
    Sheets(Cur_Sht).Cells(Target_Row, 20).Value = Cells(i, FAR_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 21).Value = Cells(i, Useful_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 1).Value = "New Assets:" & i
    
    
    
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

' end new project assets

'check if project has renewed assets

If Sht_Renew.Visible = xlSheetVisible Then



 Sht_Renew.Select
 
 
 Capitalise_Col = 0
 Allocate_Col = 0
 Total_Col = 0
 Unit_Col = 0
 Quantity_Col = 0
 ID_Col = 0
 Type_Col = 0
 Class_Col = 0
 

 'Determine Asset columns
  
  For k = 2 To 40
  
  
         
        
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
  
       If InStr(1, Cells(9, k).Value, "Allocate Project", vbTextCompare) > 0 Then
        
       Allocate_Col = k
       
              
       End If
  
       If InStr(1, Cells(9, k).Value, "Capitalise This", vbTextCompare) > 0 Then
        
       Capitalise_Col = k
       
              
       End If
  
       If InStr(1, Cells(9, k).Value, "Upgrade (%)", vbTextCompare) > 0 Then
        
       Renewal_Col = k
       
              
       End If
  


       If InStr(1, Cells(9, k).Value, "Valuation Record ID", vbTextCompare) > 0 Then
        
       FAR_Col = k
       
              
       End If
  
       If InStr(1, Cells(9, k).Value, "Useful Life", vbTextCompare) > 0 Then
        
       Useful_Col = k
       
              
       End If


       'added subclass, subtype and component
       
       If InStr(1, Cells(9, k).Value, "Asset SubClass", vbTextCompare) > 0 Then
        
       SubClass_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset SubType", vbTextCompare) > 0 Then
        
       SubType_Col = k
                     
       End If
       
       If InStr(1, Cells(9, k).Value, "Component Name", vbTextCompare) = 1 Then
        
       Component_Col = k
                     
       End If
       


  
  Next k
  
  
  'Prompt if any column is not found
  
  If Capitalise_Col = 0 Or Allocate_Col = 0 Or Total_Col = 0 Or Unit_Col = 0 Or Quantity_Col = 0 Or ID_Col = 0 Or Type_Col = 0 Or Class_Col = 0 Or Renewal_Col = 0 Then
  
  errdesx = MsgBox("Incorrect template. There is an issue with renewed assets sheet - Please contact Governance & Performance!", vbError, "Project Asset Information")

  Exit Sub
  
  
  End If
  
  'Populate Cost of Project Details
    'set blank counter
  Blank_Counter = 0
  
  
  
  For i = 10 To Cells.SpecialCells(xlCellTypeLastCell).Row
  
    If Len(Cells(i, Quantity_Col).Value) + Len(Cells(i, Unit_Col).Value) > 0 Then  '+ Len(Cells(i, Total_Col).Value)
    
    Sheets(Cur_Sht).Cells(Target_Row, 2).Value = "Renewed Asset"
    Sheets(Cur_Sht).Cells(Target_Row, 3).Value = Cells(i, Class_Col).Value & "-" & Cells(i, SubClass_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 4).Value = "=Iferror(VLOOKUP(""" & Cells(i, SubClass_Col).Value & """,Asset_Class!A:D,4,FALSE),"""")"
    Sheets(Cur_Sht).Cells(Target_Row, 5).Value = Format(1 - Cells(i, Renewal_Col).Value, "0.00%")
    Sheets(Cur_Sht).Cells(Target_Row, 6).Value = Cells(i, Type_Col).Value & "-" & Cells(i, SubType_Col).Value & "-" & Cells(i, Component_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 7).Value = Cells(i, ID_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 8).Value = Cells(i, Quantity_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 9).Value = Format(Cells(i, Unit_Col).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 10).Value = Format(Cells(i, Total_Col).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 11).Value = Cells(i, Allocate_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 12).Value = "=IF(K" & Target_Row & "=""Yes""," & "J" & Target_Row & "/SUMIF(K:K,""Yes"",J:J)*PW_Total_Costs,0)"
    Sheets(Cur_Sht).Cells(Target_Row, 13).Value = "=J" & Target_Row & "+L" & Target_Row
    Sheets(Cur_Sht).Cells(Target_Row, 14).Value = "=If(B" & Target_Row & "=""Write-off"",0,Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_CY_Expenditure)*FI_Overhead_Percentage,0)+ Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_Prev_Overhead),0))"
    Sheets(Cur_Sht).Cells(Target_Row, 15).Value = Cells(i, Capitalise_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 16).Value = IIf(Cells(i, Capitalise_Col).Value = "No", "", "=M" & Target_Row & "+N" & Target_Row)
    Sheets(Cur_Sht).Cells(Target_Row, 17).Value = "=P" & Target_Row & "*(1-E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 18).Value = "=P" & Target_Row & "*(E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 19).Value = IIf(Cells(i, Capitalise_Col).Value = "No", "=M" & Target_Row & "+N" & Target_Row, "")
    Sheets(Cur_Sht).Cells(Target_Row, 20).Value = Cells(i, FAR_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 21).Value = Cells(i, Useful_Col).Value
    Sheets(Cur_Sht).Cells(Target_Row, 1).Value = "Renewed Assets:" & i
    
    
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

' end renew project assets

'add write off costs

 Sht_ProjectWide.Select
 
 'set blank Counter
 Blank_Counter = 0
 
   For i = 10 To 1012
    
   If Cells(i, 2).Value = "Write-Off" Then
   
    Sheets(Cur_Sht).Cells(Target_Row, 2).Value = "Write-off"
     
    Sheets(Cur_Sht).Cells(Target_Row, 3).Value = Cells(i, 3).Value
    Sheets(Cur_Sht).Cells(Target_Row, 4).Value = ""
    Sheets(Cur_Sht).Cells(Target_Row, 5).Value = ""
    Sheets(Cur_Sht).Cells(Target_Row, 6).Value = Cells(i, 4).Value
    Sheets(Cur_Sht).Cells(Target_Row, 7).Value = ""
    Sheets(Cur_Sht).Cells(Target_Row, 8).Value = 1
    Sheets(Cur_Sht).Cells(Target_Row, 9).Value = Format(Cells(i, 5).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 10).Value = Format(Cells(i, 5).Value, "$#,##0.00")
    Sheets(Cur_Sht).Cells(Target_Row, 11).Value = "No"
    Sheets(Cur_Sht).Cells(Target_Row, 12).Value = "=IF(K" & Target_Row & "=""Yes""," & "J" & Target_Row & "/SUMIF(K:K,""Yes"",J:J)*PW_Total_Costs,0)"
    Sheets(Cur_Sht).Cells(Target_Row, 13).Value = "=J" & Target_Row & "+L" & Target_Row
    Sheets(Cur_Sht).Cells(Target_Row, 14).Value = "=If(B" & Target_Row & "=""Write-off"",0,Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_CY_Expenditure)*FI_Overhead_Percentage,0)+ Iferror((M" & Target_Row & "/CPD_Total_Assets_Costs*FI_Prev_Overhead),0))"
    Sheets(Cur_Sht).Cells(Target_Row, 15).Value = "No"
    Sheets(Cur_Sht).Cells(Target_Row, 16).Value = IIf("No" = "No", "", "=M" & Target_Row & "+N" & Target_Row)
    Sheets(Cur_Sht).Cells(Target_Row, 17).Value = "=P" & Target_Row & "*(1-E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 18).Value = "=P" & Target_Row & "*(E" & Target_Row & ")"
    Sheets(Cur_Sht).Cells(Target_Row, 19).Value = IIf("No" = "No", "=M" & Target_Row & "+N" & Target_Row, "")
    Sheets(Cur_Sht).Cells(Target_Row, 20).Value = ""
    Sheets(Cur_Sht).Cells(Target_Row, 21).Value = ""
    
    Target_Row = Target_Row + 1
   
   
   End If
   
   
    If Len(Cells(i, 2).Value) = 0 Then
    
     Blank_Counter = Blank_Counter + 1
    
       If Blank_Counter > 10 Then
       
        Exit For
       
       End If
     
    End If
    
   
   

   Next i
   
'End add writeoff costs


Application.StatusBar = False
'Application.ScreenUpdating = True

Sheets(Cur_Sht).Select
SHt_HandoverCost.Protect ("ips")


End Sub
