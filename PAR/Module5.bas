Attribute VB_Name = "Module5"
Sub Populate_AsseticDisposedAssets()

Cur_Sht = ActiveSheet.Name
Application.StatusBar = True
Application.StatusBar = "Please wait while Assetic Disposed Asset worksheets are being populated!"
Application.ScreenUpdating = False

If Sht_Dispose.Visible = xlSheetVisible Then


Assetic_DisposedAssets.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":F" & Target_Row + 1000).Select
Selection.EntireRow.Delete

Assetic_DisposedValuations.Select
Target_Row = 2
'Clear Target_Row
Range("A" & Target_Row & ":L" & Target_Row + 1000).Select
Selection.EntireRow.Delete


Sht_Summary.Select

'Get T1 Number
'Get Project Description

PRCode = Range("PR_T1_Number").Cells(1, 1).Value
PRDesc = Range("PR_Project_Name").Cells(1, 1).Value

DisposalType_Col = 12

Sht_Dispose.Select


'Determine Column Locations
  
  For k = 2 To 20
  
  
         
       If InStr(1, Cells(9, k).Value, "Asset Class", vbTextCompare) > 0 Then
        
       Class_Col = k
       
              
       End If



       If InStr(1, Cells(9, k).Value, "Asset Type", vbTextCompare) > 0 Then
        
       Type_Col = k
       
              
       End If
       
       If InStr(1, Cells(9, k).Value, "Asset ID", vbTextCompare) > 0 Then
        
       ID_Col = k
       
              
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
       
             
       If InStr(1, Cells(9, k).Value, "Asset Name", vbTextCompare) > 0 Then
        
       AssetName_Col = k
                     
       End If
       
       


      
       If InStr(1, Cells(9, k).Value, "Valuation Record ID", vbTextCompare) > 0 Then
        
       ValRecID_Col = k
                     
       End If
  
       If InStr(1, Cells(9, k).Value, "Disposal Date", vbTextCompare) > 0 Then
        
       DisposeDate_Col = k
                     
       End If


       If InStr(1, Cells(9, k).Value, "Reason", vbTextCompare) > 0 Then
        
       Reason_Col = k
                     
       End If



       If InStr(1, Cells(9, k).Value, "Valuation Component Name", vbTextCompare) > 0 Then
        
       ValCompName_Col = k
                     
       End If


       If InStr(1, Cells(9, k).Value, "Valuation Date", vbTextCompare) > 0 Then
        
       ValDate_Col = k
                     
       End If


      If InStr(1, Cells(9, k).Value, "Valuation Record Type", vbTextCompare) > 0 Then
        
       ValRecType_Col = k
                     
       End If
       
       
      If InStr(1, Cells(9, k).Value, "Comments", vbTextCompare) > 0 Then
        
       Comments_Col = k
                     
       End If
       
        If InStr(1, Cells(9, k).Value, "Disposal Type", vbTextCompare) > 0 Then
        
       DisposalType_Col = k
                     
       End If
       
  
  Next k
  
  'Promot error if any columns are not found


  If Class_Col * Type_Col * ID_Col * SubClass_Col * SubType_Col * Component_Col * AssetName_Col * ValRecID_Col * DisposeDate_Col = 0 Then
  
  errdesx = MsgBox("Incorrect template. There is an issue with renewed assets sheet - Please contact Governance & Performance!", vbError, "Project Asset Information")

  Exit Sub
  
  
  End If


    Assetic_DisposedAssets.Cells(1, 1).Value = "Asset Id"
    Assetic_DisposedAssets.Cells(1, 2).Value = "To State"
    Assetic_DisposedAssets.Cells(1, 3).Value = "Buyer"
    Assetic_DisposedAssets.Cells(1, 4).Value = "Sell Value"
    Assetic_DisposedAssets.Cells(1, 5).Value = "Disposal Date"
    Assetic_DisposedAssets.Cells(1, 6).Value = "Reason"


    Assetic_DisposedValuations.Cells(1, 1).Value = "Valuation Record Id"
    Assetic_DisposedValuations.Cells(1, 2).Value = "Asset Id"
    Assetic_DisposedValuations.Cells(1, 3).Value = "Component Name"
    Assetic_DisposedValuations.Cells(1, 4).Value = "Valuation Component Name"
    Assetic_DisposedValuations.Cells(1, 5).Value = "Valuation Date"
    Assetic_DisposedValuations.Cells(1, 6).Value = "Valuation Record Type"
    Assetic_DisposedValuations.Cells(1, 7).Value = "Is End Of Day"
    Assetic_DisposedValuations.Cells(1, 8).Value = "Disposal Proceeds"
    Assetic_DisposedValuations.Cells(1, 9).Value = "Disposal Expense"
    Assetic_DisposedValuations.Cells(1, 10).Value = "Project Code"
    Assetic_DisposedValuations.Cells(1, 11).Value = "Description"
    Assetic_DisposedValuations.Cells(1, 12).Value = "Comments"



 For i = 10 To Cells.SpecialCells(xlCellTypeLastCell).Row
  
    If Len(Cells(i, Component_Col).Value) + Len(Cells(i, ID_Col).Value) > 0 And Cells(i, DisposalType_Col).Value = "Full Asset Disposal" Then 'Ignore if there are no components or ID or not Full Disposal
    
    'Populate Assetic Disposals
    
    
    Assetic_DisposedAssets.Cells(Target_Row, 1).Value = Cells(i, ID_Col).Value
    Assetic_DisposedAssets.Cells(Target_Row, 2).Value = "Disposed"
    'Assetic_DisposedAssets.Cells(Target_Row, 3).Value = ""
    'Assetic_DisposedAssets.Cells(Target_Row, 4).Value = ""
    Assetic_DisposedAssets.Cells(Target_Row, 5).Value = Cells(i, DisposeDate_Col).Value
    Assetic_DisposedAssets.Cells(Target_Row, 6).Value = Cells(i, Reason_Col).Value
    
    
    
    Assetic_DisposedValuations.Cells(Target_Row, 1).Value = Cells(i, ValRecID_Col).Value
    Assetic_DisposedValuations.Cells(Target_Row, 2).Value = Cells(i, ID_Col).Value
    Assetic_DisposedValuations.Cells(Target_Row, 3).Value = Trim(Cells(i, Component_Col).Value)
    Assetic_DisposedValuations.Cells(Target_Row, 4).Value = Cells(i, ValCompName_Col).Value
    Assetic_DisposedValuations.Cells(Target_Row, 5).Value = Cells(i, ValDate_Col).Value
    Assetic_DisposedValuations.Cells(Target_Row, 6).Value = Cells(i, ValRecType_Col).Value
    Assetic_DisposedValuations.Cells(Target_Row, 7).Value = "No"
    'Assetic_DisposedValuations.Cells(Target_Row, 8).Value = ""
    'Assetic_DisposedValuations.Cells(Target_Row, 9).Value = ""
    Assetic_DisposedValuations.Cells(Target_Row, 10).Value = PRCode
    Assetic_DisposedValuations.Cells(Target_Row, 11).Value = PRDesc
    Assetic_DisposedValuations.Cells(Target_Row, 12).Value = Cells(i, Comments_Col).Value
       
    
 


    
    Target_Row = Target_Row + 1
    
    End If
    
 
  Next i
  


End If



Application.StatusBar = False
'Application.ScreenUpdating = True

Sheets(Cur_Sht).Select
Assetic_DisposedAssets.Name = PRCode & "_Assetic_DisposedAssets"
Assetic_DisposedValuations.Name = PRCode & "_Assetic_DisposedVals"


End Sub

