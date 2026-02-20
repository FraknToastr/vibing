Attribute VB_Name = "Module7"
Sub CurSheet_To_CSV()



Spath = "C:\Assetic_Extract"



    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    If fdObj.FolderExists(Spath) Then
       
    Else
        fdObj.CreateFolder (Spath)
        
    End If
    
    Set fdObj = Nothing
 



If Dir(Spath, vbDirectory) = "" Then
Shell ("cmd /c mkdir """ & Path & """")
End If


ThisWorkbook.Activate

Cur_Sht = ActiveSheet.Name



Dim wbkExport As Workbook
Dim shtToExport As Worksheet

Set shtToExport = ThisWorkbook.Worksheets(Cur_Sht)     'Sheet to export as CSV
Set wbkExport = Application.Workbooks.Add
shtToExport.Copy Before:=wbkExport.Worksheets(wbkExport.Worksheets.Count)
Application.DisplayAlerts = False                       'Possibly overwrite without asking
wbkExport.SaveAs Filename:=Spath & "\" & Cur_Sht & ".csv", FileFormat:=xlCSV


Application.DisplayAlerts = True
wbkExport.Close SaveChanges:=False


MsgBox ("The generated Assetic extract can be found in " & Spath & "\" & Cur_Sht & ".csv")



End Sub
