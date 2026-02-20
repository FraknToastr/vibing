Attribute VB_Name = "Module6"
Sub Generate_Cost_and_Assetic()


Application.StatusBar = True




Call Populate_Handover_Cost


Assetic_NewAssets.Visible = xlSheetVisible
Assetic_NewComponent.Visible = xlSheetVisible
Assetic_NewNetworkMeasure.Visible = xlSheetVisible
Assetic_NewValuations.Visible = xlSheetVisible
Assetic_DisposedAssets.Visible = xlSheetVisible
Assetic_DisposedValuations.Visible = xlSheetVisible
Assetic_CapExRenewals.Visible = xlSheetVisible


Call Populate_AsseticNewAssets
Call Populate_AsseticRenewedAssets
Call Populate_AsseticDisposedAssets

Assetic_NewAssets.Visible = xlSheetHidden
Assetic_NewComponent.Visible = xlSheetHidden
Assetic_NewNetworkMeasure.Visible = xlSheetHidden
Assetic_NewValuations.Visible = xlSheetHidden
Assetic_DisposedAssets.Visible = xlSheetHidden
Assetic_DisposedValuations.Visible = xlSheetHidden
Assetic_CapExRenewals.Visible = xlSheetHidden



Application.ScreenUpdating = True

MsgBox ("Cost of project details has been generated!")


Application.StatusBar = False



End Sub
