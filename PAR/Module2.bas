Attribute VB_Name = "Module2"
'IPS Project Asset Register
'Developed by Amer Acosta
'8-November-2018

Sub View_Summary()

Sht_Summary.Select

End Sub


Sub View_AsCons()

Sht_AsCons.Select

End Sub

Sub View_Project_WideCosts()

Sht_ProjectWide.Select

End Sub


Sub View_HandoverCost()

SHt_HandoverCost.Select

End Sub

Sub View_NewAssets()

Sht_New.Select

End Sub

Sub View_RenewedAssets()

Sht_Renew.Select

End Sub

Sub View_disposedAssets()

Sht_Dispose.Select

End Sub


Sub Hide_Lookup()

Sht_Lookup_AssetClass.Visible = xlSheetVeryHidden
Sht_Lookup_CostCategory.Visible = xlSheetVeryHidden
Sht_Lookup_CostItem.Visible = xlSheetVeryHidden
Sht_CorpOH.Visible = xlSheetVeryHidden
Sht_AHClass.Visible = xlSheetVeryHidden
Sht_AHSubClass.Visible = xlSheetVeryHidden
Sht_AHType.Visible = xlSheetVeryHidden
Sht_AHSubType.Visible = xlSheetVeryHidden
Sht_CoASchema.Visible = xlSheetVeryHidden
Sht_SCComponent.Visible = xlSheetVeryHidden
Sht_SCFinancials.Visible = xlSheetVeryHidden
Sht_UoM.Visible = xlSheetVeryHidden
Sht_TreatmentType.Visible = xlSheetVeryHidden

End Sub

Sub UnHide_Lookup()

Sht_Lookup_AssetClass.Visible = xlSheetVisible
Sht_Lookup_CostCategory.Visible = xlSheetVisible
Sht_Lookup_CostItem.Visible = xlSheetVisible
Sht_CorpOH.Visible = xlSheetVisible
Sht_AHClass.Visible = xlSheetVisible
Sht_AHSubClass.Visible = xlSheetVisible
Sht_AHType.Visible = xlSheetVisible
Sht_AHSubType.Visible = xlSheetVisible
Sht_CoASchema.Visible = xlSheetVisible
Sht_SCComponent.Visible = xlSheetVisible
Sht_SCFinancials.Visible = xlSheetVisible
Sht_UoM.Visible = xlSheetVisible
Sht_TreatmentType.Visible = xlSheetVisible

End Sub


Sub View_Transactions()

Sht_Transactions.Select

End Sub
