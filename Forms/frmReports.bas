Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub canlLabWObyPN_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLabWO_Search"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub costDocuments_Click()
On Error GoTo Err_Handler

Dim mainPath As String, adFilter As String
mainPath = mainFolder(Me.ActiveControl.name)

If Nz(Form_DASHBOARD.partNumberSearch, "") = "" Then
    adFilter = ""
Else
    adFilter = "?FilterField1=Part_x0020_Number&FilterValue1=" & Form_DASHBOARD.partNumberSearch
End If

openPath (mainPath & adFilter)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cstmXref_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmCustomerXref", , , "[NAM] = '" & Form_DASHBOARD.partNumberSearch & "'"
Form_frmCustomerXref.NAMsrchBox = Form_DASHBOARD.partNumberSearch
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub forecastOrders_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmPartForecast"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Item_Search_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
DoCmd.OpenForm "frmItemSearch"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub itemCategories_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)

Dim filterVal
filterVal = "[PN] = '" & Form_DASHBOARD.partNumberSearch & "'"

DoCmd.OpenForm "frmItemCategories", , , filterVal
Form_frmItemCategories.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub itemCost_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)

DoCmd.OpenForm "frmPartCost", , , "[ITEM_NUMBER] = '" & Nz(Form_DASHBOARD.partNumberSearch) & "'"
Form_frmPartCost.NAMsrchBox = Form_DASHBOARD.partNumberSearch

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub matSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
DoCmd.OpenForm "frmMaterialSearch"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NPIF_Click()
On Error GoTo Err_Handler

Dim mainPath As String, adFilter As String
mainPath = mainFolder(Me.ActiveControl.name)

If Nz(Form_DASHBOARD.partNumberSearch, "") = "" Then
    adFilter = ""
Else
    adFilter = "?FilterField1=Part%5Fx0020%5FNumber&FilterValue1=" & Form_DASHBOARD.partNumberSearch
End If

openPath (mainPath & adFilter)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openITRs_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmITRreport"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openOrders_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmPartOpenOrders"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub poSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmPOsearch", , , "capNum = '2024-K124'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub qtyOnHand_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)

Dim filterVal, pNum
filterVal = idNAM(Nz(Form_DASHBOARD.partNumberSearch), "NAM")
If filterVal = "" Then filterVal = "44388"
filterVal = "[INVENTORY_ITEM_ID] = " & filterVal

DoCmd.OpenForm "frmOnHandQty", , , filterVal
Form_frmOnHandQty.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub routingToolSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmToolPartNumRoutings", , , "SEGMENT1 = '" & Nz(Form_DASHBOARD.partNumberSearch) & "'"
Form_frmToolPartNumRoutings.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sifItemReport_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmPartSalesList"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub SIFsearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmSIF"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub SIFsearchPN_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmSIFpartHistory", , , "Nifco_Part_Number = '" & Nz(Form_DASHBOARD.partNumberSearch) & "'"
Form_frmSIFpartHistory.srchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srchBOM_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
Dim filterVal, pNum
filterVal = idNAM(Nz(Form_DASHBOARD.partNumberSearch), "NAM")
If Nz(filterVal) = "" Then filterVal = "3147035"
filterVal = "[ASSEMBLY_ITEM_ID] = " & filterVal

DoCmd.OpenForm "frmBOMsearch", , , filterVal
Form_frmBOMsearch.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srchECO_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmECOs", , , "[Change_Notice] = 'CNL10000'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srchMaterials_Click()
On Error GoTo Err_Handler
DoCmd.OpenForm "frmCstItemCostRpt"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cnlLabWOs_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmLabWOs"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub ECOhistory_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmECOpartHistory"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub slbToolNotes_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmSLBtoolNotes"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolingInfo_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)
DoCmd.OpenForm "frmToolingInfo"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialSearch_Click()
On Error GoTo Err_Handler

Call logClick(Me.ActiveControl.name, Me.name, Form_DASHBOARD.partNumberSearch)

Dim filterVal, pNum
filterVal = Nz(Form_DASHBOARD.partNumberSearch)
If filterVal = "" Then filterVal = "29120"
filterVal = "[PRIMARY_ITEM_ID] = " & idNAM(filterVal, "NAM")

DoCmd.OpenForm "frmTrialJobs", , , filterVal
Form_frmTrialJobs.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
