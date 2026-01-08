Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addNewPartNumber_Click()
On Error GoTo Err_Handler

If DLookup("paramVal", "tblDBinfoBE", "parameter = 'allowNPNadditions'") = True Or Environ("username") = DLookup("message", "tblDBinfoBE", "parameter = 'allowNPNadditions'") Then
    DoCmd.OpenForm "frmNewPartNumber"
Else
    MsgBox "This system has been turned off. The new part number system is being replaced with the Insight Quote System. Please use that to create New Part Numbers", vbOKOnly, "Sorry about that"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub Reset_Labels()
On Error GoTo Err_Handler

Dim ctrl As Control

For Each ctrl In Me.Controls
    If TypeOf ctrl Is label Then
        ctrl.Caption = Replace(ctrl.Caption, ">", "-")
        ctrl.Caption = Replace(ctrl.Caption, "<", "-")
    End If
Next ctrl

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler

If validate Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate Then
    DoCmd.OpenReport "rptNewPart", acViewPreview, , "[newPartNumber]= " & Me.newPartNumber
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function updatePrefix()

ncmPrefix = Me.NCMcategory.column(1) & Left(Me.NCMsubCategory.column(2), 1)

End Function

Function validate()

validate = False

Dim errorMsg As String
errorMsg = ""

'check if NCM or not
If Me.partNumberType = 2 Then 'NCM
    If IsNull(Me.NCMcategory) Then errorMsg = "NCM category"
    If IsNull(Me.NCMsubCategory) Then errorMsg = "NCM sub-category"
End If

'BOTH
If IsNull(Me.PartDescription) Then errorMsg = "Part Description"
If IsNull(Me.customerId) Then errorMsg = "Customer"

If errorMsg <> "" Then
    MsgBox "Please fill out " & errorMsg, vbInformation, "Please fix"
    Exit Function
End If

validate = True

End Function

Private Sub Color_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub customerID_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub customerPartNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

If Me.Dirty Then
    If Not validate Then Me.Undo
End If

End Sub

Private Sub Form_Current()
Call partNumberType_AfterUpdate
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

If Nz(userData("org"), 0) = 5 Then Call toggleLanguage_Click 'Spanish for NCM

Me.OrderBy = "newPartNumber Desc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "tblSalesUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCreatedBy_Click()
On Error GoTo Err_Handler

Dim newLabel As String
newLabel = labelUpdate(Me.lblCreatedBy.Caption)

Reset_Labels
Me.lblCreatedBy.Caption = newLabel

Me.OrderBy = "[creator] " & labelDirection(newLabel)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCreatedBy_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
Dim X
X = InputBox("Enter Creator", "Search by Creator")
If X = "" Or X = vbCancel Then Exit Sub
    
Me.Form.filter = "[creator] Like '*" & X & "*'"
Me.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCreatedDate_Click()
On Error GoTo Err_Handler

Dim newLabel As String
newLabel = labelUpdate(Me.lblCreatedDate.Caption)

Reset_Labels
Me.lblCreatedDate.Caption = newLabel

Me.OrderBy = "[createdDate] " & labelDirection(newLabel)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNewPartDescription_Click()
On Error GoTo Err_Handler

Dim newLabel As String
newLabel = labelUpdate(Me.lblNewPartDescription.Caption)

Reset_Labels
Me.lblNewPartDescription.Caption = newLabel

Me.OrderBy = "[partDescription] " & labelDirection(newLabel)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNewPartDescription_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
Dim X
X = InputBox("Enter Part Description", "Search by Part Description")
If X = "" Or X = vbCancel Then Exit Sub
    
Me.Form.filter = "[partDescription] Like '*" & X & "*'"
Me.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNewPartNumber_Click()
On Error GoTo Err_Handler

Dim newLabel As String
newLabel = labelUpdate(Me.lblNewPartNumber.Caption)

Reset_Labels
Me.lblNewPartNumber.Caption = newLabel

Me.OrderBy = "[newPartNumber] " & labelDirection(newLabel)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNewPartNumber_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
Dim X
X = InputBox("Enter Part Number", "Filter by Part Number")
If X = "" Or X = vbCancel Then Exit Sub
    
Me.Form.filter = "[newPartNumber] = " & X
Me.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialType_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NCMcategory_AfterUpdate()
On Error GoTo Err_Handler

Me.NCMsubCategory = Null
Me.NCMsubCategory.RowSource = "SELECT recordid, NCMpnSubCategoryLetter, NCMpnSubCategory From tblDropDownsSP WHERE NCMpnSubCategory Is Not Null AND NCMpnSubCategoryLetter = '" & Me.NCMcategory.column(1) & "'"

Call updatePrefix
Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NCMsubCategory_AfterUpdate()
On Error GoTo Err_Handler

Call updatePrefix
Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newPartNumberHelp_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NJPpartNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub notes_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partDescription_AfterUpdate()
On Error GoTo Err_Handler

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "tblSalesUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = '" & Me.newPartNumber & "'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partNumberType_AfterUpdate()
On Error Resume Next

Dim ncm As Boolean

ncm = Me.partNumberType = 2
    
If ncm = False Then
    If Not IsNull(Me.NCMcategory) Then Me.NCMcategory = Null
    If Not IsNull(Me.NCMsubCategory) Then Me.NCMsubCategory = Null
End If

Me.lblNCMcat.Visible = ncm

Me.NCMcategory.Visible = ncm
Me.Command59.Visible = ncm
Me.lblNCMsubCat.Visible = ncm
Me.Command62.Visible = ncm

Me.NCMsubCategory.Visible = ncm
Me.ncmPrefix.Visible = ncm
Me.lblPrefix.Visible = ncm

Call updatePrefix

Call registerSalesUpdates("tblPartNumbers", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.newPartNumber)

End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub resetFilter_Click()
On Error GoTo Err_Handler

Me.FilterOn = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toggleLanguage_Click()
On Error GoTo Err_Handler

Dim lang(0 To 21, 0 To 1) As String 'first column is english, second is spanish

lang(0, 0) = "New Part Number"
lang(1, 0) = "Created"
lang(2, 0) = "by"
lang(3, 0) = "PN Type*"
lang(4, 0) = "NCM Category*"
lang(5, 0) = "NCM Sub-Category*"
lang(6, 0) = "Full Part Number"
lang(7, 0) = "Prefix"
lang(8, 0) = "Part Number"
lang(9, 0) = "Part Description*"
lang(10, 0) = "Customer*"
lang(11, 0) = "Customer Part #"
lang(12, 0) = "Material"
lang(13, 0) = "Color"
lang(14, 0) = "Nifco Global Part #"
lang(15, 0) = "Notes"
lang(16, 0) = " Save"
lang(17, 0) = " See All"
lang(18, 0) = "Print"
lang(19, 0) = "General Info"
lang(20, 0) = " New Part Number"
lang(21, 0) = "Part Description"

lang(0, 1) = "Sistema de Código Consecutivo"
lang(1, 1) = "Fecha de Creación"
lang(2, 1) = "Creado Por"
lang(3, 1) = "Los Sucursal*"
lang(4, 1) = "Categoría*"
lang(5, 1) = "Subcategoría*"
lang(6, 1) = "No. Parte Completo"
lang(7, 1) = "Prefijo"
lang(8, 1) = "No. Parte"
lang(9, 1) = "Descripcion*"
lang(10, 1) = "Cliente*"
lang(11, 1) = "No. Parte del Cliente"
lang(12, 1) = "Resina"
lang(13, 1) = "Color"
lang(14, 1) = "No. Parte Global"
lang(15, 1) = "Notas"
lang(16, 1) = " Guardar"
lang(17, 1) = " Ver todo"
lang(18, 1) = "Imprimir"
lang(19, 1) = "Información General"
lang(20, 1) = " Agregar Nuevo"
lang(21, 1) = "Descripcion"

Dim langMark As Long
If Me.toggleLanguage.Caption = "English" Then
    Me.toggleLanguage.Caption = "Español"
    langMark = 0
Else
    Me.toggleLanguage.Caption = "English"
    langMark = 1
End If

'Me.lblTitleBar.Caption = lang(0, langMark)
Me.lblPNtype.Caption = lang(3, langMark)
Me.lblNCMcat.Caption = lang(4, langMark)
Me.lblNCMsubCat.Caption = lang(5, langMark)
Me.lblFullPartNumber.Caption = lang(6, langMark)
Me.lblPrefix.Caption = lang(7, langMark)
Me.lblPartNumber.Caption = lang(8, langMark)
Me.lblDescription.Caption = lang(9, langMark)
Me.lblCustomer.Caption = lang(10, langMark)
Me.lblCustomerPN.Caption = lang(11, langMark)
Me.lblMaterial.Caption = lang(12, langMark)
Me.lblColor.Caption = lang(13, langMark)
Me.lblNJP.Caption = lang(14, langMark)
Me.lblNotes.Caption = lang(15, langMark)
Me.resetFilter.Caption = lang(17, langMark)
Me.cmdPrint.Caption = lang(18, langMark)
Me.lblGeneralInfo.Caption = lang(19, langMark)
Me.addNewPartNumber.Caption = lang(20, langMark)

Dim lblClicker As String, lblLen, newCap As String
lblClicker = Me.lblNewPartNumber.Caption
lblLen = Len(lblClicker)
newCap = Replace(lblClicker, Left(lblClicker, lblLen - 2), lang(8, langMark))
Me.lblNewPartNumber.Caption = newCap

lblClicker = Me.lblNewPartDescription.Caption
lblLen = Len(lblClicker)
newCap = Replace(lblClicker, Left(lblClicker, lblLen - 2), lang(21, langMark))
Me.lblNewPartDescription.Caption = newCap

lblClicker = Me.lblCreatedBy.Caption
lblLen = Len(lblClicker)
newCap = Replace(lblClicker, Left(lblClicker, lblLen - 2), lang(2, langMark))
Me.lblCreatedBy.Caption = newCap

lblClicker = Me.lblCreatedDate.Caption
lblLen = Len(lblClicker)
newCap = Replace(lblClicker, Left(lblClicker, lblLen - 2), lang(1, langMark))
Me.lblCreatedDate.Caption = newCap

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.toggleLanguage.name, Err.DESCRIPTION, Err.number)
End Sub
