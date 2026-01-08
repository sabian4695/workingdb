Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

Dim val As String
val = validate
If val <> "" Then
    MsgBox "Please fill out " & val, vbInformation, "Please fix"
    Exit Sub
End If

DoCmd.CLOSE

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
Dim val As String
val = validate
If val <> "" Then
    MsgBox "Please fill out " & val, vbInformation, "Please fix"
    Exit Sub
End If

DoCmd.OpenReport "rptNewPart", acViewPreview, , "[newPartNumber]= " & Me.newPartNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler

Dim val As String
val = validate
If val <> "" Then
    MsgBox "Please fill out " & val, vbInformation, "Please fix"
    Exit Sub
End If

DoCmd.GoToRecord , , acNewRec
Me.newPartNumber = Nz(DMax("[newPartNumber]", "[tblPartNumbers]"), 0) + 1
Me.creator = Environ("username")
Me.Repaint

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.newPartNumber = Nz(DMax("[newPartNumber]", "[tblPartNumbers]"), 0) + 1
Me.creator = Environ("username")
If Me.Dirty Then Me.Dirty = False

Me.Repaint

If Nz(userData("org"), "") = 5 Then 'ORG = NCM
    partNumberType = 2
    Call toggleLanguage_Click
Else
    Me.partNumberType = 1
End If

Call partNumberType_AfterUpdate

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function updatePrefix()

ncmPrefix = Me.NCMcategory.column(1) & Left(Me.NCMsubCategory.column(2), 1)

End Function

Function validate() As String
validate = ""

'check if NCM or not
If Me.partNumberType = 2 Then 'NCM
    If IsNull(Me.NCMcategory) Then validate = "NCM category"
    If IsNull(Me.NCMsubCategory) Then validate = "NCM sub-category"
End If

'BOTH
If IsNull(Me.PartDescription) Then validate = "Part Description"
If IsNull(Me.customerId) Then validate = "Customer"

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If validate <> "" Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.ID) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If


Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub NCMcategory_AfterUpdate()
On Error GoTo Err_Handler

Me.NCMsubCategory = Null
Me.NCMsubCategory.RowSource = "SELECT recordid, NCMpnSubCategoryLetter, NCMpnSubCategory From tblDropDownsSP WHERE NCMpnSubCategory Is Not Null AND NCMpnSubCategoryLetter = '" & Me.NCMcategory.column(1) & "'"

Call updatePrefix

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NCMsubCategory_AfterUpdate()
On Error GoTo Err_Handler

Call updatePrefix

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

Private Sub partNumberType_AfterUpdate()
On Error GoTo Err_Handler

Dim ncm As Boolean

ncm = Me.partNumberType = 2

Me.lblNCMcat.Visible = ncm
Me.NCMcategory.Visible = ncm
Me.NCMcategory = Null
Me.Command59.Visible = ncm
Me.lblNCMsubCat.Visible = ncm
Me.Command62.Visible = ncm
Me.NCMsubCategory.Visible = ncm
Me.NCMsubCategory = Null
Me.ncmPrefix.Visible = ncm
Me.lblPrefix.Visible = ncm

Call updatePrefix

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toggleLanguage_Click()
On Error GoTo Err_Handler

Dim lang(0 To 19, 0 To 1) As String 'first column is english, second is spanish

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
lang(16, 0) = " Save + Close"
lang(17, 0) = "Save / Add New"
lang(18, 0) = "Print"
lang(19, 0) = "General Info"

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
lang(16, 1) = " Guardar y Cerrar"
lang(17, 1) = "Guardar y Agregar Buevo"
lang(18, 1) = "Imprimir"
lang(19, 1) = "Información General"

Dim langMark As Integer
If Me.toggleLanguage.Caption = "English" Then
    Me.toggleLanguage.Caption = "Español"
    langMark = 0
Else
    Me.toggleLanguage.Caption = "English"
    langMark = 1
End If

'Me.lblTitleBar.Caption = lang(0, langMark)
Me.lblCreated.Caption = lang(1, langMark)
Me.lblBy.Caption = lang(2, langMark)
Me.lblPNtype.Caption = lang(3, langMark)
Me.lblNCMcat.Caption = lang(4, langMark)
Me.lblNCMsubCat.Caption = lang(5, langMark)
Me.lblFullPartNumber.Caption = lang(6, langMark)
Me.lblPrefix.Caption = lang(7, langMark)
Me.lblPartNumber.Caption = lang(8, langMark)
Me.lblPartDescription.Caption = lang(9, langMark)
Me.lblCustomer.Caption = lang(10, langMark)
Me.lblCustomerPN.Caption = lang(11, langMark)
Me.lblMaterial.Caption = lang(12, langMark)
Me.lblColor.Caption = lang(13, langMark)
Me.lblNJP.Caption = lang(14, langMark)
Me.lblNotes.Caption = lang(15, langMark)
Me.cmdSave.Caption = lang(16, langMark)
Me.cmdAdd.Caption = lang(17, langMark)
Me.cmdPrint.Caption = lang(18, langMark)
Me.lblGeneralInfo.Caption = lang(19, langMark)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.toggleLanguage.name, Err.DESCRIPTION, Err.number)
End Sub
