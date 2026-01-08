Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdPrint_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
DoCmd.OpenReport "rptPackingList", acViewPreview, , "tblPackList.recordId = " & Me.recordId ', acHidden
DoCmd.RunCommand acCmdPrint

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

DoCmd.GoToRecord , , acNewRec

Me.packedBy = getFullName()
Me.sendTrackingTo = getEmail(Environ("username"))

Select Case userData("Org")
    Case 1 'CNL
        Me.senderAddress = "8015 Dove Pkwy, Canal Winchester, OH 43110"
    Case 2 'SLB
        Me.senderAddress = "300 Hudson Blvd, Shelbyville, KY 40065"
    Case 3 'LVG
        Me.senderAddress = "130 Wheeler St, La Vergne, TN 37086"
    Case 4 'CUU
        Me.senderAddress = "AV. Nicolas Gogal #11301, Chihuahua Industrial, Chihuahua, Chihuahua Mexico 31125"
    Case 5 'NCM
        Me.senderAddress = "Parque Industrial Castro del Rio, Avenida Rio San Lorenzo #1117, Barrio de San Vicente, 36815 Irapuato, Gto., Mexico"
End Select

Exit Sub
Err_Handler:: Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number): End Sub

Function validate() As Boolean

validate = False

'Select Case True
'    Case Nz(Me.modelCode) = ""
'        MsgBox "Please enter a model code.", vbOKOnly, "Warning"
'        Exit Function
'    Case Nz(Me.modelYear) = ""
'        MsgBox "Please enter a model year.", vbOKOnly, "Warning"
'        Exit Function
'    Case Nz(Me.OEM) = ""
'        MsgBox "Please select an OEM.", vbOKOnly, "Warning"
'        Exit Function
'    Case Nz(Me.modelName) = ""
'        MsgBox "Please enter a model name.", vbOKOnly, "Warning"
'        Exit Function
'    Case Len(Me.modelYear) <> 4
'        MsgBox "Please adjust your model year to be a 4 digit number", vbOKOnly, "Warning"
'        Exit Function
'End Select

validate = True

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If validate = False Then Exit Sub

If Me.Dirty Then Me.Dirty = False

Call registerWdbUpdates("tblPackList", Me.recordId, "Pack List", "", "Created")

DoCmd.CLOSE

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub packedBy_AfterUpdate()
On Error GoTo Err_Handler

On Error Resume Next
Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT * FROM tblPermissions WHERE firstName & ' ' & lastName = '" & Me.packedBy & "'")

Me.sendTrackingTo = rs1!userEmail

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub prevAddress_AfterUpdate()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT * FROM tblPackList WHERE recordId = " & Me.prevAddress)

Me.fullName = rs1!fullName
Me.addressLine1 = rs1!addressLine1
Me.addressLine2 = rs1!addressLine2
Me.city = rs1!city
Me.stateOrProvince = rs1!stateOrProvince
Me.zipOrPostalCode = rs1!zipOrPostalCode
Me.countryOrRegion = rs1!countryOrRegion

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
