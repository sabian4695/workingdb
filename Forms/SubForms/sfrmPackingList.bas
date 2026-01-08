Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub item_AfterUpdate()
On Error GoTo Err_Handler

Me.itemDescription = ""
Me.materialType = "Plastic"
Me.countryOfOrigin = "United States"
Me.unitWeight = 0
Me.weightUnit = "Lb"
Me.unitCost = 0

On Error Resume Next
Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset

Set rs1 = db.OpenRecordset("SELECT * FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & Me.ITEM & "'", dbOpenSnapshot)
Me.itemDescription = rs1("DESCRIPTION")
Select Case rs1!ITEM_TYPE
    Case "COMPTM", "FGM"
        Me.materialType = "Plastic"
    Case "EQ"
        Me.materialType = "Metal"
        Me.countryOfOrigin = "China"
End Select

rs1.FindFirst "UNIT_WEIGHT is not null"
Me.unitWeight = rs1!UNIT_WEIGHT
Me.weightUnit = rs1!WEIGHT_UOM_CODE

rs1.CLOSE
Set rs1 = Nothing

Dim rs2 As Recordset
Set rs2 = db.OpenRecordset("SELECT ITEM_COST FROM APPS_CST_ITEM_COST_TYPE_V WHERE ITEM_NUMBER = '" & Me.ITEM & "' AND COST_TYPE = 'Frozen' AND ITEM_COST > 0")
Me.unitCost = CDbl(rs2!ITEM_COST)

rs2.CLOSE
Set rs2 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
