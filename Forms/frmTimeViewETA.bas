Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub lblBetween_Click()
    Me.betweenDates.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblOverdue_Click()
    Me.overdueETA.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblUser_Click()
    Me.userName.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub
