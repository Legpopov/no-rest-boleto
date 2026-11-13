VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPlanningRoutingEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Application

Private Sub App_AfterCalculate()
    RoutingClearCache
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    RoutingClearCache
End Sub
