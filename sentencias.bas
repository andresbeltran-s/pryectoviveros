Attribute VB_Name = "sentencias"
Option Explicit
Sub main()
With BASE
.CursorLocation = adUseClient
.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\basevivero.mdb;Persist Security Info=False"
principal.Show
End With
End Sub

Sub admin()
With rsadmin
If .State = 1 Then .Close
    .Open "select * from ADMINISTRADOR", BASE, adOpenStatic, adLockOptimistic
End With
End Sub

Sub plan()
With rsp

If .State = 1 Then .Close

    .Open "select * from PLANTA", BASE, adOpenStatic, adLockOptimistic
End With
End Sub
Sub detallefactura()
With detallefac

If .State = 1 Then .Close
    .Open "select * from DETALLE_PEDIDO", BASE, adOpenStatic, adLockOptimistic
End With
End Sub


