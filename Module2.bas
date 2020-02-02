Attribute VB_Name = "Module2"
Sub main()
With base
    .CursorLocation = adUseClient
    .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Jack 3D;Data Source=RAMIRO-B7ED4DBC\SQLEXPRESS"
    FrmPrincipal.Show
End With
End Sub
Sub AbrirTablaCliente()
With RsCliente
    If .State = 1 Then .Close
    .Open "select * from Cliente", base, adOpenStatic, adLockOptimistic
    
End With
End Sub

Sub AbrirTablaFabricacion()
With RsFabricacion
    If .State = 1 Then .Close
    .Open "select * from Fabricacion", base, adOpenStatic, adLockOptimistic
End With
End Sub


Sub AbrirTablaConstantes()
With RsConstantes
    If .State = 1 Then .Close
    .Open "select * from Constantes", base, adOpenStatic, adLockOptimistic
End With
End Sub


