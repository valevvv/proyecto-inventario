Attribute VB_Name = "Procedimiento"
Global con As New ADODB.Connection
Global rsropa As New ADODB.Recordset
Global rsinventario As New ADODB.Recordset
Sub main()
    With con
        .CursorLocation = adUseClient 'Vamos a ser clientes de la base de datos
        'Conexion a la base de datos
        .Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & App.Path & "\Inventario.accdb;Persist Security Info=False"
        frmsplash.Show
    End With
    
End Sub

Sub tablaAutor()
    With rsropa
        
        If .State = 1 Then .Close
            .Source = "Ropa"
            .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
            .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            .Open "select * from Ropa", con
    End With
End Sub
Sub tablaInventario()
    With rsinventario
        
        If .State = 1 Then .Close
            .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
            .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            .Open "select * from inventario", con
            .MoveFirst
    End With
    
End Sub

