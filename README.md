# CobroPeaje
Macro para insertar un número de placa, genera un ticket que permite imprimir y por último almacena cada transacción, totalizando por día.

Así luce esta sencilla aplicación:

![image](https://github.com/vjceballosj/CobroPeaje/assets/108242764/b0d66102-065e-4afa-8063-2b0788525d95)

Para utilizar esta macro, puede hacerlo manual siguiendo estos pasos:

1. Abre un archivo de Excel.
2. Debe crear 4 hojas:
   "CobroPeaje":
   //Hoja que muestra el ticket que se genera y el botón que permite buscar un número de placa:
   
   ![Sheet_CobroPeaje](https://github.com/vjceballosj/CobroPeaje/assets/108242764/3edad160-69be-48cf-b5cb-29b9eade3d86)

   "plates":
   //Hoja que muestra la BBDD de placas ya predeterminadas para saber la categría a cobrar:
   
   ![Sheet_plates](https://github.com/vjceballosj/CobroPeaje/assets/108242764/e408741f-bb89-43a2-b38e-319777690cfd)

   "categories":
   
   //Hoja que muestra el valor a cobrar por cada categoría, en este caso hay 7:
   ![Sheet_categories](https://github.com/vjceballosj/CobroPeaje/assets/108242764/67ad5c4e-ca76-4dae-bc56-b7f8deb46cde)

   "totales":
   
   //Hoja que muestra el cobrado en peajes por día calendario:
    ![Sheet_totales](https://github.com/vjceballosj/CobroPeaje/assets/108242764/d0f2f970-76c7-4867-a94e-96e412d36e3b)

4. Presiona Alt + F11 para abrir el Editor de Visual Basic.
5. Haz clic con el botón derecho del ratón en el nombre de tu archivo en el panel de la izquierda
   y selecciona "Insertar" -> "Módulo". Se creará un nuevo módulo en tu archivo.
6. Copia y pega el siguiente código de la macro en el módulo abierto:
//****************************************************************************************************************//
Dim plateFound As Boolean ' Variable de control

Sub searchplate()

    Dim categories As Worksheet
    Dim plates As Worksheet
    Dim totales As Worksheet
    Dim cobroPeaje As Worksheet
    Dim tbl1 As ListObject
    Dim tbl2 As ListObject
    Dim foundCell As Range
    Dim valueColumn As Range
    Dim valuetosearch As String
    Dim cat As String
    Dim found As Boolean
    Dim i As Integer
    Dim cobro As Double
    Dim totalCobros As Double
    Dim value As Double
    Dim fecha As Date

    ' Pide al usuario que ingrese el número de placa
    valuetosearch = InputBox("Ingrese el número de placa:")
    
    ' Reinicia la variable de control
    plateFound = False
    
    ' Obtiene las hojas "plates", "categories" y "totales"
    Set plates = ThisWorkbook.Sheets("plates")
    Set categories = ThisWorkbook.Sheets("categories")
    Set totales = ThisWorkbook.Sheets("totales")
    Set cobroPeaje = ThisWorkbook.Sheets("CobroPeaje")
    
    ' Busca la tabla en la hoja "categories"
        On Error Resume Next
            Set tbl1 = categories.ListObjects(1)
        On Error GoTo 0
        
    ' Busca las tablas "category" en la hoja "plates"
    For i = 1 To 7
        On Error Resume Next
            Set tbl2 = plates.ListObjects(i)
        On Error GoTo 0
    
        ' Si encontró la tabla, busca la placa en la columna plate de la tabla encontrada
        If Not tbl2 Is Nothing Then
            Set foundCell = tbl2.ListColumns("plate cat. " & i).DataBodyRange.Find(What:=valuetosearch, LookIn:=xlValues, LookAt:=xlWhole)
            
                ' Si encuentra la placa, obtiene tanto el valor correspondiente de la columna "value"
                ' como la categoría a cobrar de la hoja "categories"
                If Not foundCell Is Nothing Then
                    Set valueColumn = categories.Range("categories[cat]").Find(What:=tbl2.Name, LookIn:=xlValues, LookAt:=xlWhole)
                    value = valueColumn.Offset(0, 1).value
                    found = True
                
                        ' Obtener la fecha y el valor del cobro actual de la hoja "cobroPeaje"
                        fecha = cobroPeaje.Range("B6").value
                        cobro = value
        
                        ' Buscar la fila correspondiente a la fecha en la hoja "totales"
                        Dim fechaColumn As Range
                        Set fechaColumn = totales.Columns("A").Find(What:=fecha, LookIn:=xlValues, LookAt:=xlWhole)
        
                        ' Si la fecha no existe en la hoja "totales", agregarla y asignar el valor de cobro
                        If fechaColumn Is Nothing Then
                            Dim lastTotalRow As Long
                            lastTotalRow = totales.Cells(totales.Rows.Count, "A").End(xlUp).Row
                            totales.Cells(lastTotalRow + 1, "A").value = fecha
                            totales.Cells(lastTotalRow + 1, "B").value = cobro
                        Else
                        
                        ' Si la fecha ya existe, sumar el valor de cobro al total existente
                            totalCobros = totales.Cells(fechaColumn.Row, "B").value
                            totales.Cells(fechaColumn.Row, "B").value = totalCobros + cobro
                        End If
                        
                    ' Genera el ticket de factura en la hoja llamada "CobroPeaje"
                    Set cobroPeaje = ThisWorkbook.Sheets("CobroPeaje")
                    Dim valorconsecutivo As Long
                    
                    ' Configura el contenido del ticket
                    valorconsecutivo = cobroPeaje.Range("B4")
                    cobroPeaje.Range("B4").value = valorconsecutivo + 1
                    cobroPeaje.Range("A6").value = "Fecha:"
                    cobroPeaje.Range("B6").value = Format(Now, "dd/mm/yyyy")
                    cobroPeaje.Range("A7").value = "Hora:"
                    cobroPeaje.Range("B7").value = Format(Now, "hh:mm:ss")
                    cobroPeaje.Range("A8").value = "Número de placa:"
                    cobroPeaje.Range("B8").value = valuetosearch
                    cobroPeaje.Range("A9").value = "Valor a pagar:"
                    cobroPeaje.Range("B9").value = value
            
                    ' Da formato al ticket
                    With cobroPeaje.Range("A6:A8")
                        .HorizontalAlignment = xlLeft
                        .Font.Bold = True
                    End With
                    With cobroPeaje.Range("B6:B8")
                        .HorizontalAlignment = xlRight
                        .Font.Bold = False
                    End With
        
                    ' Ajusta el ancho de las columnas
                    cobroPeaje.Columns("A:B").ColumnWidth = 19
        
                    ' Imprime el ticket
                    If MsgBox("¿Desea imprimir el Ticket?", vbQuestion + vbYesNo) = vbYes Then
                    cobroPeaje.PrintOut
                    End If
                    Exit Sub
                End If
            End If
    Next i
        
    ' Si no se encuentra la placa en ninguna tabla, muestra un cuadro de diálogo y permite buscar otra placa
    If Not found Then
        MsgBox "Placa no encontrada"
        If MsgBox("¿Desea buscar otra placa?", vbQuestion + vbYesNo) = vbYes Then
            searchplate
        End If
        Exit Sub
    End If
    
    ' Actualiza la variable de control
    plateFound = True
    
End Sub

Sub EjecutarBuscarPlaca()
    ' Verifica si la variable de control es verdadera (se encontró una placa)
    ' Si es verdadera, muestra un mensaje y sale de la macro
    ' Si no es verdadera, ejecuta la macro searchplate
    If plateFound Then
        MsgBox "Ya se ha encontrado una placa. Reinicie la macro para buscar otra placa."
    Else
        searchplate
    End If
End Sub
//****************************************************************************************************************//

7. Guarda y cierra el Editor de Visual Basic.
8. En tu archivo de Excel, selecciona una celda o un botón donde quieras asignar esta macro.
9. Ve a la pestaña "Desarrollador" (si no la tienes visible, debes habilitarla en las opciones de Excel).
10. Haz clic en "Macros" y selecciona "searchplate".
11. Haz clic en "Ejecutar".
