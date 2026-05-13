'========================================
'SISTEMA DE GESTIÓN DE HABITACIONES - GUI
'Versión completa con Interfaz Gráfica
'========================================

Option Explicit

' ========== VARIABLES GLOBALES ==========
Public wsHabitaciones As Worksheet
Public wsInventario As Worksheet
Public wsObservaciones As Worksheet
Public wsReportes As Worksheet
Public wsHistorial As Worksheet
Public wsDashboard As Worksheet

Public Const NUM_HABITACIONES = 30
Public Const NUM_ITEMS = 3

' ========== INICIALIZACIÓN ==========

Sub Auto_Open()
    'Se ejecuta automáticamente al abrir el libro
    On Error Resume Next
    
    Call InicializarHojas
    Call CrearDashboard
    Call CrearEstructuraBase
    
End Sub

Sub InicializarHojas()
    'Inicializa referencias a las hojas
    On Error Resume Next
    
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsHabitaciones = ThisWorkbook.Sheets("Habitaciones")
    Set wsInventario = ThisWorkbook.Sheets("Inventario")
    Set wsObservaciones = ThisWorkbook.Sheets("Observaciones")
    Set wsReportes = ThisWorkbook.Sheets("Reportes")
    Set wsHistorial = ThisWorkbook.Sheets("Historial")
    
End Sub

Sub CrearDashboard()
    'Crea la interfaz principal del Dashboard
    Dim btnAgregar As Shape
    Dim btnEditar As Shape
    Dim btnActualizar As Shape
    Dim btnBuscar As Shape
    Dim btnDarDeBaja As Shape
    Dim btnObservacion As Shape
    Dim btnCompletarObs As Shape
    Dim btnReporte As Shape
    Dim btnImprimir As Shape
    Dim btnHistorial As Shape
    
    On Error Resume Next
    
    With wsDashboard
        .Cells.Clear
        
        ' Título Principal
        .Range("A1:H1").Merge
        .Range("A1").Value = "🏨 SISTEMA DE GESTIÓN DE HABITACIONES"
        With .Range("A1")
            .Font.Bold = True
            .Font.Size = 18
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 102, 204)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Row(1).RowHeight = 35
        
        ' Subtítulo
        .Range("A2:H2").Merge
        .Range("A2").Value = "Control de 30 Habitaciones | Inventario | Mantenimiento"
        With .Range("A2")
            .Font.Size = 11
            .Font.Italic = True
            .Interior.Color = RGB(220, 230, 241)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Sección 1: GESTIÓN DE HABITACIONES
        .Range("A4:H4").Merge
        .Range("A4").Value = "📋 GESTIÓN DE HABITACIONES"
        With .Range("A4")
            .Font.Bold = True
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(192, 0, 0)
            .HorizontalAlignment = xlLeft
        End With
        
        ' Botones de Habitaciones
        Set btnAgregar = .Shapes.AddShape(msoShapeRoundedRectangle, 20, 140, 150, 35)
        With btnAgregar
            .Name = "btnAgregarHab"
            .TextFrame.Characters.Text = "AGREGAR HABITACIÓN"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(0, 128, 0)
        End With
        
        Set btnEditar = .Shapes.AddShape(msoShapeRoundedRectangle, 180, 140, 150, 35)
        With btnEditar
            .Name = "btnEditarHab"
            .TextFrame.Characters.Text = "EDITAR HABITACIÓN"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(0, 128, 0)
        End With
        
        Set btnActualizar = .Shapes.AddShape(msoShapeRoundedRectangle, 340, 140, 150, 35)
        With btnActualizar
            .Name = "btnActualizarItems"
            .TextFrame.Characters.Text = "ACTUALIZAR ITEMS"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(0, 128, 0)
        End With
        
        Set btnBuscar = .Shapes.AddShape(msoShapeRoundedRectangle, 500, 140, 150, 35)
        With btnBuscar
            .Name = "btnBuscarHab"
            .TextFrame.Characters.Text = "BUSCAR HABITACIÓN"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(0, 128, 0)
        End With
        
        ' Sección 2: INVENTARIO
        .Range("A8:H8").Merge
        .Range("A8").Value = "📦 INVENTARIO DE ALMACÉN"
        With .Range("A8")
            .Font.Bold = True
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 153, 76)
            .HorizontalAlignment = xlLeft
        End With
        
        Set btnDarDeBaja = .Shapes.AddShape(msoShapeRoundedRectangle, 20, 244, 150, 35)
        With btnDarDeBaja
            .Name = "btnDarDeBaja"
            .TextFrame.Characters.Text = "DAR DE BAJA ITEMS"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(204, 102, 0)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(153, 76, 0)
        End With
        
        ' Sección 3: OBSERVACIONES
        .Range("A12:H12").Merge
        .Range("A12").Value = "⚠️ OBSERVACIONES DE MANTENIMIENTO"
        With .Range("A12")
            .Font.Bold = True
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(153, 76, 0)
            .HorizontalAlignment = xlLeft
        End With
        
        Set btnObservacion = .Shapes.AddShape(msoShapeRoundedRectangle, 20, 348, 150, 35)
        With btnObservacion
            .Name = "btnAgregarObs"
            .TextFrame.Characters.Text = "AGREGAR OBSERVACIÓN"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 9
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(204, 153, 0)
        End With
        
        Set btnCompletarObs = .Shapes.AddShape(msoShapeRoundedRectangle, 180, 348, 150, 35)
        With btnCompletarObs
            .Name = "btnCompletarObs"
            .TextFrame.Characters.Text = "COMPLETAR OBSERVACIÓN"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 9
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(204, 153, 0)
        End With
        
        ' Sección 4: REPORTES
        .Range("A16:H16").Merge
        .Range("A16").Value = "📊 REPORTES"
        With .Range("A16")
            .Font.Bold = True
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(112, 48, 160)
            .HorizontalAlignment = xlLeft
        End With
        
        Set btnReporte = .Shapes.AddShape(msoShapeRoundedRectangle, 20, 452, 150, 35)
        With btnReporte
            .Name = "btnGenerarReporte"
            .TextFrame.Characters.Text = "GENERAR REPORTE"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(112, 48, 160)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(80, 32, 120)
        End With
        
        Set btnImprimir = .Shapes.AddShape(msoShapeRoundedRectangle, 180, 452, 150, 35)
        With btnImprimir
            .Name = "btnImprimirReporte"
            .TextFrame.Characters.Text = "IMPRIMIR REPORTE"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(112, 48, 160)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(80, 32, 120)
        End With
        
        Set btnHistorial = .Shapes.AddShape(msoShapeRoundedRectangle, 340, 452, 150, 35)
        With btnHistorial
            .Name = "btnVerHistorial"
            .TextFrame.Characters.Text = "VER HISTORIAL"
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .Fill.ForeColor.RGB = RGB(112, 48, 160)
            .Line.Weight = 2
            .Line.Color.RGB = RGB(80, 32, 120)
        End With
        
        ' Instrucciones
        .Range("A20:H20").Merge
        .Range("A20").Value = "💡 Haz clic en cualquier botón para comenzar. Todo es intuitivo y seguro."
        With .Range("A20")
            .Font.Italic = True
            .Font.Size = 10
            .Interior.Color = RGB(240, 248, 255)
        End With
        
        .Columns("A:H").ColumnWidth = 15
        
    End With
    
End Sub

Sub CrearEstructuraBase()
    'Crea la estructura base del sistema
    On Error Resume Next
    Dim i As Integer
    
    ' Encabezados Habitaciones
    With wsHabitaciones
        .Range("A1").Value = "#Habitación"
        .Range("B1").Value = "Estado"
        .Range("C1").Value = "Control"
        .Range("D1").Value = "Toallas"
        .Range("E1").Value = "Llaves"
        .Range("F1").Value = "Responsable"
        .Range("G1").Value = "Fecha Última Inspección"
        .Range("H1").Value = "Observaciones Rápidas"
        
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Llenar números de habitaciones
        For i = 1 To NUM_HABITACIONES
            .Range("A" & (i + 1)).Value = i
            .Range("B" & (i + 1)).Value = "Disponible"
            .Range("C" & (i + 1)).Value = 1
            .Range("D" & (i + 1)).Value = 2
            .Range("E" & (i + 1)).Value = 1
            .Range("G" & (i + 1)).Value = Date
        Next i
        
        .Columns("A:H").AutoFit
    End With
    
    ' Encabezados Inventario
    With wsInventario
        .Range("A1").Value = "Item"
        .Range("B1").Value = "En Almacén"
        .Range("C1").Value = "En Uso"
        .Range("D1").Value = "Dadas de Baja"
        .Range("E1").Value = "Total Original"
        .Range("F1").Value = "Última Actualización"
        
        With .Range("A1:F1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 153, 76)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        .Range("A2").Value = "Controles"
        .Range("A3").Value = "Toallas"
        .Range("A4").Value = "Llaves"
        
        For i = 2 To 4
            .Range("B" & i).Value = 30
            .Range("C" & i).Value = 0
            .Range("D" & i).Value = 0
            .Range("E" & i).Formula = "=B" & i & "+C" & i & "+D" & i
            .Range("F" & i).Value = Date
        Next i
        
        .Columns("A:F").AutoFit
    End With
    
    ' Encabezados Observaciones
    With wsObservaciones
        .Range("A1").Value = "Habitación"
        .Range("B1").Value = "Categoría"
        .Range("C1").Value = "Observación"
        .Range("D1").Value = "Prioridad"
        .Range("E1").Value = "Estado"
        .Range("F1").Value = "Fecha Registro"
        .Range("G1").Value = "Asignado a"
        
        With .Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(204, 102, 0)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns("A:G").AutoFit
    End With
    
    ' Encabezados Historial
    With wsHistorial
        .Range("A1").Value = "Fecha y Hora"
        .Range("B1").Value = "Acción"
        
        With .Range("A1:B1")
            .Font.Bold = True
            .Interior.Color = RGB(100, 100, 100)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        .Columns("A:B").AutoFit
    End With
    
End Sub

' ========== FUNCIONES DE CONTROL ==========

Sub AgregarHabitacion()
    Dim habitacion As Integer
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación (1-30):", "Nueva Habitación")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido. Debe estar entre 1 y 30.", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    
    With wsHabitaciones
        .Range("B" & fila).Value = "Disponible"
        .Range("G" & fila).Value = Date
    End With
    
    Call RegistrarHistorial("Habitación " & habitacion & " agregada")
    MsgBox "✅ Habitación " & habitacion & " agregada correctamente", vbInformation
End Sub

Sub EditarHabitacion()
    Dim habitacion As Integer
    Dim estado As String
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación a editar:", "Editar Habitación")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido.", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    
    estado = InputBox("Nuevo estado:" & vbCrLf & "Disponible" & vbCrLf & "Ocupada" & vbCrLf & "Mantenimiento", "Estado")
    
    If estado = "" Then Exit Sub
    
    If estado <> "Disponible" And estado <> "Ocupada" And estado <> "Mantenimiento" Then
        MsgBox "Estado inválido.", vbExclamation
        Exit Sub
    End If
    
    With wsHabitaciones
        .Range("B" & fila).Value = estado
        .Range("G" & fila).Value = Date
    End With
    
    Call RegistrarHistorial("Habitación " & habitacion & " editada. Estado: " & estado)
    MsgBox "✅ Habitación actualizada", vbInformation
End Sub

Sub ActualizarItems()
    Dim habitacion As Integer
    Dim control As Integer
    Dim toallas As Integer
    Dim llaves As Integer
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación:", "Actualizar Items")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido.", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    
    control = InputBox("Cantidad de controles:", "Items")
    toallas = InputBox("Cantidad de toallas:", "Items")
    llaves = InputBox("Cantidad de llaves:", "Items")
    
    With wsHabitaciones
        .Range("C" & fila).Value = control
        .Range("D" & fila).Value = toallas
        .Range("E" & fila).Value = llaves
        .Range("G" & fila).Value = Date
    End With
    
    Call RegistrarHistorial("Items habitación " & habitacion & " actualizados")
    MsgBox "✅ Items actualizados correctamente", vbInformation
End Sub

Sub BuscarHabitacion()
    Dim numeroHab As Integer
    Dim fila As Long
    Dim encontrada As Boolean
    Dim info As String
    
    numeroHab = InputBox("Ingrese número de habitación a buscar:", "Búsqueda")
    
    encontrada = False
    
    With wsHabitaciones
        For fila = 2 To NUM_HABITACIONES + 1
            If .Range("A" & fila).Value = numeroHab Then
                encontrada = True
                info = "HABITACIÓN #" & numeroHab & vbCrLf & vbCrLf & _
                       "Estado: " & .Range("B" & fila).Value & vbCrLf & _
                       "Controles: " & .Range("C" & fila).Value & vbCrLf & _
                       "Toallas: " & .Range("D" & fila).Value & vbCrLf & _
                       "Llaves: " & .Range("E" & fila).Value & vbCrLf & _
                       "Responsable: " & .Range("F" & fila).Value & vbCrLf & _
                       "Última Inspección: " & .Range("G" & fila).Value
                MsgBox info, vbInformation
                Exit For
            End If
        Next fila
    End With
    
    If Not encontrada Then
        MsgBox "Habitación no encontrada.", vbExclamation
    End If
End Sub

Sub DarDeBajaItems()
    Dim itemNombre As String
    Dim cantidad As Integer
    Dim fila As Long
    
    itemNombre = InputBox("Seleccione item:" & vbCrLf & "Controles" & vbCrLf & "Toallas" & vbCrLf & "Llaves", "Dar de Baja")
    cantidad = InputBox("Cantidad a dar de baja:", "Dar de Baja")
    
    With wsInventario
        For fila = 2 To 4
            If .Range("A" & fila).Value = itemNombre Then
                .Range("D" & fila).Value = .Range("D" & fila).Value + cantidad
                .Range("B" & fila).Value = .Range("B" & fila).Value - cantidad
                .Range("F" & fila).Value = Now
                Call RegistrarHistorial(cantidad & " " & itemNombre & " dadas de baja")
                MsgBox "✅ Items dados de baja correctamente", vbInformation
                Exit For
            End If
        Next fila
    End With
End Sub

Sub AgregarObservacion()
    Dim habitacion As Integer
    Dim categoria As String
    Dim observacion As String
    Dim prioridad As String
    Dim fila As Long
    
    habitacion = InputBox("Número de habitación:", "Nueva Observación")
    categoria = InputBox("Categoría:" & vbCrLf & "Cuarto" & vbCrLf & "Camas" & vbCrLf & "Muebles" & vbCrLf & "Electrónicos" & vbCrLf & "Baño" & vbCrLf & "Otros", "Categoría")
    observacion = InputBox("Descripción del problema:", "Observación")
    prioridad = InputBox("Prioridad:" & vbCrLf & "Baja" & vbCrLf & "Media" & vbCrLf & "Alta", "Prioridad")
    
    fila = wsObservaciones.Cells(wsObservaciones.Rows.Count, "A").End(xlUp).Row + 1
    
    With wsObservaciones
        .Range("A" & fila).Value = habitacion
        .Range("B" & fila).Value = categoria
        .Range("C" & fila).Value = observacion
        .Range("D" & fila).Value = prioridad
        .Range("E" & fila).Value = "Pendiente"
        .Range("F" & fila).Value = Date
    End With
    
    Call RegistrarHistorial("Observación agregada a habitación " & habitacion)
    MsgBox "✅ Observación registrada correctamente", vbInformation
End Sub

Sub CompletarObservacion()
    Dim filaObservacion As Long
    Dim responsable As String
    
    filaObservacion = InputBox("Número de fila de observación:", "Completar")
    responsable = InputBox("Responsable:", "Responsable")
    
    With wsObservaciones
        .Range("E" & filaObservacion).Value = "Completada"
        .Range("G" & filaObservacion).Value = responsable
    End With
    
    Call RegistrarHistorial("Observación completada por " & responsable)
    MsgBox "✅ Observación actualizada", vbInformation
End Sub

Sub GenerarReporteDiario()
    Dim ws As Worksheet
    Dim fila As Long
    Dim i As Integer
    
    Set ws = wsReportes
    ws.Cells.Clear
    
    With ws
        .Range("A1:E1").Merge
        .Range("A1").Value = "REPORTE DIARIO DE GESTIÓN DE HABITACIONES"
        With .Range("A1")
            .Font.Bold = True
            .Font.Size = 14
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        fila = 3
        .Range("A" & fila).Value = "Fecha: " & Format(Date, "DD/MM/YYYY HH:MM:SS")
        fila = fila + 2
        
        ' Estado de Habitaciones
        .Range("A" & fila).Value = "ESTADO DE HABITACIONES"
        With .Range("A" & fila)
            .Font.Bold = True
            .Interior.Color = RGB(192, 192, 192)
        End With
        fila = fila + 1
        
        .Range("A" & fila).Value = "Habitación"
        .Range("B" & fila).Value = "Estado"
        .Range("C" & fila).Value = "Controles"
        .Range("D" & fila).Value = "Toallas"
        .Range("E" & fila).Value = "Llaves"
        
        With .Range("A" & fila & ":E" & fila)
            .Font.Bold = True
            .Interior.Color = RGB(192, 192, 192)
        End With
        
        fila = fila + 1
        
        For i = 2 To NUM_HABITACIONES + 1
            .Range("A" & fila).Value = wsHabitaciones.Range("A" & i).Value
            .Range("B" & fila).Value = wsHabitaciones.Range("B" & i).Value
            .Range("C" & fila).Value = wsHabitaciones.Range("C" & i).Value
            .Range("D" & fila).Value = wsHabitaciones.Range("D" & i).Value
            .Range("E" & fila).Value = wsHabitaciones.Range("E" & i).Value
            fila = fila + 1
        Next i
        
        fila = fila + 2
        
        ' Estado de Inventario
        .Range("A" & fila).Value = "ESTADO DE INVENTARIO"
        With .Range("A" & fila)
            .Font.Bold = True
            .Interior.Color = RGB(192, 192, 192)
        End With
        fila = fila + 1
        
        .Range("A" & fila).Value = "Item"
        .Range("B" & fila).Value = "En Almacén"
        .Range("C" & fila).Value = "En Uso"
        .Range("D" & fila).Value = "Dados de Baja"
        
        With .Range("A" & fila & ":D" & fila)
            .Font.Bold = True
            .Interior.Color = RGB(192, 192, 192)
        End With
        
        fila = fila + 1
        
        For i = 2 To 4
            .Range("A" & fila).Value = wsInventario.Range("A" & i).Value
            .Range("B" & fila).Value = wsInventario.Range("B" & i).Value
            .Range("C" & fila).Value = wsInventario.Range("C" & i).Value
            .Range("D" & fila).Value = wsInventario.Range("D" & i).Value
            fila = fila + 1
        Next i
        
        .Columns("A:E").AutoFit
    End With
    
    Call RegistrarHistorial("Reporte diario generado")
    MsgBox "✅ Reporte generado correctamente", vbInformation
End Sub

Sub ImprimirReporteDiario()
    Call GenerarReporteDiario
    wsReportes.PrintOut
    MsgBox "✅ Reporte enviado a imprimir", vbInformation
End Sub

Sub RegistrarHistorial(accion As String)
    Dim fila As Long
    
    With wsHistorial
        fila = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Range("A" & fila).Value = Now
        .Range("B" & fila).Value = accion
    End With
End Sub

Sub ActualizarFechas()
    Dim i As Integer
    
    With wsHabitaciones
        For i = 2 To NUM_HABITACIONES + 1
            If .Range("G" & i).Value = "" Then
                .Range("G" & i).Value = Date
            End If
        Next i
    End With
    
    Call RegistrarHistorial("Fechas actualizadas")
    MsgBox "✅ Fechas actualizadas", vbInformation
End Sub

Sub LimpiarDatos()
    If MsgBox("¿Está seguro de que desea limpiar todos los datos?", vbYesNo) = vbYes Then
        wsHabitaciones.Range("A2:H" & (NUM_HABITACIONES + 1)).Clear
        wsInventario.Range("B2:F4").Clear
        wsObservaciones.Range("A2:G1000").Clear
        Call RegistrarHistorial("Datos limpiados")
        MsgBox "✅ Datos limpiados correctamente", vbInformation
    End If
End Sub
