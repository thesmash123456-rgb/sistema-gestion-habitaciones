'========================================
'SISTEMA DE GESTIÓN DE HABITACIONES
'Desarrollado para control de 30 habitaciones
'========================================

Option Explicit

' ========== MÓDULO: VARIABLES GLOBALES ==========
Public wsHabitaciones As Worksheet
Public wsInventario As Worksheet
Public wsObservaciones As Worksheet
Public wsReportes As Worksheet
Public wsHistorial As Worksheet
Public wsDashboard As Worksheet

Public Const NUM_HABITACIONES = 30
Public Const NUM_ITEMS = 3

' ========== MÓDULO: INICIALIZACIÓN ==========

Sub InicializarSistema()
    'Inicializa todas las hojas al abrir el libro
    On Error Resume Next
    
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsHabitaciones = ThisWorkbook.Sheets("Habitaciones")
    Set wsInventario = ThisWorkbook.Sheets("Inventario")
    Set wsObservaciones = ThisWorkbook.Sheets("Observaciones")
    Set wsReportes = ThisWorkbook.Sheets("Reportes")
    Set wsHistorial = ThisWorkbook.Sheets("Historial")
    
    Call CrearEstructuraBase
    MsgBox "Sistema inicializado correctamente", vbInformation, "Éxito"
End Sub

Sub CrearEstructuraBase()
    'Crea la estructura base del sistema
    On Error Resume Next
    
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
        
        ' Formatear encabezados
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Llenar números de habitaciones
        Dim i As Integer
        For i = 1 To NUM_HABITACIONES
            .Range("A" & (i + 1)).Value = i
            .Range("B" & (i + 1)).Value = "Disponible"
            .Range("C" & (i + 1)).Value = 1
            .Range("D" & (i + 1)).Value = 2
            .Range("E" & (i + 1)).Value = 1
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
        
        ' Items predefinidos
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
    
End Sub

' ========== MÓDULO: FUNCIONES DE HABITACIONES ==========

Sub AgregarHabitacion()
    Dim habitacion As Integer
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación (1-30):", "Nueva Habitación")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    
    With wsHabitaciones
        .Range("B" & fila).Value = "Disponible"
        .Range("G" & fila).Value = Date
    End With
    
    MsgBox "Habitación " & habitacion & " agregada", vbInformation
End Sub

Sub EditarHabitacion()
    Dim habitacion As Integer
    Dim estado As String
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación a editar:", "Editar Habitación")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    
    estado = InputBox("Nuevo estado (Disponible/Ocupada/Mantenimiento):", "Estado")
    
    If estado = "" Then Exit Sub
    
    With wsHabitaciones
        .Range("B" & fila).Value = estado
        .Range("G" & fila).Value = Date
    End With
    
    Call RegistrarHistorial("Habitación " & habitacion & " editada. Estado: " & estado)
    MsgBox "Habitación actualizada", vbInformation
End Sub

Sub ActualizarItems()
    Dim habitacion As Integer
    Dim control As Integer
    Dim toallas As Integer
    Dim llaves As Integer
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación:", "Actualizar Items")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido", vbExclamation
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
    
    Call ActualizarInventario("Controles", control)
    Call ActualizarInventario("Toallas", toallas)
    Call ActualizarInventario("Llaves", llaves)
    
    Call RegistrarHistorial("Items habitación " & habitacion & " actualizados")
    MsgBox "Items actualizados", vbInformation
End Sub

Sub BorrarObservacion()
    Dim habitacion As Integer
    Dim fila As Long
    
    habitacion = InputBox("Ingrese número de habitación:", "Borrar Observación")
    
    If habitacion < 1 Or habitacion > 30 Then
        MsgBox "Número de habitación inválido", vbExclamation
        Exit Sub
    End If
    
    fila = habitacion + 1
    wsHabitaciones.Range("H" & fila).Value = ""
    
    MsgBox "Observación eliminada", vbInformation
End Sub

' ========== MÓDULO: FUNCIONES DE INVENTARIO ==========

Sub ActualizarInventario(itemNombre As String, cantidad As Integer)
    Dim fila As Long
    Dim filaInventario As Long
    
    ' Buscar el item en inventario
    With wsInventario
        For fila = 2 To 4
            If .Range("A" & fila).Value = itemNombre Then
                filaInventario = fila
                Exit For
            End If
        Next fila
    End With
    
    If filaInventario = 0 Then Exit Sub
    
    With wsInventario
        .Range("C" & filaInventario).Value = cantidad
        .Range("B" & filaInventario).Value = .Range("E" & filaInventario).Value - cantidad
        .Range("F" & filaInventario).Value = Now
    End With
End Sub

Sub DarDeBajaItems()
    Dim itemNombre As String
    Dim cantidad As Integer
    Dim fila As Long
    
    itemNombre = InputBox("Ingrese nombre del item (Controles/Toallas/Llaves):", "Dar de Baja")
    cantidad = InputBox("Cantidad a dar de baja:", "Dar de Baja")
    
    With wsInventario
        For fila = 2 To 4
            If .Range("A" & fila).Value = itemNombre Then
                .Range("D" & fila).Value = .Range("D" & fila).Value + cantidad
                .Range("B" & fila).Value = .Range("B" & fila).Value - cantidad
                .Range("F" & fila).Value = Now
                Call RegistrarHistorial(cantidad & " " & itemNombre & " dadas de baja")
                MsgBox "Items dados de baja", vbInformation
                Exit For
            End If
        Next fila
    End With
End Sub

' ========== MÓDULO: FUNCIONES DE OBSERVACIONES ==========

Sub AgregarObservacion()
    Dim habitacion As Integer
    Dim categoria As String
    Dim observacion As String
    Dim prioridad As String
    Dim fila As Long
    
    habitacion = InputBox("Número de habitación:", "Nueva Observación")
    categoria = InputBox("Categoría (Cuarto/Camas/Muebles/Electrónicos/Baño/Otros):", "Categoría")
    observacion = InputBox("Descripción de la observación:", "Observación")
    prioridad = InputBox("Prioridad (Baja/Media/Alta):", "Prioridad")
    
    ' Encontrar siguiente fila vacía
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
    MsgBox "Observación registrada", vbInformation
End Sub

Sub CompletarObservacion()
    Dim filaObservacion As Long
    Dim responsable As String
    
    filaObservacion = InputBox("Número de fila de observación a completar:", "Completar")
    responsable = InputBox("Asignado a (nombre del responsable):", "Responsable")
    
    With wsObservaciones
        .Range("E" & filaObservacion).Value = "Completada"
        .Range("G" & filaObservacion).Value = responsable
    End With
    
    Call RegistrarHistorial("Observación completada por " & responsable)
    MsgBox "Observación actualizada", vbInformation
End Sub

' ========== MÓDULO: FUNCIONES DE REPORTES ==========

Sub GenerarReporteDiario()
    Dim ws As Worksheet
    Dim fila As Long
    Dim i As Integer
    
    Set ws = wsReportes
    ws.Cells.Clear
    
    ' Encabezado del reporte
    With ws
        .Range("A1").Value = "REPORTE DIARIO DE GESTIÓN DE HABITACIONES"
        .Range("A2").Value = "Fecha: " & Format(Date, "DD/MM/YYYY")
        .Range("A3").Value = "Hora: " & Format(Time, "HH:MM:SS")
        
        With .Range("A1")
            .Font.Bold = True
            .Font.Size = 14
        End With
        
        fila = 5
        
        ' Estado de Habitaciones
        .Range("A" & fila).Value = "ESTADO DE HABITACIONES"
        .Range("A" & fila).Font.Bold = True
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
        
        ' Copiar datos de habitaciones
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
        .Range("A" & fila).Font.Bold = True
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
    
    MsgBox "Reporte generado correctamente", vbInformation
End Sub

Sub ImprimirReporteDiario()
    Call GenerarReporteDiario
    wsReportes.PrintOut
    MsgBox "Reporte enviado a imprimir", vbInformation
End Sub

' ========== MÓDULO: HISTORIAL ==========

Sub RegistrarHistorial(accion As String)
    Dim fila As Long
    
    With wsHistorial
        fila = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Range("A" & fila).Value = Now
        .Range("B" & fila).Value = accion
    End With
End Sub

' ========== MÓDULO: FUNCIONES DE BÚSQUEDA ==========

Sub BuscarHabitacion()
    Dim numeroHab As Integer
    Dim fila As Long
    Dim encontrada As Boolean
    
    numeroHab = InputBox("Ingrese número de habitación a buscar:", "Búsqueda")
    
    encontrada = False
    
    With wsHabitaciones
        For fila = 2 To NUM_HABITACIONES + 1
            If .Range("A" & fila).Value = numeroHab Then
                encontrada = True
                MsgBox "Habitación " & numeroHab & vbCrLf & _
                        "Estado: " & .Range("B" & fila).Value & vbCrLf & _
                        "Controles: " & .Range("C" & fila).Value & vbCrLf & _
                        "Toallas: " & .Range("D" & fila).Value & vbCrLf & _
                        "Llaves: " & .Range("E" & fila).Value & vbCrLf & _
                        "Responsable: " & .Range("F" & fila).Value, vbInformation
                Exit For
            End If
        Next fila
    End With
    
    If Not encontrada Then
        MsgBox "Habitación no encontrada", vbExclamation
    End If
End Sub

' ========== MÓDULO: UTILIDADES ==========

Sub LimpiarDatos()
    If MsgBox("¿Está seguro de que desea limpiar todos los datos?", vbYesNo) = vbYes Then
        wsHabitaciones.Range("A2:H" & (NUM_HABITACIONES + 1)).Clear
        wsInventario.Range("B2:F4").Clear
        wsObservaciones.Range("A2:G1000").Clear
        MsgBox "Datos limpiados", vbInformation
    End If
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
    
    MsgBox "Fechas actualizadas", vbInformation
End Sub
