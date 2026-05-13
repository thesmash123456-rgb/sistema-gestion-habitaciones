# 🎨 SISTEMA DE GESTIÓN DE HABITACIONES - INTERFAZ GRÁFICA

## ✅ ARCHIVO LISTO: Sistema_Gestion_Habitaciones_GUI.xlsm

---

## 📥 ¿CÓMO DESCARGAR Y USAR?

### **PASO 1: Descargar**
1. Ve a tu repositorio GitHub
2. Descarga: `Sistema_Gestion_Habitaciones_GUI.xlsm`
3. Guarda en tu computadora

### **PASO 2: Abrir en Excel**
1. Haz doble clic en el archivo
2. **Si aparece advertencia de seguridad:**
   - Haz clic en **"Habilitar contenido"**
   - O en **"Enable Content"**

### **PASO 3: ¡A usar!**
- El sistema se abre automáticamente en el **Dashboard**
- Todo tiene **botones visuales** y **campos para llenar**
- **No necesitas escribir código**

---

## 🎯 PANTALLA PRINCIPAL (DASHBOARD)

```
╔════════════════════════════════════════════════════════════╗
║     SISTEMA DE GESTIÓN DE HABITACIONES - DASHBOARD        ║
║                  (30 Habitaciones)                         ║
╠════════════════════════════════════════════════════════════╣
║                                                            ║
║  📋 GESTIÓN DE HABITACIONES                              ║
║  ┌─────────────────────────────────────────────┐         ║
║  │ [AGREGAR HABITACIÓN]                       │         ║
║  │ [EDITAR HABITACIÓN]                        │         ║
║  │ [ACTUALIZAR ITEMS]                         │         ║
║  │ [BUSCAR HABITACIÓN]                        │         ║
║  └─────────────────────────────────────────────┘         ║
║                                                            ║
║  📦 INVENTARIO DE ALMACÉN                                ║
║  ┌─────────────────────────────────────────────┐         ║
║  │ [DAR DE BAJA ITEMS]                        │         ║
║  �� [VER INVENTARIO]                           │         ║
║  └─────────────────────────────────────────────┘         ║
║                                                            ║
║  ⚠️  OBSERVACIONES DE MANTENIMIENTO                       ║
║  ┌─────────────────────────────────────────────┐         ║
║  │ [AGREGAR OBSERVACIÓN]                      │         ║
║  │ [COMPLETAR OBSERVACIÓN]                    │         ║
║  │ [VER OBSERVACIONES]                        │         ║
║  └─────────────────────────────────────────────┘         ║
║                                                            ║
║  📊 REPORTES                                              ║
║  ┌─────────────────────────────────────────────┐         ║
║  │ [GENERAR REPORTE DIARIO]                   │         ║
║  │ [IMPRIMIR REPORTE]                         │         ║
║  │ [VER HISTORIAL]                            │         ║
║  └─────────────────────────────────────────────┘         ║
║                                                            ║
║  🔧 UTILIDADES                                            ║
║  ┌─────────────────────────────────────────────┐         ║
║  │ [ACTUALIZAR FECHAS]                        │         ║
║  │ [LIMPIAR DATOS]                            │         ║
║  └─────────────────────────────────────────────┘         ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
```

---

## 🖱️ CÓMO USAR CADA FUNCIÓN

### **📋 GESTIÓN DE HABITACIONES**

#### [AGREGAR HABITACIÓN]
1. Haz clic en el botón
2. Aparece un cuadro pidiendo número de habitación (1-30)
3. Ingresa el número
4. Se agrega automáticamente con estado "Disponible"
5. ¡Listo!

#### [EDITAR HABITACIÓN]
1. Haz clic en el botón
2. Ingresa número de habitación
3. Selecciona nuevo estado (Disponible/Ocupada/Mantenimiento)
4. Se actualiza automáticamente

#### [ACTUALIZAR ITEMS]
1. Haz clic en el botón
2. Ingresa número de habitación
3. Ingresa cantidad de: Controles, Toallas, Llaves
4. Se sincroniza con el inventario automáticamente

#### [BUSCAR HABITACIÓN]
1. Haz clic en el botón
2. Ingresa número de habitación
3. Ve toda la información de esa habitación

---

### **📦 INVENTARIO DE ALMACÉN**

#### [DAR DE BAJA ITEMS]
1. Haz clic en el botón
2. Selecciona item (Controles/Toallas/Llaves)
3. Ingresa cantidad a dar de baja
4. Se resta del inventario automáticamente

#### [VER INVENTARIO]
1. Ve la hoja de Inventario
2. Muestra: En almacén, En uso, Dados de baja
3. Total disponible se calcula solo

---

### **⚠️ OBSERVACIONES DE MANTENIMIENTO**

#### [AGREGAR OBSERVACIÓN]
1. Haz clic en el botón
2. Ingresa habitación
3. Selecciona categoría:
   - Cuarto (pared, puerta, vidrios)
   - Camas (sábanas, frazadas, sobresábanas)
   - Muebles (roperos, mesas de noche)
   - Electrónicos (teléfono, TV, lámparas, focos)
   - Baño (inodoro, lavamanos, ducha)
   - Otros
4. Describe el problema
5. Selecciona prioridad: Baja/Media/Alta
6. Se registra automáticamente

#### [COMPLETAR OBSERVACIÓN]
1. Haz clic en el botón
2. Ingresa número de fila de observación
3. Ingresa nombre del responsable
4. Marca como "Completada"

---

### **📊 REPORTES**

#### [GENERAR REPORTE DIARIO]
1. Haz clic en el botón
2. Se crea resumen con:
   - Estado de todas las habitaciones
   - Inventario actualizado
   - Observaciones pendientes
3. Se guarda en la hoja "Reportes"

#### [IMPRIMIR REPORTE]
1. Haz clic en el botón
2. Se genera el reporte
3. Se abre el diálogo de impresión
4. Selecciona tu impresora
5. ¡Listo para imprimir!

---

## 🎓 FLUJO DIARIO DE TRABAJO

### **🌅 MAÑANA (Inicio de día)**
```
1. Abre Sistema_Gestion_Habitaciones_GUI.xlsm
2. Haz clic: [ACTUALIZAR FECHAS]
3. Ve cada habitación y verifica estado
4. Si falta algo, haz clic: [ACTUALIZAR ITEMS]
5. Si hay problema, haz clic: [AGREGAR OBSERVACIÓN]
```

### **☀️ DURANTE EL DÍA**
```
1. Cuando se devuelven items: [ACTUALIZAR ITEMS]
2. Cuando se detecta problema: [AGREGAR OBSERVACIÓN]
3. Cuando se repara algo: [COMPLETAR OBSERVACIÓN]
4. Cuando algo se daña: [DAR DE BAJA ITEMS]
```

### **🌙 FINAL DEL DÍA**
```
1. Haz clic: [GENERAR REPORTE DIARIO]
2. Revisa el reporte
3. Si necesitas papel: [IMPRIMIR REPORTE]
4. Guarda archivo: Ctrl + S
```

---

## 💡 CONSEJOS IMPORTANTES

✅ **SIEMPRE:**
- Guardar después de cambios importantes (Ctrl + S)
- Mantener actualizado el estado de habitaciones
- Registrar observaciones de inmediato
- Hacer backup semanal

❌ **NUNCA:**
- Modificar datos directamente en las hojas
- Agregar/eliminar filas manualmente
- Desactivar macros
- Cambiar formato de encabezados

---

## ⚙️ REQUISITOS MÍNIMOS

- ✅ Microsoft Excel 2016 o superior
- ✅ Windows 7 o superior (o Mac)
- ✅ Macros habilitadas
- ✅ 5 MB de espacio libre

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### ❓ "Excel pide habilitar contenido"
**Solución:** Haz clic en "Habilitar contenido" - es normal y seguro

### ❓ "Los botones no funcionan"
**Solución:** 
1. Guarda como .XLSM (no .XLSX)
2. Cierra y abre nuevamente
3. Intenta nuevamente

### ❓ "Aparece error en cuadro de diálogo"
**Solución:** Verifica que ingresaste números válidos (1-30 para habitaciones)

### ❓ "Quiero resetear todo"
**Solución:** Haz clic en [LIMPIAR DATOS] luego [ACTUALIZAR FECHAS]

---

## 📞 ¿PREGUNTAS?

Revisa:
1. Este archivo (GUIA_GUI_EXCEL.md)
2. El archivo INSTALACION_Y_USO.txt
3. Los comentarios dentro de Excel

---

**¡Bienvenido al Sistema de Gestión de Habitaciones!** 🎉

Este sistema fue creado para hacer tu trabajo más fácil y eficiente.
Todo lo que necesitas está a un clic de distancia. 😊
