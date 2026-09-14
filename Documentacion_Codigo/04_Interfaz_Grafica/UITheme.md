# 🎨 Explicación de UITheme.vb (Diseño y Estilos Visuales)

* **Ubicación**: [GestionComercial/UI/UITheme.vb]
* **Propósito**: Centralizar la apariencia gráfica del sistema (colores, fuentes, bordes redondeados y estilos de tablas) para que todas las ventanas se vean modernas, elegantes y uniformes.

---

## 🌟 La importancia de un Sistema de Diseño (Theme)

En lugar de que cada formulario configure sus propios colores a mano (lo que genera incoherencia visual), todos llaman a las funciones de `UITheme`:

```vb
' Ejemplo en cualquier ventana:
UITheme.StyleButton(btnGuardar, "Success")
UITheme.StyleTextBox(txtNombre)
UITheme.StyleDataGrid(dgvProductos)
```

---

## 🎨 Paleta de Colores Definida

* `ColorPrimary` e `ColorPrimaryDark`: Azul Índigo (`#6366F1`) para acciones principales y elementos destacados.
* `ColorSidebar`: Azul noche profundo (`#111827`) para el menú lateral de navegación.
* `ColorBackground`: Gris muy suave (`#F8FAFC`) para el fondo de las pantallas.
* `ColorSurface`: Blanco puro (`#FFFFFF`) para las tarjetas y paneles de contenido.
* `ColorSuccess`: Verde esmeralda (`#10B981`) para confirmar cobros, ventas y guardados.
* `ColorDanger`: Rojo coral (`#EF4444`) para botones de eliminar, anular o alertas de error.
* `ColorWarning`: Ámbar (`#F59E0B`) para alertas de stock bajo.

---

## 🔤 Tipografías (`Font`)

Usa la tipografía moderna de Windows **Segoe UI** en diferentes tamaños y pesos:
* `FontHeading`: 16pt Negrita (títulos principales).
* `FontSubheading`: 12pt Negrita (secciones y subtítulos).
* `FontRegular` / `FontBold`: 9.5pt (textos habituales y etiquetas).
* `FontPriceBig`: 22pt Negrita (el número gigante que muestra el total a cobrar en el punto de venta).

---

## 🛠️ Métodos de Estilo Automático

### 1. `StyleButton(btn, btnType, rounded)`
Convierte un botón gris anticuado de Windows en un botón moderno:
* Elimina el borde tosco (`btn.FlatStyle = FlatStyle.Flat`).
* Cambia el cursor a "mano" (`Cursors.Hand`).
* Aplica el color según el tipo: `"Primary"`, `"Success"`, `"Danger"`, `"Warning"`, `"Secondary"`, `"Sidebar"`.
* Aplica color de realce al pasar el mouse por encima (`MouseOverBackColor`).
* **Bordes redondeados**: Llama a `ApplyRoundedRegion(btn, 8)` usando matemáticas de dibujo (`System.Drawing.Drawing2D.GraphicsPath`) para recortar las esquinas en forma curva.

### 2. `StyleTextBox(txt)`
* Aplica tipografía prolija, fondo blanco y borde suave a las cajas de texto de entrada.

### 3. `StyleDataGrid(dgv)`
Transforma las tablas de datos (`DataGridView`):
* Elimina bordes dobles anticuados.
* Encabezados oscuros con texto en blanco.
* Filas con altura generosa (36 píxeles) para que no se vean apretadas.
* Efecto visual al pasar el cursor o seleccionar: resalta la fila en azul claro suave (`#EEF2FF`).
* Oculta la columna gris izquierda vacía que viene por defecto en Windows Forms.

### 4. `CreateStatWidget(...) As Panel`
* Construye las tarjetas de resumen del Dashboard (ej: "Total Vendido Hoy: $150.000").
* Genera un contenedor con icono, texto descriptivo y valor numérico destacado.
