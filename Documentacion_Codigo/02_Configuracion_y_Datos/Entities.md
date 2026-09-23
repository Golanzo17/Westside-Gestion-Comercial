# 📦 Explicación de Entities.vb (Modelos de Datos)

* **Ubicación**: [GestionComercial/Models/Entities.vb]
* **Propósito**: Definir las estructuras de datos (entidades) que viajan por todo el sistema. Cada clase representa un concepto de la vida real de una tienda de indumentaria.

---

## 💡 Concepto de VB.NET: ¿Qué es una `Property`?
En VB.NET, una propiedad es una variable dentro de una clase que permite guardar y leer valores de forma controlada:
```vb
Public Property Nombre As String = String.Empty
```
* `Public`: Accesible desde cualquier parte del código.
* `String.Empty`: Inicializa la propiedad con texto vacío `""` para evitar errores de valor nulo (`NullReferenceException`).

---

## 📋 Catálogo de Entidades del Sistema

### 1. `Usuario`
Representa a los empleados o administradores que inician sesión:
* `Id`, `Username`: Nombre de usuario único para acceso al sistema.
* `Dni`: Documento Nacional de Identidad / Cédula única del empleado o administrador.
* `PasswordHash`: Contraseña encriptada (PBKDF2/SHA256, nunca en texto plano).
* `Nombre`, `Apellido`, `NombreCompleto`: Datos del personal.
* `Rol`: Puede ser `"Administrador"`, `"Gerente"` o `"Vendedor"`.
* `Telefono`, `Email`, `Direccion`, `Ciudad`, `FechaNacimiento`, `Notas`.
* `Activo`, `UltimoLogin`, `CreatedAt`.

---

### 2. `Categoria` y `Talle`
Organizan la ropa dentro del local:
* **`Categoria`**: Nombre (ej: "Remeras", "Pantalones"), descripción y estado activo.
  * Tiene `Overrides Function ToString() As String` para que cuando se cargue en un `ComboBox` (menú desplegable), muestre directamente el nombre de la categoría en pantalla.
* **`Talle`**: Nombre (ej: "S", "M", "L", "XL", "42") y un número de `Orden` para mostrarlos ordenados de menor a mayor.

---

### 3. `Producto` y `ProductoTalle` (Matriz de Talles y Colores)
En una tienda de ropa, una prenda tiene variantes de talles y colores:
* **`Producto`**:
  * `CodigoBarra`: Código escaneable (ej: `779001001`).
  * `Nombre`, `Descripcion`, `CategoriaId`, `CategoriaNombre`.
  * `PrecioCosto`: Cuánto te costó comprar la prenda.
  * `PrecioVenta`: A cuánto se vende al público.
  * `PorcentajeGanancia`: Margen de ganancia calculado automáticamente.
  * `TotalStock`: Suma total de todas las unidades de todos los talles de esta prenda.
  * `TallesStock`: Lista de `ProductoTalle` asociados.
* **`ProductoTalle`**:
  * La combinación exacta: Producto + Talle + Color.
  * `StockActual`, `StockMinimo` (aviso de reposición), `SkuEspecifico`.
  * `DescripcionDisplay`: Propiedad de solo lectura (`ReadOnly`) que devuelve texto formateado como `"M (Negro) - Stock: 8"` para mostrar en la interfaz.

---

### 4. `Cliente`
Datos del comprador:
* `DniCuit`, `Nombre`, `Apellido`, `Telefono`, `Email`, `Direccion`, `Ciudad`, `Notas`, `Activo`.
* `NombreCompleto`: Propiedad calculada que concatena `"{Apellido}, {Nombre}"`.

---

### 5. `Caja` y `MovimientoCaja`
Control del dinero físico y digital:
* **`Caja`**:
  * `MontoInicial`: Cambio con el que se abrió la caja.
  * `TotalVentasEfectivo`, `TotalVentasDigital`.
  * `TotalIngresos`, `TotalEgresos`: Movimientos manuales.
  * `MontoEsperado`: El dinero que el sistema calcula que debería haber en el cajón (`MontoInicial + VentasEfectivo + Ingresos - Egresos`).
  * `MontoReal`, `Diferencia`: Arqueo al momento del cierre.
  * `Estado`: `"Abierta"` o `"Cerrada"`.
* **`MovimientoCaja`**:
  * Cada entrada o salida manual de dinero (ej: pago a proveedor, retiro del dueño).

---

### 6. `Venta` y `DetalleVenta`
El comprobante de compra y sus renglones:
* **`Venta`**:
  * `NumeroTicket`: Identificador único (ej: `TICK-20260914-0001`).
  * `Fecha`, `UsuarioId`, `ClienteId`, `CajaId`.
  * `TipoComprobante` (Ticket, Factura), `MetodoPago` (Efectivo, Tarjeta, etc.).
  * `Subtotal`, `DescuentoPorcentaje`, `DescuentoMonto`, `RecargoMonto`, `Total`.
  * `MontoAbonado`, `Vuelto`, `Estado` ("Completada", "Anulada").
  * `Detalles`: Lista con los artículos incluidos en la venta.
* **`DetalleVenta`**:
  * Cada renglón vendido: `ProductoId`, `TalleId`, `Color`, `CodigoBarra`, `DescripcionArticulo`, `PrecioUnitario`, `CostoUnitario`, `Cantidad`, `Subtotal`.

---

### 7. `MovimientoStock` (Auditoría y Trazabilidad)
Cada vez que el stock cambia, queda un registro indeleble:
* `TipoMovimiento`: `"Venta"`, `"Ingreso_Compra"`, `"Ajuste_Manual"`, `"Anulacion_Venta"`.
* `StockAnterior`, `StockPosterior`, `Cantidad`, `Motivo`, `UsuarioId`, `Fecha`.

---

### 8. `ConfiguracionComercio`
Datos de la empresa que se imprimen en los tickets:
* `NombreComercio`, `Cuit`, `Direccion`, `Telefono`, `CondicionIva`, `MensajeTicket`, `MonedaSimbolo`.
