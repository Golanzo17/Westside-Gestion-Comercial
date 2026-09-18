# 🖥️ Explicación de los Formularios y Ventanas (Forms)

* **Ubicación**: [GestionComercial/Forms/](file:///c:/Users/gonza/Desktop/Proyecto/GestionComercial/Forms/)
* **Propósito**: Contiene todas las pantallas gráficas con las que interactúa el usuario final.

---

## 💡 Conceptos Clave de Windows Forms en VB.NET

### 1. `Inherits Form`
Indica que nuestra clase es una ventana de Windows. Hereda automáticamente todas las capacidades visuales del sistema operativo (minimizar, maximizar, cerrar, redimensionar).

### 2. Controles Dinámicos por Código vs Diseñador Visual
En este proyecto, los formularios se crean **100% por código limpio** en un método llamado `InitializeUI()`.  
Esto evita que se corrompan los archivos `.Designer.vb` que a menudo causan dolores de cabeza en Visual Studio.

### 3. Eventos (`AddHandler ... AddressOf ...`)
Para que un botón haga algo cuando el usuario le hace clic:
```vb
AddHandler btnCobrar.Click, AddressOf BtnCobrar_Click
```
* `btnCobrar.Click`: El evento que dispara Windows cuando se presiona el botón.
* `AddressOf BtnCobrar_Click`: La dirección de la subrutina (`Sub`) que contiene el código a ejecutar.

---

## 📑 Resumen de los Formularios Principales

### 1. [FrmLogin.vb](Inicio de Sesión)
* **Controles**: Cajas de texto para Usuario y Contraseña (con `PasswordChar = "*"c`), botón "Ingresar".
* **Atajo de teclado**: Captura la tecla `Enter` en el campo contraseña para que no sea necesario tocar el ratón.
* **Lógica**: Llama a `AuthService.Login(...)`. Si es correcto, oculta el login (`Me.Hide()`) y abre el menú principal (`FrmMain`).

---

### 2. [FrmMain.vb](Menú Principal y Navegación)
* **Estructura**:
  * **Barra lateral izquierda (`pnlSidebar`)**: Botones de navegación (Punto de Venta, Catálogo, Stock, Caja, Clientes, Reportes, Sincronización Web, etc.).
  * **Panel de contenido central (`pnlContainer`)**: El espacio donde se cargan los diferentes formularios hijos sin abrir ventanas flotantes desordenadas.
* **Seguridad por Roles**:
  * Al iniciar, comprueba `AuthService.IsAdmin`. Si el usuario es un cajero o vendedor, oculta automáticamente los botones de "Usuarios", "Configuración" y "Reportes Gerenciales".

---

### 3. [FrmVentasPOS.vb] (Punto de Venta TPV)
Es la pantalla más utilizada en el mostrador del local:
* **Lector de Código de Barras**:
  * Tiene una caja de texto con foco permanente (`txtBusqueda`).
  * Al escanear con la pistola láser o escribir y dar Enter, busca en `CatalogService.GetProductoPorCodigo`.
* **Matriz de Talles y Colores**:
  * Si la prenda tiene varios talles o colores, abre un selector rápido para elegir la combinación deseada.
* **Carrito de Compras (`dgvCarrito`)**:
  * Muestra los artículos agregados, cantidades, precios unitarios y subtotales.
  * Permite sumar/restar cantidades con teclado (`+` / `-`).
* **Cálculo de Descuentos, Recargos y Vuelto**:
  * Actualiza en tiempo real el precio final si se ingresa un porcentaje de descuento o recargo por tarjeta.
  * Calcula el vuelto exacto a entregar cuando se ingresa el monto pagado en efectivo.
* **Cobro**:
  * Llama a `VentaService.ProcesarVenta(...)`. Si la venta es exitosa, emite el ticket, descuenta el stock y vacía el carrito listo para el siguiente cliente.

---

### 4. [FrmProductos.vb] y [FrmStock.vb]
* **`FrmProductos`**:
  * Listado completo con filtros de búsqueda rápida.
  * Formulario emergente para crear o editar prendas, definir precios de costo y venta, margen de ganancia automático y asignar los talles disponibles.
* **`FrmStock`**:
  * Vista especializada en inventario. Resalta en color rojo/naranja las prendas con **stock crítico o agotado**.
  * Botón "Ajuste de Stock": Permite ingresar compras a proveedores o descontar mermas con registro obligatorio de motivo.

---

### 5. [FrmCaja.vb] (Control de Efectivo)
* Muestra el estado del turno actual (si la caja está abierta o cerrada).
* Botón **"Abrir Caja"**: Solicita el dinero de cambio inicial.
* Botón **"Ingreso / Egreso Manual"**: Registra gastos menores del local.
* Botón **"Cerrar Caja (Arqueo)"**: Pide contar el efectivo real en mano, calcula la diferencia (sobrante/faltante) e imprime o guarda el balance del día.

---

### 6. [FrmUsuarios.vb] (Gestión de Usuarios y Equipo)
* Permite dar de alta, modificar y desactivar empleados con sus roles (Administrador, Vendedor, Cajero).
* Campos de información personal: Nombre, Apellido, DNI (con validación de duplicados), Teléfono, Email, Ciudad, Dirección y Notas.
* Gestión segura de contraseñas mediante hashing PBKDF2 y control de bloqueo por intentos fallidos.

