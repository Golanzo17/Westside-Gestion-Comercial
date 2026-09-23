# 🖥️ Explicación de los Formularios y Ventanas (Forms)

* **Ubicación**: 
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
  * **Barra lateral izquierda (`pnlSidebar`)**: Botones de navegación (Panel Principal, Punto de Venta, Catálogo, Stock, Clientes, Caja, Reportes, Usuarios, Configuración).
  * **Barra superior (`pnlTopBar`)**: Sesión activa con distintivo del rol actual, indicador de estado de caja en tiempo real, botón de venta rápida (F1) y reloj sincronizado.
  * **Panel de contenido central (`pnlContentHost`)**: El espacio donde se cargan los diferentes formularios hijos sin abrir ventanas flotantes desordenadas.
* **Seguridad y Segregación de Funciones por Rol (Principle of Least Privilege)**:
  El sistema implementa una arquitectura deliberada de **segregación corporativa de funciones** para prevenir fraudes internos y garantizar auditoría contable. Cada rol posee un ámbito operativo claramente delimitado:
  * **Administrador (Perfil Auditor / TI / Propietario)**:
    * Gestiona la infraestructura del sistema: Configuración general del local, altas/bajas de empleados en el módulo de Usuarios, mantenimiento de Catálogo y visualización de Reportes/Métricas.
    * **Restricción intencional**: No opera el Punto de Venta (POS) ni abre turnos de Caja física diaria, preservando la separación estricta entre la administración/auditoría del negocio y el manejo material de dinero en efectivo.
  * **Gerente (Perfil Supervisión / Operaciones)**:
    * Supervisa el salón y la logística: Control de Stock y reposición, recepción de mercadería, altas y modificaciones de precios y prendas en Catálogo, gestión del equipo en Usuarios, apertura y arqueo de Caja, y análisis en Reportes.
    * Posee facultad de **autorización presencial** (mediante contraseña) para operaciones sensibles de los vendedores (descuentos mayores al 15%, retiros manuales de caja o ajustes de inventario).
  * **Vendedor (Perfil Atención al Público / Caja de Mostrador)**:
    * Enfocado en la atención operativa: Venta rápida en Punto de Venta (POS), operación de su turno de Caja diaria (apertura, cobros y arqueo), consulta del Catálogo de prendas y gestión de la cartera de Clientes.
    * No tiene acceso a la Configuración global del sistema ni a la edición libre de Stock o creación de cuentas de usuario.

| Módulo / Función | Administrador | Gerente | Vendedor |
| :--- | :---: | :---: | :---: |
| **Panel Principal (KPIs)** | ✅ Visible | ✅ Visible | ✅ Visible |
| **Punto de Venta (POS)** | ❌ Restringido (Auditoría) | ✅ Operativo | ✅ Operativo |
| **Catálogo y Talles** | ✅ Gestión total | ✅ Gestión total | ✅ Solo consulta / venta |
| **Stock y Reposición** | ❌ Restringido | ✅ Operativo | 🔒 Requiere Autorización |
| **Caja Diaria y Arqueo** | ❌ Restringido (Auditoría) | ✅ Operativo | ✅ Operativo (su turno) |
| **Clientes** | ✅ Visible | ✅ Visible | ✅ Visible |
| **Reportes y Ventas** | ✅ Completo | ✅ Completo | ✅ Métricas de ventas |
| **Usuarios y Equipo** | ✅ Gestión total | ✅ Gestión de equipo | ❌ Restringido |
| **Configuración General** | ✅ Exclusivo | ❌ Restringido | ❌ Restringido |

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
* Permite dar de alta, modificar y desactivar empleados con sus roles (Administrador, Gerente, Vendedor).
* Campos de información personal: Nombre, Apellido, DNI (con validación de duplicados), Teléfono, Email, Ciudad, Dirección y Notas.
* Gestión segura de contraseñas mediante hashing PBKDF2 y control de bloqueo por intentos fallidos.

