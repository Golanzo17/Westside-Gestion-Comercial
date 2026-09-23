# 🎓 Guía Definitiva para la Defensa Oral del Proyecto
## Westside - Sistema de Gestión Comercial para Locales de Indumentaria

Esta guía fue creada para que nosotros, como equipo de desarrollo, tengamos a mano **todos los conceptos clave, decisiones de diseño, flujo de datos y respuestas precisas** para defender el proyecto ante el profesor con total seguridad y solvencia técnica.

---

## 🧭 1. Resumen Ejecutivo (El "Elevator Pitch" de 1 minuto)

> *"Profesor, nuestro proyecto es un **Sistema de Gestión Comercial integral enfocado en locales de indumentaria y ropa**, desarrollado en **VB.NET con .NET 10** sobre **Windows Forms**. 
> 
> Resuelve los problemas más críticos del rubro: 
> 1. El manejo de stock en matriz de **Talle y Color** (evitando vender prendas que no existen físicamente).
> 2. La **agilidad en el mostrador** con un Punto de Venta (POS) rápido que lee códigos de barra, calcula vueltos y emite comprobantes.
> 3. El **control diario de caja y arqueo de efectivo** para detectar faltantes o sobrantes por turno.
> 4. La **seguridad interna mediante roles estrictos**, asegurando que los vendedores no toquen la configuración y que el Administrador audite sin manipular dinero de caja.
> La arquitectura fue diseñada en capas limpias con persistencia en **SQLite** (embebido, transaccional ACID, portátil y sin necesidad de instalar servidores externos)."*

---

## 🏛️ 2. Arquitectura del Software (Capas y Responsabilidades)

Diseñamos el sistema siguiendo el patrón de **Arquitectura Multicapa (N-Tier)** para desacoplar responsabilidades y hacer el código mantenible y testeable:

```
┌─────────────────────────────────────────────────────────────┐
│ 1. CAPA DE PRESENTACIÓN (UI - Windows Forms)                │
│    FrmMain, FrmVentasPOS, FrmCaja, FrmProductos, etc.       │
│    - 100% construido por código (sin diseñador visual .resx)│
│    - Aplica UITheme para consistencia estética y modo DPI.  │
└──────────────────────────────┬──────────────────────────────┘
                               │ Invoca DTOs y Servicios
                               ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. CAPA DE SERVICIOS / LÓGICA DE NEGOCIO (Services)        │
│    VentaService, AuthService, CajaService, CatalogService...│
│    - Aplica las reglas del local (máx 15% desc., stock mín) │
│    - Orquesta transacciones ACID (commit / rollback).       │
└──────────────────────────────┬──────────────────────────────┘
                               │ Invoca comandos SQL seguros
                               ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. CAPA DE ACCESO A DATOS (Data)                            │
│    DatabaseHelper.vb & AppConfig.vb                         │
│    - Persistencia embebida en SQLite (Microsoft.Data.Sqlite)│
│    - Consultas preparadas y parametrizadas (Anti-SQLi).     │
│    - Hashing criptográfico PBKDF2 (estándar OWASP).         │
└──────────────────────────────┬──────────────────────────────┘
                               │ Lectura / Escritura
                               ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. BASE DE DATOS FÍSICA (Database)                          │
│    gestion_comercial.db (SQLite local, cero configuración)  │
└─────────────────────────────────────────────────────────────┘
```

---

## 🛡️ 3. Seguridad y Segregación de Funciones por Rol (Pregunta Estrella del Profesor)

### ¿Por qué cada rol tiene acceso a lo que tiene?
Aplicamos el principio de seguridad corporativa de **Mínimo Privilegio (Least Privilege)** y **Segregación de Funciones**:

1. **Administrador (Perfil Auditor / TI / Dueño)**:
   * **Qué hace**: Configura datos del negocio, crea y desactiva cuentas de empleados, gestiona catálogo y audita reportes globales de recaudación.
   * **Por qué NO opera POS ni Caja**: Porque un auditor o dueño no debe manipular físicamente el dinero de los turnos diarios. Si el administrador pudiera vender y cerrar caja, podría alterar registros sin contraparte de control.
2. **Gerente (Perfil Supervisión / Operaciones de Salón)**:
   * **Qué hace**: Recibe compras de proveedores, ajusta inventarios con justificación, supervisa al equipo, abre o cierra cajas y **autoriza excepciones** en el salón mediante su contraseña.
   * **Potestad clave**: Si un vendedor quiere dar más del 15% de descuento o hacer un retiro de efectivo, el gerente autoriza de forma presencial sin transferir su cuenta.
3. **Vendedor (Perfil Mostrador / Atención al Cliente)**:
   * **Qué hace**: Cobra rápido en el POS, atiende su propio turno de caja (ingresos, cobros y arqueo), consulta talles disponibles y da de alta clientes.
   * **Restricción**: No puede alterar configuraciones de la empresa, ni ver datos de usuarios, ni aplicar descuentos desmedidos sin firma del supervisor.

---

## 🎯 4. Preguntas Típicas del Profesor y Cómo Responderlas

### P1: ¿Cómo evitan ataques de Inyección SQL (SQL Injection)?
* **Nuestra respuesta**:  
  *"En ningún lugar del sistema concatenamos variables de texto directamente en las sentencias SQL. En su lugar, utilizamos el método `DatabaseHelper.CreateCommand` con parámetros formales mediante `cmd.Parameters.Add(p)`. De esta forma, el motor de base de datos trata cualquier entrada del usuario estrictamente como un dato literal y nunca como código ejecutable."*

### P2: ¿Cómo evitan que el stock quede inconsistente si se corta la luz en medio de una venta?
* **Nuestra respuesta**:  
  *"En `VentaService.ProcesarVenta`, utilizamos transacciones atómicas de base de datos (`DbTransaction`). Todo el cobro, la inserción del ticket en la tabla `ventas`, los renglones en `detalle_ventas` y el descuento de unidades en `producto_talles` se ejecutan dentro del mismo bloque transaccional. Si cualquiera de los pasos falla o falta stock, se ejecuta un `trans.Rollback()`, dejando la base de datos en su estado anterior intacto."*

### P3: ¿Cómo guardan las contraseñas de los usuarios? ¿Están en texto plano?
* **Nuestra respuesta**:  
  *"No, bajo ningún concepto. Implementamos el estándar internacional OWASP utilizando el algoritmo de derivación de claves **PBKDF2** con `HMAC-SHA-256`, una sal criptográfica pseudoaleatoria única de 16 bytes generada por `RandomNumberGenerator` y **100.000 iteraciones**. Además, la verificación de contraseñas se realiza con `CryptographicOperations.FixedTimeEquals` para prevenir ataques de temporización (timing attacks)."*

### P4: ¿Cómo manejan los intentos reiterados de contraseña errónea?
* **Nuestra respuesta**:  
  *"En `AuthService`, programamos un mecanismo en memoria contra ataques de fuerza bruta. Si una cuenta acumula 3 intentos fallidos consecutivos, el sistema bloquea temporalmente los inicios de sesión para ese usuario durante 60 segundos, informando el tiempo de espera restante."*

### P5: ¿Cómo resolvieron el stock de ropa (que viene en diferentes talles y colores)?
* **Nuestra respuesta**:  
  *"No creamos un producto distinto para cada talle, porque saturaría el catálogo. Modelamos una tabla principal `productos` (que almacena el nombre, código de barra principal, categoría y precios) y una tabla relacional `producto_talles` vinculada con `talles`. En `producto_talles` guardamos la cantidad real y el stock mínimo para cada combinación específica de talle y color."*

### P6: ¿Por qué construyeron la interfaz 100% por código en lugar de arrastrar controles en Visual Studio?
* **Nuestra respuesta**:  
  *"Elegimos inicializar los controles por código en `InitializeUI()` por tres razones técnicas:
  1. Evita que los archivos `.Designer.vb` y `.resx` se corrompan ante conflictos de Git o cambios de versión de Visual Studio.
  2. Nos permite desacoplar estilos globales en un módulo centralizado (`UITheme.vb`), facilitando cambios de colores, fuentes y márgenes en todo el software con un solo ajuste.
  3. Soporta de forma nativa el escalado dinámico en pantallas de alta resolución (`AutoScaleMode = AutoScaleMode.Dpi`)."*

### P7: ¿Por qué eligieron SQLite como motor principal de base de datos?
* **Nuestra respuesta**:  
  *"Elegimos **SQLite** por diseño de arquitectura: para un sistema de punto de venta (POS) en local comercial, la confiabilidad, la velocidad y la portabilidad son críticas. SQLite no requiere instalar ni mantener servicios de fondo de bases de datos como MySQL o SQL Server, eliminando puntos de falla si el servicio de base de datos se detiene. El archivo `gestion_comercial.db` es autocontenido, soporta transacciones ACID reales, claves foráneas (`PRAGMA foreign_keys = ON;`) y copias de seguridad inmediatas copiando el archivo."*

---

## 📊 5. Mapa de Archivos Clave para Mostrar en la Pantalla

Si el profesor pide: *"Muéstrenme dónde hacen..."*, abran estos archivos exactos:

| Lo que el profesor quiere ver | Archivo a mostrar | Método / Función |
| :--- | :--- | :--- |
| **Transacción de venta y descuento de stock** | `GestionComercial/Services/VentaService.vb` | `ProcesarVenta(...)` |
| **Seguridad de contraseñas (PBKDF2 y Salt)** | `GestionComercial/Data/DatabaseHelper.vb` | `HashPasswordSecure(...)` |
| **Validación de Login y bloqueo fuerza bruta** | `GestionComercial/Services/AuthService.vb` | `Login(...)` |
| **Consulta SQL segura y parametrizada** | `GestionComercial/Data/DatabaseHelper.vb` | `ExecuteQuery(...)` y `AddParam(...)` |
| **Matriz de talles y stock crítico** | `GestionComercial/Services/CatalogService.vb` | `GetStockMatrizPorProducto(...)` |
| **Arqueo y cierre de caja diaria** | `GestionComercial/Services/CajaService.vb` | `CerrarCaja(...)` |
| **Punto de Venta con atajos de teclado** | `GestionComercial/Forms/FrmVentasPOS.vb` | `FrmVentasPOS_KeyDown(...)` |
| **Matriz de roles y menú lateral** | `GestionComercial/Forms/FrmMain.vb` | `InitializeUI()` (visibilidad de botones) |

---

## 💡 6. Consejos Prácticos para el Momento de Exponer

1. **Repartirse los temas**:
   * Uno puede explicar la **arquitectura general y la base de datos**.
   * Otro puede mostrar el **flujo del Punto de Venta (POS) y la Caja**.
   * Otro puede detallar la **seguridad (AuthService, roles, PBKDF2) y Reportes**.
2. **Hacer una demostración en vivo (Live Demo)**:
   * **Paso 1**: Iniciar sesión como `vendedor` (clave `Vendedor123!`).
   * **Paso 2**: Mostrar que no puede entrar a Configuración ni ver Usuarios (Segregación de roles).
   * **Paso 3**: Abrir caja con $10.000 de cambio inicial.
   * **Paso 4**: Escanear o buscar una prenda en POS (ej. `7790001000018`), seleccionar talle `M`, aplicar descuento del 20% y mostrar cómo salta el cartel de autorización de supervisor.
   * **Paso 5**: Cobrar en efectivo, emitir el comprobante y mostrar cómo se descuenta el stock en la base de datos.
   * **Paso 6**: Cerrar sesión e ingresar como `admin` (clave `Admin123!`) para ver el reporte actualizado y la configuración.
3. **Seguridad y calma**:
   * Si el profesor pregunta algo inesperado, recuerden: *"Diseñamos el sistema de forma modular; esa regla de negocio está centralizada en la capa de servicios para que sea fácil de mantener y modificar sin tocar la pantalla."*
