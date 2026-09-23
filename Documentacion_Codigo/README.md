# 📚 Guía Integral de Aprendizaje del Código: Westside Gestión Comercial

¡Bienvenido! Esta carpeta fue creada especialmente para que puedas entender **absolutamente todo el código de tu sistema**, aunque comiences desde cero sin saber programación en **VB.NET** (Visual Basic .NET).

El sistema está desarrollado con **.NET 10.0** sobre **Windows Forms** y utiliza una arquitectura profesional multicapa.

---

## 🗺️ Mapa de Documentación y Ruta de Aprendizaje Recomendada

Para aprender sin abrumarte, te recomendamos seguir este orden de lectura:

### Fase 1: Fundamentos del Lenguaje y Arquitectura
1. [01_Curso_Rapido_VB_NET.md]:
   * ¿Qué es VB.NET? Sintaxis básica, variables (`Dim`), funciones (`Function`), procedimientos (`Sub`), condicionales y bucles.
2. [02_Arquitectura_Del_Sistema.md]:
   * Cómo viajan los datos: desde que haces clic en un botón en la pantalla hasta que se guardan en la base de datos (SQLite).

---

### Fase 2: El Núcleo de Datos y Configuración
3. [Program.md]:
   * Explicación de [Program.vb] (el archivo donde arranca la aplicación).
4. [AppConfig.md]:
   * Explicación de [AppConfig.vb] y [appsettings.json] (cómo se configuran las conexiones a las bases de datos).
5. [Entities.md]:
   * Explicación de [Entities.vb] (las clases y modelos de datos: Producto, Venta, Cliente, Caja, etc.).
6. [DatabaseHelper.md]:
   * Explicación de [DatabaseHelper.vb] (el motor de conexión que ejecuta consultas SQL de forma segura).

---

### Fase 3: La Lógica del Negocio (Servicios)
7. [AuthService.md]:
   * Inicio de sesión, contraseñas encriptadas (PBKDF2) y bloqueo por fuerza bruta.
8. [VentaService.md]:
   * Cobro de tickets, validación y descuento de stock, trazabilidad y anulación de ventas.
9. [CatalogService.md]:
   * Gestión de productos, categorías, talles, matriz de stock y ajustes manuales.
10. [CajaService.md]:
    * Apertura, ingresos/egresos, arqueo de efectivo y cierre de caja.
11. [Otros_Servicios.md]:
    * Explicación de `ClienteService.vb`, `UsuarioService.vb`, `ReporteService.vb` y `ConfiguracionService.vb`.


---

### Fase 4: La Interfaz Gráfica (Ventanas y Formularios)
13. [UITheme.md]:
    * Colores, fuentes y estilos modernos para todos los controles visuales.
14. [Formularios_Principales.md]:
    * Explicación de `FrmLogin.vb`, `FrmMain.vb`, `FrmVentasPOS.vb`, `FrmProductos.vb`, `FrmStock.vb`, etc.

---

### Fase 5: Preparación para la Defensa Oral
15. [05_Guia_Defensa_Oral/Guia_Defensa_Proyecto.md](file:///c:/Users/gonza/Desktop/Proyecto/Documentacion_Codigo/05_Guia_Defensa_Oral/Guia_Defensa_Proyecto.md):
    * Resumen ejecutivo ("Elevator Pitch" de 1 minuto).
    * Justificación arquitectónica y seguridad por roles.
    * Preguntas típicas del profesor y respuestas técnicas exactas.
    * Guía paso a paso para la demostración en vivo (Live Demo).

---

## 💡 Consejo para estudiar
Cada archivo de esta guía contiene:
* **¿Qué hace este archivo?** (en lenguaje simple y sin tecnicismos difíciles).
* **Explicación bloque por bloque / función por función**.
* **Traducción línea por línea** de las partes más importantes.
* **Conceptos clave de VB.NET** que aparecen en ese archivo para que amplíes tu vocabulario como programador.
