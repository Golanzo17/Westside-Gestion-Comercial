# 🏁 Explicación de Program.vb

* **Ubicación en el proyecto**: [GestionComercial/Program.vb]
* **Propósito**: Es el **punto de entrada principal** (*Entry Point*) de toda la aplicación. Es lo primero que ejecuta Windows cuando haces doble clic sobre el ejecutable.

---

## 🔍 Estructura General

El archivo contiene un `Module` llamado `Program` con un método especial llamado `Main`.

### ¿Qué es un `Module` en VB.NET?
Un `Module` es un contenedor de funciones y variables estáticas que están disponibles en toda la aplicación sin necesidad de crear una instancia con `New`.

### ¿Qué significa `<STAThread()>`?
Es un atributo de .NET obligatorio para aplicaciones con interfaz gráfica (Windows Forms). Significa *Single-Threaded Apartment*, e indica que los controles de la ventana interactúan con el sistema operativo en un hilo principal seguro.

---

## 📝 Explicación Función por Función y Bloque por Bloque

```vb
Friend Module Program
    <STAThread()>
    Friend Sub Main(args As String())
```
* `args As String()`: Permite recibir parámetros desde la consola de comandos (por ejemplo, si ejecutas `GestionComercial.exe --test`).

---

### Bloque 1: Inicialización de Base de Datos vía Consola (`--init-db`)
```vb
If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--init-db" Then
    Dim outMsg As String = ""
    Dim ok = Data.DatabaseHelper.InitializeDatabaseAndTables(outMsg)
    Console.WriteLine("INIT_RESULT:" & ok.ToString() & ":" & outMsg)
    Return
End If
```
* **¿Qué hace?**: Si ejecutas el programa pasando `--init-db`, no abre ninguna ventana visual. En su lugar, llama a `DatabaseHelper.InitializeDatabaseAndTables` para crear el archivo SQLite y todas las tablas necesarias automáticamente, imprime el resultado en la consola y termina con `Return`.

---

### Bloque 2: Pruebas Automáticas del Sistema (`--test`)
```vb
If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--test" Then
```
* **¿Qué hace?**: Es una suite completa de diagnóstico integrada que permite verificar sin intervención humana:
  1. **Conexión a la base de datos**: `DatabaseHelper.TestConnection(...)`.
  2. **Autenticación y Login**: `AuthService.Login("admin", "admin123", ...)`.
  3. **Migración de seguridad de contraseñas**: Revisa si el hash en la base empieza con `PBKDF2$SHA256$`.
  4. **Protección Anti-Fuerza Bruta**: Ejecuta un bucle simulando 6 intentos con contraseña incorrecta para verificar que el sistema bloquee al usuario.
  5. **Resolución de scripts**: Comprueba que encuentre `schema_sqlite.sql`.
  6. **Servicios de negocio**: Prueba cargar configuraciones, productos, talles, apertura de caja y simula el cobro de una venta con ticket.
  7. Finaliza con `Return` para no abrir la interfaz gráfica.

---

### Bloque 3: Arranque Visual de Windows Forms
Si el programa se abre normalmente (sin argumentos de consola), llega a las líneas finales:

```vb
Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
Application.EnableVisualStyles()
Application.SetCompatibleTextRenderingDefault(False)
Application.Run(New Forms.FrmLogin())
```

* `SetHighDpiMode`: Garantiza que las pantallas modernas de alta resolución (monitores 4K o pantallas de laptops con escalado al 125% o 150%) no se vean borrosas.
* `EnableVisualStyles`: Activa los estilos visuales modernos de Windows (bordes redondeados, sombras del sistema operativo).
* `Application.Run(New Forms.FrmLogin())`: **Crea y muestra la pantalla de inicio de sesión (`FrmLogin`)**. El programa permanecerá en ejecución mientras esta ventana principal siga activa.
