# ⚙️ Explicación de AppConfig.vb y appsettings.json

* **Ubicación**: [GestionComercial/Config/AppConfig.vb] y [GestionComercial/appsettings.json]
* **Propósito**: Administrar la configuración persistente del sistema, principalmente la conexión a la base de datos (SQLite o MySQL).

---

## 📄 El archivo `appsettings.json`

Es un archivo de texto en formato JSON donde se guarda la configuración:
```json
{
  "Provider": "SQLite",
  "SqliteFileName": "gestion_comercial.db",
  "Host": "localhost",
  "Port": 3306,
  "Database": "gestion_comercial_db",
  "Username": "root",
  "Password": ""
}
```

---

## 🔍 Explicación de las Clases en `AppConfig.vb`

### 1. Clase `DatabaseSettings`
Esta clase representa en memoria los mismos campos que están en el JSON:
* `Provider`: Indica si el motor es `"SQLite"` o `"MySQL"`.
* `SqliteFileName`: El nombre del archivo de base de datos local.
* `Host`, `Port`, `Database`, `Username`, `Password`: Parámetros para conectar con MySQL.

#### Métodos de `DatabaseSettings`:
* **`GetSqlitePath() As String`**:
  * Busca la carpeta `Database/` en varias ubicaciones relativas (directorio de compilación `bin/Debug/...`, carpeta raíz del proyecto, etc.) y devuelve la ruta completa absoluta del archivo `.db`.
* **`GetConnectionString() As String`**:
  * Devuelve la cadena de conexión completa:
    * Para SQLite: `Data Source=C:\...\Database\gestion_comercial.db;`
    * Para MySQL: Arma una cadena segura con `MySqlConnectionStringBuilder` especificando `Server`, `Port`, `Database`, `utf8mb4`, timeouts, etc.
* **`GetServerOnlyConnectionString() As String`**:
  * Similar a la anterior, pero conecta al servidor MySQL sin seleccionar ninguna base de datos específica (se usa para comprobar si el servidor está encendido y crear la base de datos si no existe).

---

### 2. Módulo `AppConfig`
Es el administrador central estático:
* **`Settings As DatabaseSettings`**:
  * Propiedad que devuelve la configuración activa. Si aún no fue leída del disco, llama automáticamente a `LoadSettings()`.
* **`LoadSettings()`**:
  * Usa `File.ReadAllText(ConfigFilePath)` y `JsonSerializer.Deserialize(Of DatabaseSettings)(json)` para convertir el texto JSON en un objeto real de VB.NET.
  * Si el archivo no existe, crea una configuración con valores predeterminados y llama a `SaveSettings()`.
* **`SaveSettings()`**:
  * Toma el objeto `Settings`, lo convierte a texto JSON con formato ordenado (`WriteIndented = True`) y lo graba en el disco con `File.WriteAllText`.
