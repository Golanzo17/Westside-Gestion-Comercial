# ⚙️ Explicación de AppConfig.vb y appsettings.json

* **Ubicación**: [GestionComercial/Config/AppConfig.vb] y [GestionComercial/appsettings.json]
* **Propósito**: Administrar la configuración persistente del sistema, principalmente la conexión a la base de datos relacional SQLite.

---

## 📄 El archivo `appsettings.json`

Es un archivo de texto en formato JSON donde se guarda la configuración:
```json
{
  "Provider": "SQLite",
  "SqliteFileName": "gestion_comercial.db"
}
```

---

## 🔍 Explicación de las Clases en `AppConfig.vb`

### 1. Clase `DatabaseSettings`
Esta clase representa en memoria los campos de configuración del JSON:
* `Provider`: Indica el motor activo (`"SQLite"`).
* `SqliteFileName`: El nombre del archivo de base de datos local (`"gestion_comercial.db"`).

#### Métodos de `DatabaseSettings`:
* **`GetSqlitePath() As String`**:
  * Busca la carpeta `Database/` en varias ubicaciones relativas (directorio de ejecución `bin/Debug/...`, carpeta raíz del proyecto, etc.) y devuelve la ruta completa absoluta del archivo `.db`.
* **`GetConnectionString() As String`**:
  * Devuelve la cadena de conexión completa para SQLite:
    `Data Source=C:\...\Database\gestion_comercial.db;`

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
