# 🗄️ Explicación de DatabaseHelper.vb (El Motor de Base de Datos)

* **Ubicación**: [GestionComercial/Data/DatabaseHelper.vb]
* **Propósito**: Es el módulo que centraliza **toda la comunicación con la base de datos**. Ningún formulario ni servicio crea conexiones por su cuenta; todos pasan por `DatabaseHelper`.

---

## 🌟 La gran ventaja: Soporte Híbrido (SQLite y MySQL)

El sistema detecta automáticamente qué base de datos usar mediante la propiedad:
```vb
Public ReadOnly Property IsSQLite As Boolean
    Get
        Return AppConfig.Settings.Provider.Equals("SQLite", StringComparison.OrdinalIgnoreCase)
    End Get
End Property
```
Si es `True`, usa **Microsoft.Data.Sqlite** (guarda todo en un archivo `.db` en tu máquina). Si es `False`, usa **MySqlConnector** (para conectar a un servidor MySQL en red).

---

## 🔍 Explicación de Métodos Principales

### 1. `GetConnection() As DbConnection`
Crea y devuelve un objeto de conexión (`SqliteConnection` o `MySqlConnection`).  
Usa la interfaz base genérica `DbConnection` de .NET, lo que permite que el resto del código no tenga que preocuparse por si la base es MySQL o SQLite.

---

### 2. `PrepareQuery(query As String) As String`
Permite escribir consultas SQL con funciones estándar de MySQL (como `NOW()` o `FOR UPDATE`) y las traduce automáticamente al dialecto de SQLite:
* Reemplaza `NOW()` por `datetime('now', 'localtime')`.
* Reemplaza `CURDATE()` por `date('now', 'localtime')`.
* Elimina `FOR UPDATE` (que en SQLite no existe porque SQLite bloquea todo el archivo en escritura).

---

### 3. Comandos y Parámetros Seguros (Protección contra Inyecciones SQL)
```vb
Public Function CreateCommand(conn As DbConnection, query As String, Optional trans As DbTransaction = Nothing) As DbCommand
Public Sub AddParam(cmd As DbCommand, name As String, value As Object)
```
* **¿Por qué es fundamental?**:
  Nunca debes concatenar variables directamente en la consulta SQL (por ejemplo: `"SELECT * FROM usuarios WHERE nombre = '" & texto & "'"`), porque un atacante podría escribir `' OR '1'='1` y hackear el sistema.
  Con `AddParam(cmd, "@nom", nombre)`, los datos viajan separados de la consulta, impidiendo cualquier ataque de inyección SQL.

---

### 4. Métodos Rápidos de Consulta

* **`ExecuteQuery(query, params) As DataTable`**:
  * Ejecuta una consulta de lectura (`SELECT`) y devuelve una tabla en memoria (`DataTable`) con filas y columnas listas para recorrer.
* **`ExecuteNonQuery(query, params) As Integer`**:
  * Ejecuta acciones que modifican datos (`INSERT`, `UPDATE`, `DELETE`). Devuelve el número de filas que fueron afectadas.
* **`ExecuteScalar(query, params) As Object`**:
  * Ejecuta una consulta y devuelve únicamente la primera columna de la primera fila (ej: `SELECT COUNT(*) FROM productos`).
* **`GetLastInsertedId(conn, trans) As Long`**:
  * Devuelve el último `ID` autoincremental generado (ej: para saber qué ID de ticket se acaba de crear). En SQLite ejecuta `last_insert_rowid()` y en MySQL `LAST_INSERT_ID()`.

---

### 5. Criptografía y Seguridad de Contraseñas (OWASP)

`DatabaseHelper` incluye la seguridad de contraseñas de nivel bancario:

* **`HashPasswordSecure(password As String) As String`**:
  * Encripta la contraseña usando **PBKDF2** con `SHA-256`, una sal aleatoria (*Salt*) de 16 bytes y **100.000 iteraciones**. Esto hace imposible descifrar contraseñas con tablas de arcoíris (*rainbow tables*).
* **`VerifyPassword(password, storedHash, ByRef needsRehash) As Boolean`**:
  * Comprueba si la clave escrita coincide con el hash.
  * Usa `CryptographicOperations.FixedTimeEquals`: comparación en tiempo constante para evitar ataques de temporización (*timing attacks*).
  * **Auto-migración transparente**: Si detecta que el usuario tenía una contraseña con el hash antiguo (`SHA-256 simple`), valida la clave y avisa con `needsRehash = True` para que el sistema la actualice automáticamente al nuevo formato PBKDF2 sin molestar al usuario.

---

### 6. `InitializeDatabaseAndTables(ByRef outMessage) As Boolean`
* Busca automáticamente los archivos `schema_sqlite.sql` y `seed_sqlite.sql` (o sus versiones MySQL) y crea la estructura inicial con tablas y datos por defecto (usuario admin, categorías, talles) si es la primera vez que se ejecuta el sistema.
