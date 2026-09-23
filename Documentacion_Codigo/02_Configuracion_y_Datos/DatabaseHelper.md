# 🗄️ Explicación de DatabaseHelper.vb (El Motor de Base de Datos)

* **Ubicación**: [GestionComercial/Data/DatabaseHelper.vb]
* **Propósito**: Es el módulo que centraliza **toda la comunicación con la base de datos**. Ningún formulario ni servicio crea conexiones por su cuenta; todos pasan por `DatabaseHelper`.

---

## 🌟 Motor de Persistencia: SQLite Embebido y Portátil

El sistema utiliza **SQLite** mediante `Microsoft.Data.Sqlite`:
```vb
Public ReadOnly Property IsSQLite As Boolean
    Get
        Return True
    End Get
End Property
```
Toda la base de datos reside en un archivo local (`gestion_comercial.db`), lo que garantiza:
1. **Portabilidad total**: No requiere instalar ni configurar servidores de bases de datos externos.
2. **Cero fricción en despliegue**: Al abrir el sistema en una máquina limpia, inicializa automáticamente las tablas y semillas.
3. **Conformidad ACID**: Integridad referencial con claves foráneas habilitadas (`PRAGMA foreign_keys = ON;`) y soporte transaccional completo.

---

## 🔍 Explicación de Métodos Principales

### 1. `GetConnection() As DbConnection`
Crea y devuelve un objeto de conexión `SqliteConnection`.  
Usa la interfaz base genérica `DbConnection` de `System.Data.Common`, asegurando que la carpeta contenedora exista antes de intentar abrir el archivo.

---

### 2. `PrepareQuery(query As String) As String`
Adapta sintaxis SQL al dialecto nativo de SQLite:
* Reemplaza `NOW()` por `datetime('now', 'localtime')`.
* Reemplaza `CURDATE()` por `date('now', 'localtime')`.
* Elimina cláusulas como `FOR UPDATE`.

---

### 3. Comandos y Parámetros Seguros (Protección contra Inyecciones SQL)
```vb
Public Function CreateCommand(conn As DbConnection, query As String, Optional trans As DbTransaction = Nothing) As DbCommand
Public Sub AddParam(cmd As DbCommand, name As String, value As Object)
```
* **¿Por qué es fundamental?**:
  Nunca debes concatenar variables directamente en la consulta SQL (por ejemplo: `"SELECT * FROM usuarios WHERE nombre = '" & texto & "'"`), porque un atacante podría escribir `' OR '1'='1` y vulnerar el sistema.
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
  * Devuelve el último `ID` autoincremental generado tras un `INSERT` dentro de una transacción mediante `last_insert_rowid()`.

---

### 5. Criptografía y Seguridad de Contraseñas (Estándar OWASP)

`DatabaseHelper` implementa seguridad de contraseñas de nivel bancario:

* **`HashPasswordSecure(password As String) As String`**:
  * Encripta la contraseña usando **PBKDF2** con `HMAC-SHA-256`, una sal aleatoria (*Salt*) criptográfica de 16 bytes y **100.000 iteraciones**. Esto previene ataques de fuerza bruta y tablas de arcoíris (*rainbow tables*).
* **`VerifyPassword(password, storedHash) As Boolean`**:
  * Comprueba si la clave escrita coincide con el hash almacenado.
  * Utiliza `CryptographicOperations.FixedTimeEquals`: comparación en tiempo constante para mitigar ataques de canal lateral / temporización (*timing attacks*).

---

### 6. `InitializeDatabaseAndTables(ByRef outMessage) As Boolean`
* Busca automáticamente los archivos `schema_sqlite.sql` y `seed_sqlite.sql` y crea la estructura inicial con tablas y datos por defecto (usuario admin, categorías, talles) si es la primera vez que se ejecuta el sistema.
