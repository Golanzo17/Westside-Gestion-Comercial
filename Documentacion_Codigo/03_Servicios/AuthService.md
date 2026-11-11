# 🔐 Explicación de AuthService.vb (Autenticación y Seguridad)

* **Ubicación**: [GestionComercial/Services/AuthService.vb]
* **Propósito**: Gestionar el inicio de sesión, la sesión del usuario conectado, la consulta de roles, la autorización de supervisores (doble control) y la protección contra ataques de fuerza bruta.
* **Tipo**: es un `Module` de VB.NET, es decir, una clase estática: se usa directamente como `AuthService.Login(...)` sin crear objetos.

---

## 🔍 Conceptos Clave de VB.NET en este archivo

### 1. `Private Structure LoginAttemptInfo`
Una `Structure` en VB.NET es un tipo de dato personalizado y liviano que agrupa campos en memoria:
```vb
Private Structure LoginAttemptInfo
    Public FailedCount As Integer       ' Cantidad de intentos erróneos
    Public LockoutUntil As DateTime     ' Momento hasta el cual está bloqueado
End Structure
```
Se guarda un registro por cada clave en un diccionario en memoria (`_attempts`).

### 2. `SyncLock _lockObj ... End SyncLock`
Garantiza seguridad en entornos multihilo (*Thread-Safety*). Si dos usuarios intentaran iniciar sesión al mismo milisegundo, `SyncLock` pone una fila de espera para que no se corrompa el diccionario de intentos en la memoria.

---

## 📝 Propiedades de Sesión y de Roles

| Propiedad | Qué devuelve |
|---|---|
| `CurrentUser As Usuario` | El usuario con la sesión abierta. Vale `Nothing` si nadie inició sesión. |
| `LastAuthorizingSupervisor As Usuario` | El último supervisor que autorizó una acción (se guarda en `ventas.autorizado_por_id` y `movimientos_caja.autorizado_por_id` para auditoría). |
| `IsAuthenticated` | `True` si hay un usuario conectado. |
| `IsAdmin` | `True` si el rol es `"Administrador"`. |
| `IsManager` | `True` si el rol es `"Gerente"`. |
| `IsVendor` | `True` si el rol es `"Vendedor"`. |
| `IsAdminOrManager` | `True` si es Administrador **o** Gerente. |

Los formularios usan estas propiedades para mostrar u ocultar botones y menús según el rol (segregación de funciones). Por ejemplo, en `FrmMain` el menú de **Reportes** solo es visible si `IsManager`, y el de **Usuarios** solo si `IsAdmin`.

---

## 🔑 `Login(username, password, ByRef errorMessage) As Boolean`
Es el método que se llama cuando el usuario presiona el botón "Iniciar Sesión":

1. **Control Anti-Fuerza Bruta**: tras **5 intentos fallidos consecutivos** (`MaxIntentos`), la clave queda bloqueada **60 segundos** (`SegundosBloqueo`) y los intentos se rechazan al instante sin consultar la base.
2. **Búsqueda del Usuario**: busca un usuario activo con ese nombre. Si no existe, cuenta como intento fallido.
3. **Verificación Criptográfica**: llama a `DatabaseHelper.VerifyPassword`, que computa **PBKDF2 (SHA-256, 100.000 iteraciones)** y lo compara con el hash guardado. Nunca se compara texto plano.
4. **Auto-Migración de Seguridad**: si la contraseña estaba en un formato viejo (SHA-256 sin sal), se actualiza silenciosamente a PBKDF2.
5. **Éxito**: limpia el contador de fallos, actualiza `ultimo_login` y guarda al usuario en `CurrentUser`. Se registra en el log de auditoría.

---

## 🚪 `Logout()`
* Cierra la sesión: deja `CurrentUser` y `LastAuthorizingSupervisor` en `Nothing` y registra el evento `LOGOUT` en auditoría.

---

## 🛡️ Autorización de Supervisores (Doble Control)

Algunas acciones requieren que **otra persona con más jerarquía** las apruebe. El sistema abre el formulario `FrmAutorizacionSupervisor`, que pide **solo la contraseña** del supervisor; el sistema busca entre los usuarios activos del rol requerido a quién pertenece esa contraseña.

### `AutenticarSupervisorPorRoles(password, rolesPermitidos, motivo, ByRef supervisor, ByRef errorMessage)`
* Es el motor de la validación. Aplica el mismo bloqueo anti fuerza bruta (clave `__SUPERVISOR_AUTH__`, 5 intentos / 60 s).
* Recorre los roles permitidos y, con un helper privado (`ValidarCredencialesPorRol`), consulta los usuarios activos de ese rol verificando el hash.
* Si coincide, guarda al supervisor en `LastAuthorizingSupervisor` y registra en auditoría: supervisor, operador y motivo.

### Las cuatro funciones `SolicitarAutorizacion...`

| Función | Quién puede autorizar | Se saltea el diálogo si... | Uso típico |
|---|---|---|---|
| `SolicitarAutorizacionManager` | Gerente | el usuario actual ya es Gerente | Anular una venta |
| `SolicitarAutorizacionGerente` | Gerente | **nunca**: siempre pide la clave, incluso a un Administrador | Crear, exportar y restaurar backups (control dual) |
| `SolicitarAutorizacionAdminOManager` | Administrador o Gerente | el usuario actual ya es uno de ellos | Descuentos de más de 15%, egresos de caja, altas y cambios de catálogo o stock de un Vendedor |
| `SolicitarAutorizacionAdmin` | Administrador | el usuario actual ya es Administrador (salvo `forzarConfirmacionClave = True`) | Entrar a Configuración |

Todas devuelven `True` si la acción fue autorizada y `False` si se canceló o falló.
