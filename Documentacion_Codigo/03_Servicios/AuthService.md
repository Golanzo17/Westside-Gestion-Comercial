# 🔐 Explicación de AuthService.vb (Autenticación y Seguridad)

* **Ubicación**: [GestionComercial/Services/AuthService.vb]
* **Propósito**: Gestionar el inicio de sesión, la sesión del usuario conectado, los permisos de administrador y la protección contra ataques de fuerza bruta.

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

### 2. `SyncLock _lockObj ... End SyncLock`
Garantiza seguridad en entornos multihilo (*Thread-Safety*). Si dos usuarios intentaran iniciar sesión al mismo milisegundo, `SyncLock` pone una fila de espera para que no se corrompa el diccionario de intentos en la memoria.

---

## 📝 Explicación de Métodos y Propiedades

### 1. Propiedades de Sesión
* **`CurrentUser As Usuario`**:
  * Guarda el objeto del usuario que tiene la sesión abierta actualmente. Si nadie inició sesión, vale `Nothing`.
* **`IsAuthenticated As Boolean`**:
  * Devuelve `True` si hay un usuario conectado (`CurrentUser IsNot Nothing`).
* **`IsAdmin As Boolean`**:
  * Devuelve `True` solo si el rol del usuario conectado es `"Administrador"`. Los gerentes o vendedores recibirán `False`.

---

### 2. `Login(username, password, ByRef errorMessage) As Boolean`
Es el método que se llama cuando el usuario presiona el botón "Iniciar Sesión":

1. **Paso 1: Control Anti-Fuerza Bruta**:
   * Si el usuario erró la contraseña 5 veces consecutivas, el sistema lo bloquea temporalmente por 30 segundos y rechaza el intento al instante sin consultar la base de datos.
2. **Paso 2: Búsqueda del Usuario**:
   * Busca en la base de datos un usuario activo con ese nombre. Si no existe, cuenta como intento fallido.
3. **Paso 3: Verificación Criptográfica**:
   * Llama a `DatabaseHelper.VerifyPassword`. No compara texto directo, sino que computa el algoritmo seguro **PBKDF2** para verificar si la clave coincide con el hash guardado.
4. **Paso 4: Auto-Migración de Seguridad**:
   * Si el usuario tenía una contraseña con formato viejo (SHA-256 sin sal), el sistema la actualiza silenciosamente al nuevo estándar PBKDF2 sin que el usuario tenga que cambiar su contraseña.
5. **Paso 5: Éxito**:
   * Limpia el contador de intentos fallidos, actualiza `ultimo_login = NOW()` en la base de datos y guarda al usuario en `CurrentUser`.

---

### 3. `Logout()`
* Cierra la sesión activa: pone `_currentUser = Nothing`.

---

### 4. `ShowAdminAuthDialog(parentForm, requiredRole, actionDescription) As Boolean`
* **¿Qué hace?**: Muestra una ventana emergente que solicita usuario y contraseña de un supervisor.
* **¿Para qué sirve?**: Permite que un vendedor común pueda realizar una acción restringida (por ejemplo, aplicar un descuento especial o anular una venta) pidiendo la autorización presencial de un Administrador, sin necesidad de cerrar su sesión de ventas.
