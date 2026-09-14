# 🚀 Curso Rápido de VB.NET para Principiantes

Visual Basic .NET (VB.NET) es un lenguaje de programación orientado a objetos creado por Microsoft. Su sintaxis está pensada para ser muy cercana al inglés hablado, lo que lo hace muy legible.

A continuación encontrarás los conceptos básicos esenciales que necesitas para entender todo el código de tu proyecto:

---

## 1. Variables: ¿Dónde se guardan los datos?

En VB.NET, para crear una variable se usa la palabra clave `Dim` (que históricamente viene de *Dimension*):

```vb
' Forma tradicional explícita:
Dim precio As Decimal = 1500.50D
Dim nombre As String = "Remera Oversize"
Dim cantidad As Integer = 5
Dim estaActivo As Boolean = True

' Forma moderna con inferencia de tipo (la que usamos en el proyecto):
Dim total = 2500D            ' El compilador sabe que es Decimal por la D
Dim cliente = "Juan Pérez"   ' El compilador sabe que es String
Dim edad = 30                ' El compilador sabe que es Integer
```

### Tipos de datos más usados en este proyecto:
* `Integer`: Números enteros sin coma (ej: `-1, 0, 15, 2024`).
* `Decimal`: Números con coma de altísima precisión, ideales para dinero y precios (ej: `199.99D`). Se le pone `D` al final.
* `String`: Cadenas de texto entre comillas dobles (ej: `"Hola Mundo"`).
* `Boolean`: Solo puede ser `True` (Verdadero) o `False` (Falso).
* `DateTime`: Guarda fecha y hora (ej: `DateTime.Now`).

---

## 2. Métodos: `Sub` vs `Function`

En VB.NET, los bloques de código que hacen algo se dividen en dos categorías:

### A) `Sub` (Procedimiento o Subrutina)
Hace una acción, pero **NO devuelve ningún valor**.
```vb
Public Sub Saludar()
    MessageBox.Show("¡Hola! Bienvenido al sistema.")
End Sub
```

### B) `Function` (Función)
Hace cálculos o búsquedas y **SÍ devuelve un resultado** usando `Return`:
```vb
Public Function CalcularTotal(precio As Decimal, cantidad As Integer) As Decimal
    Dim total = precio * cantidad
    Return total
End Function
```

---

## 3. ¿Qué es `ByVal` y `ByRef`?

Cuando le pasas un dato a una función:
* **`ByVal` (por valor - por defecto)**: La función recibe una copia. Si la función modifica la variable, afuera no cambia.
* **`ByRef` (por referencia)**: La función recibe el enlace a la variable original. Si la función la cambia adentro, afuera también cambia.  
  *En tu proyecto usamos `ByRef errorMessage As String` para que la función pueda devolver `True`/`False` y al mismo tiempo llenar el texto del error si falló.*

```vb
Public Function GuardarDato(ByRef errorMsg As String) As Boolean
    If AlgoSalioMal Then
        errorMsg = "No se pudo conectar a la base de datos."
        Return False
    End If
    Return True
End Function
```

---

## 4. Condicionales: Tomar Decisiones

### `If ... Then ... Else ... End If`
```vb
If stockActual >= cantidadSolicitada Then
    MessageBox.Show("Venta permitida")
Else
    MessageBox.Show("Stock insuficiente")
End If
```

### Operadores lógicos:
* `AndAlso`: Significa "Y" lógico con cortocircuito (si la primera condición es falsa, ni siquiera evalúa la segunda).
* `OrElse`: Significa "O" lógico con cortocircuito (si la primera es verdadera, ya entra).
* `Not`: Niega una condición (ej: `If Not File.Exists(ruta) Then`).

---

## 5. Bucles: Repetir Tareas

### `For Each ... Next` (El más usado en el proyecto)
Recorre una lista o colección elemento por elemento:
```vb
For Each prenda In listaDeProductos
    Console.WriteLine(prenda.Nombre)
Next
```

### `For ... Next` (Con contador numérico)
```vb
For i As Integer = 1 To 5
    Console.WriteLine("Número: " & i)
Next
```

---

## 6. Manejo de Recursos: `Using ... End Using`

Este bloque es fundamental en bases de datos. Garantiza que conexiones, comandos o lectores se cierren y liberen de la memoria RAM inmediatamente al terminar, incluso si ocurre un error:

```vb
Using conn = DatabaseHelper.GetConnection()
    conn.Open()
    ' Trabajamos con la base de datos...
End Using ' <- Aquí la conexión se cierra automáticamente sin peligro de fugas de memoria
```

---

## 7. Manejo de Errores: `Try ... Catch ... End Try`

Permite capturar fallos imprevistos del sistema (como un corte de red o disco lleno) para que la aplicación no se cierre abruptamente:

```vb
Try
    ' Código que podría fallar (ej: leer un archivo en disco)
    Dim texto = File.ReadAllText("archivo.txt")
Catch ex As Exception
    ' Si falló, entra aquí sin romperse
    MessageBox.Show("Hubo un error: " & ex.Message)
End Try
```

---

## 8. Clases y Objetos (Programación Orientada a Objetos)

Una **Clase** es como el plano de una casa; un **Objeto** es la casa ya construida en la memoria.

```vb
Public Class Producto
    ' Propiedades (características del producto):
    Public Property Id As Integer
    Public Property Nombre As String
    Public Property Precio As Decimal

    ' Constructor: lo que se ejecuta al crearlo con "New":
    Public Sub New()
        Precio = 0D
    End Sub
End Class
```

Para crear un producto en código usamos `New`:
```vb
Dim remera = New Producto() With {
    .Id = 1,
    .Nombre = "Remera Oversize",
    .Precio = 15000D
}
```

---

Con estos 8 conceptos en mente, ya tienes el 90% del vocabulario necesario para entender los archivos de tu proyecto.
