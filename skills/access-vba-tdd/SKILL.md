---
name: access-vba-tdd
description: >
  Framework y guía estricta para escribir tests en Microsoft Access VBA compatibles con el runner test-vba (access-vba-sync).
  Trigger: Cuando vayas a escribir tests en VBA, hacer TDD en Access, diseñar un harness de pruebas o usar el comando test-vba.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.2"
---

# SKILL.md — TDD y Testing Profesional en Access VBA

## Objetivo
Instruir a la IA y a los desarrolladores en la creación de una arquitectura de pruebas sólida, mantenible y ejecutable de forma automatizada (vía COM) para bases de datos Microsoft Access (`.accdb` / `.mdb`). Este skill complementa directamente a `access-vba-sync`.

## Cuándo usar este skill
- Cuando estés implementando nueva lógica de negocio en VBA y debas asegurarla con tests.
- Cuando refactorices código legacy en Access y necesites crear una red de seguridad (tests de caracterización).
- Cuando necesites interactuar con el runner `test-vba` de la herramienta `access-vba-sync`.

---

## 1. Fundamentos (Conceptos > Código)

En Access, la automatización COM que ejecuta los tests desde afuera **no tiene magia**. Depende estrictamente de invocar métodos expuestos y leer su resultado escalar. 
Si hacés métodos privados, COM falla. Si devolvés objetos complejos que COM no puede serializar nativamente (como un Recordset o un Dictionary de Scripting), COM falla.

**La robustez se logra devolviendo SIEMPRE un JSON String.** Esto permite a herramientas modernas (como Node.js) parsear el resultado, evaluar aserciones externas e integrarse a pipelines de CI/CD.

### Patrones Críticos (REGLAS DE ORO)

1. **Visibilidad Obligatoria**: TODO procedimiento que actúe como punto de entrada de un test debe ser `Public Function`. NUNCA `Private Sub` ni `Private Function`.
2. **Firma de Retorno**: El test debe devolver un `String` que contenga un JSON válido.
3. **Cero UI**: NUNCA utilices `MsgBox`, `InputBox`, ni abras formularios modales. La ejecución COM es desatendida; cualquier UI bloqueará el proceso permanentemente.
4. **Cero Debug.Print para resultados**: La IA y el runner no pueden capturar la ventana de Inmediato de forma fiable. Todos los logs o trazas deben acumularse en una variable y devolverse en el array `logs` dentro del JSON.
5. **Estructura de Módulos**: Guarda los módulos de test separados del código de producción, usando el prefijo `Test_` (ej. `src/modules/Test_Clientes.bas`).

---

## 2. El Contrato JSON del Runner

El comando `test-vba` espera que cada función de test devuelva una estructura JSON.

### Estructura Mínima Esperada:
```json
{
  "ok": true,
  "value": 42,
  "error": null,
  "logs": [
    "1. Arrange: Usuario creado",
    "2. Act: Cálculo realizado",
    "3. Assert: Todo correcto"
  ]
}
```

---

## 3. Patrones Avanzados de Testing en Access

### 3.1. Aislamiento de Datos (Transacciones DAO)

Para evitar que los tests ensucien la base de datos de desarrollo, envuelve el bloque Arrange/Act en una transacción y haz un Rollback en el Teardown.

```vb
Public Function Test_CrearCliente() As String
    On Error GoTo EH
    Dim logs As String
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    
    ' ARRANGE
    ws.BeginTrans
    logs = """1. Transacción iniciada"""
    
    ' ACT
    Dim nuevoId As Long
    nuevoId = ClienteService_Crear("Juan Perez", "12345678")
    logs = logs & ", ""2. Cliente creado con ID " & CStr(nuevoId) & """"
    
    ' ASSERT
    Dim guardadoCorrecto As Boolean
    guardadoCorrecto = (nuevoId > 0)
    
    ' TEARDOWN
    ws.Rollback
    logs = logs & ", ""3. Rollback ejecutado"""
    
    If guardadoCorrecto Then
        Test_CrearCliente = "{""ok"":true,""value"":" & CStr(nuevoId) & ",""logs"":[" & logs & "]}"
    Else
        Test_CrearCliente = "{""ok"":false,""error"":""ID no generado"",""logs"":[" & logs & "]}"
    End If
    Exit Function

EH:
    On Error Resume Next
    ws.Rollback
    Test_CrearCliente = "{""ok"":false,""error"":""" & Replace(Err.Description, """", "'") & """,""logs"":[" & logs & "]}"
End Function
```

### 3.2. Mocks y Stubs (Inyección de Dependencias manual)

VBA no tiene frameworks de mocking modernos, pero puedes lograr lo mismo pasando interfaces (Class Modules) o usando parámetros opcionales.

**Producción (Clase `Calculadora`):**
```vb
Public Function CalcularDescuento(ByVal monto As Double, Optional ByVal isVip As Boolean = False) As Double
    ' Lógica determinista testeable sin tocar base de datos
    If isVip Then
        CalcularDescuento = monto * 0.2
    Else
        CalcularDescuento = monto * 0.05
    End If
End Function
```

**Test:**
```vb
Public Function Test_Descuento_ClienteVip() As String
    On Error GoTo EH
    Dim resultado As Double
    resultado = CalcularDescuento(100, True)   ' isVip = True → 20% de 100

    If resultado = 20 Then
        Test_Descuento_ClienteVip = "{""ok"":true,""value"":20,""logs"":[""Descuento VIP 20% correcto""]}"
    Else
        Test_Descuento_ClienteVip = "{""ok"":false,""error"":""Esperado 20, recibido " & CStr(resultado) & """}"
    End If
    Exit Function
EH:
    Test_Descuento_ClienteVip = "{""ok"":false,""error"":""" & Replace(Err.Description, """", "'") & """}"
End Function
```

### 3.3. Manejo de Errores Esperados

A veces quieres probar que una función arroje un error específico.

```vb
Public Function Test_ClienteInvalido_LanzaError() As String
    On Error GoTo ManejoErrorEsperado
    Dim logs As String
    
    ' ACT: Esto debería fallar
    Call ClienteService_Crear("", "") 
    
    ' Si llega aquí, el test FALLA (porque no arrojó error)
    Test_ClienteInvalido_LanzaError = "{""ok"":false,""error"":""Se esperaba un error de validación""}"
    Exit Function
    
ManejoErrorEsperado:
    If Err.Number = 9999 Then ' Tu código de error de negocio
        Test_ClienteInvalido_LanzaError = "{""ok"":true,""logs"":[""Error de validación capturado correctamente""]}"
    Else
        Test_ClienteInvalido_LanzaError = "{""ok"":false,""error"":""Error inesperado: " & Replace(Err.Description, """", "'") & """}"
    End If
End Function
```

---

## 4. Orquestación Externa (tests.vba.json)

Para correr múltiples tests sin tener que especificarlos por CLI, la IA debe crear un archivo `tests.vba.json` en la raíz del proyecto.

```json
{
  "tests": [
    {
      "name": "Debe calcular descuento VIP",
      "procedure": "Test_CalcularDescuento_VIP",
      "expect": { "ok": true, "value": 20 }
    },
    {
      "name": "Debe rechazar cliente sin DNI",
      "procedure": "Test_ClienteInvalido_LanzaError",
      "expect": { "ok": true }
    }
  ]
}
```

---

## 5. El Bucle de TDD para la IA

Cuando te pidan implementar una nueva feature, sigue ESTE flujo exacto:

1. **Escribir el Test Primero**:
   - Crea un nuevo módulo `src/modules/Test_<Feature>.bas`.
   - Escribe el test (que va a fallar porque la función de producción no existe o no hace lo correcto).
2. **Actualizar el Plan**:
   - Añade el test al archivo `tests.vba.json`.
3. **Escribir el Código de Producción**:
   - Crea o modifica el módulo en `src/modules/<Feature>.bas`.
4. **Verificar sync de formularios (solo si el código está en un form)**:
   - Si editaste el `.cls` de un formulario, comprueba que esté en sync con su `.form.txt`:
     `node cli.js verify-code Form_<NombreForm> --access "MI_BASE.accdb"`
   - Si reporta DESINCRONIZADO no es un error bloqueante: el siguiente `import` (modo Auto) lo resuelve automáticamente. El paso es informativo.
5. **Sincronizar a Access**:
   - Ejecuta: `node cli.js import Test_<Feature> <Feature> --access "MI_BASE.accdb"`
   - Si el comando falla al abrir la BD porque AutoExec o StartupForm bloquean la apertura, añade `--allow-startup-execution` **solo en ese caso**.
6. **Correr el Runner**:
   - Ejecuta: `node cli.js test-vba --access "MI_BASE.accdb" --json`
   - `test-vba` ejecuta el compile gate automáticamente antes de los tests; no es necesario correr `compile-vba` por separado.
7. **Analizar y Refactorizar**:
   - Si la salida indica `phase: "compile"`, arregla errores de sintaxis primero.
   - Si compila pero falla un test, lee el array `failures` y `run.logs` para entender por qué, corrige el código de producción y repite el bucle.
