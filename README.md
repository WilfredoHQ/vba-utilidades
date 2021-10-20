# Utilidades vba

Un conjunto de utilidades para desarrollar rápidamente en vba.

## Comenzando 🚀

_Estas instrucciones te permitirán obtener una copia del proyecto en funcionamiento en tu máquina local para propósitos de desarrollo y pruebas._

### Instalación 🔧

Puede descargar los archivos `.bas` e importarlo o copiar y pegar directamente lo que se necesite, para usar el módulo SQL con SQLite también deberás descargar e instalar [`sqliteodbc`](http://www.ch-werner.de/sqliteodbc/)

### Indice de ejemplos 📋

- [CRUD](./src/modulos/CRUD.bas)
  - [GetRecords](#GetRecords)
  - [ManageRecord](#ManageRecord)
  - [DeleteRecord](#DeleteRecord)
- [Functions](./src/modulos/Functions.bas)
  - [Sha256](#Sha256)
  - [GetEmptyCell](#GetEmptyCell)
  - [ValidateFields](#ValidateFields)
  - [GetSerialNumber](#GetSerialNumber)
  - [GenerateUuid](#GenerateUuid)
  - [StrFormat](#StrFormat)
- [Procedures](./src/modulos/Procedures.bas)
  - [CleanControls](#CleanControls)
  - [FillCombobox](#FillCombobox)
  - [GeneratePDF](#GeneratePDF)
  - [FormDesign](#FormDesign)
- [SQL](./src/modulos/SQL.bas)
  - [ExecuteQuery](#ExecuteQuery)

### Uso 🔥

#### GetRecords

Función para buscar registros

_Dependencias_: `GetEmptyCell`

- Buscar un registro

```bas
Dim foundRow As Long

foundRow = CLng(GetRecords(Hoja1, 1, 1, "Valor a buscar"))

MsgBox foundRow
```

- Buscar varios registros

```bas
Dim aRowsFound() As String
Dim i As Long

aRowsFound = Split(GetRecords(Hoja1, 1, 1, "Valor a buscar", True), " ")

For i = 0 To UBound(aRowsFound)
  MsgBox CLng(aRowsFound(i))
Next i
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### ManageRecord

Procedimiento para agregar o modificar un registro

- Agregando registros

```bas
Dim aData(2) As String
Dim emptyCell As Long

emptyCell = GetEmptyCell(Hoja1, 1, 1)

aData(0) = "1"
aData(1) = "Jhon"
aData(2) = "HQ"

ManageRecord Hoja1, emptyCell, 1, aData()
```

- Modificando registros

```bas
Dim aData(0) As String
Dim userRow As Long

userRow = CLng(GetRecords(Hoja1, 1, 1, "1"))

aData(0) = "David"

ManageRecord Hoja1, userRow, 2, aData()
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### DeleteRecord

Procedimiento para eliminar un registro

```bas
Dim rowToDeleted As Long

rowToDeleted = CLng(GetRecords(Hoja1, 1, 1, "1"))

DeleteRecord Hoja1, rowToDeleted
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### Sha256

Función Sha256 para encriptar carácteres

```bas
MsgBox Sha256("texto a encriptar")
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### GetEmptyCell

Función para obtener el final de una columna

```bas
Dim emptyCell As Long

emptyCell = GetEmptyCell(Hoja1, 1, 1)

MsgBox emptyCell
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### ValidateFields

Función para validar los campos de un formulario, para esto en el tag de cada control use cualquiera de las siguientes opciones **[ number | date | email ]**, si quiere validar que el campo también sea obligatorio añada **required**

```bas
If Not ValidateFields(Me) Then Exit Sub
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### GetSerialNumber

Función para obtener el serial de la unidad C

```bas
MsgBox GetSerialNumber
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### GenerateUuid

Función para crear un GUID / UUID

```bas
MsgBox GenerateUuid
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### StrFormat

Función para generar un string

```bas
Dim aValues() As String

ReDim aValues(1)
aValues(0) = "gato"
aValues(1) = "animal doméstico"

MsgBox StrFormat("El {0} es un {1}", aValues)
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### CleanControls

Procedimiento para limpiar los siguientes controles TextBox, ComboBox, CheckBox, OptionButton y ListBox

- Limpiar todos los controles

```bas
CleanControls Me
```

- Limpiar solo los controles con el tag que creas conveniente

```bas
CleanControls Me, "tag"
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### FillCombobox

Procedimiento para llenar un combobox

_Dependencias_: `GetEmptyCell`

```bas
FillComboBox Hoja1, 1, 1, Me.ComboBox1
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### GeneratePDF

Procedimiento para generar un pdf

- Generar pdf de una hoja visible

```bas
GeneratePDF Hoja1, "pdf-generado"
```

- Generar pdf de una hoja oculta

```bas
Hoja1.Visible = xlSheetVisible
GeneratePDF Hoja1, "pdf-generado"
Hoja1.Visible = xlSheetVeryHidden
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### FormDesign

Procedimiento para diseño de formularios

```bas
FormDesign Me
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

#### ExecuteQuery

Función para ejecutar una consulta en una BBDD

- Insertar un registro

```bas
Dim oData As Object
Dim sql As String

'sql = "INSERT INTO Productos VALUES (16, 'ProductoNuevo', 3, '4.2')"
sql = "INSERT INTO Productos (IdProducto, NomProducto, IdGrupo, Precio) VALUES (16, 'ProductoNuevo', 3, '4.2')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing
```

- Actualizar un registro

```bas
Dim oData As Object
Dim sql As String

sql = "UPDATE Productos SET NomProducto='ProductoNuevoModificado' WHERE IdProducto=16"
Set oData = ExecuteQuery(sql)
Set oData = Nothing
```

- Eliminar un registro

```bas
Dim oData As Object
Dim sql As String

sql = "DELETE FROM Productos WHERE IdProducto=16"
Set oData = ExecuteQuery(sql)
Set oData = Nothing
```

- Consultar un registro

```bas
Dim oData As Object
Dim sql As String

sql = "SELECT * FROM Productos WHERE IdProducto=16"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Cells(1, 1) = oData.Fields("IdProducto")
  Cells(1, 2) = oData.Fields("NomProducto")
  Cells(1, 3) = oData.Fields("IdGrupo")
  Cells(1, 4) = oData.Fields("Precio")
Else
  Msgbox "No se encontraron registros"
End if

oData.Close
Set oData = Nothing
```

- Consultar varios registros

```bas
Dim oData As Object
Dim sql As String
Dim firstRow As Long

firstRow = 1

sql = "SELECT * FROM Productos WHERE NomProducto like '%z%'"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Do While Not oData.EOF
    Cells(firstRow, 1) = oData.Fields("IdProducto")
    Cells(firstRow, 2) = oData.Fields("NomProducto")
    Cells(firstRow, 3) = oData.Fields("IdGrupo")
    Cells(firstRow, 4) = oData.Fields("Precio")

    firstRow = firstRow + 1
    oData.movenext
  Loop
Else
  Msgbox "No se encontraron registros"
End if

oData.Close
Set oData = Nothing
```

**⬆ [Regresar al Indice](#indice-de-ejemplos-)**

## Despliegue 📦

1. In process...

## Licencia 📄

Este proyecto está bajo la Licencia (GPL-2.0) - mira el archivo [LICENSE](LICENSE) para detalles.
