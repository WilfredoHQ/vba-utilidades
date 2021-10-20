<div id="top"></div>
<br />
<div align="center">
  <a href="https://github.com/wilfredohq/vba-utilidades">
    <img
      src="https://github.com/WilfredoHQ/md-readme/raw/main/images/logo.png"
      alt="Logo"
      width="80"
      height="80"
    />
  </a>
  <h3 align="center">VBA Utilities</h3>
  <p align="center">
    A set of utilities to develop quickly in vba.
    <br />
    <a href="https://github.com/wilfredohq/vba-utilidades">
      <strong>Explore the docs »</strong>
    </a>
    <br />
    <br />
    <a href="https://github.com/wilfredohq/vba-utilidades">
      View Demo
    </a>
    ·
    <a href="https://github.com/wilfredohq/vba-utilidades/issues">
      Report Bug
    </a>
    ·
    <a href="https://github.com/wilfredohq/vba-utilidades/issues">
      Request Feature
    </a>
  </p>
</div>
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a>
      <ul>
        <li>
          <a href="./src/modulos/CRUD.bas">CRUD</a>
          <ul>
            <li><a href="#GetRecords">GetRecords</a></li>
            <li><a href="#ManageRecord">ManageRecord</a></li>
            <li><a href="#DeleteRecord">DeleteRecord</a></li>
          </ul>
        </li>
        <li>
          <a href="./src/modulos/Functions.bas">Functions</a>
          <ul>
            <li><a href="#Sha256">Sha256</a></li>
            <li><a href="#GetEmptyCell">GetEmptyCell</a></li>
            <li><a href="#ValidateFields">ValidateFields</a></li>
            <li><a href="#GetSerialNumber">GetSerialNumber</a></li>
            <li><a href="#GenerateUuid">GenerateUuid</a></li>
            <li><a href="#StrFormat">StrFormat</a></li>
          </ul>
        </li>
        <li>
          <a href="./src/modulos/Procedures.bas">Procedures</a>
          <ul>
            <li><a href="#CleanControls">CleanControls</a></li>
            <li><a href="#FillCombobox">FillCombobox</a></li>
            <li><a href="#GeneratePDF">GeneratePDF</a></li>
            <li><a href="#FormDesign">FormDesign</a></li>
          </ul>
        </li>
        <li>
          <a href="./src/modulos/SQL.bas">SQL</a>
          <ul>
            <li><a href="#ExecuteQuery">ExecuteQuery</a></li>
          </ul>
        </li>
      </ul>
    </li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
  </ol>
</details>

## About The Project

In many projects that I have done with VBA I always used the same functions that I had developed, this became complex because I always copied the functions between files, even having a base file I had it locally. A lot of people liked the features and here it is.

### Built With

- [Excel](https://www.office.com/)
- [VBA](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)

<p align="right">(<a href="#top"> ↑ back to top </a>)</p>

## Getting Started

To get a working local copy, follow these simple steps.

### Prerequisites

Things you need to use the software and how to install them.

- SQLite ODBC Driver (only for using SQL module with SQLite)

  ```tex
  http://www.ch-werner.de/sqliteodbc/
  ```

### Installation

You can download the `.bas` files and import it or directly copy and paste whatever is needed.

<p align="right">(<a href="#top"> ↑ back to top </a>)</p>

## Usage

### GetRecords

Function to find records.

_Dependencies_: `GetEmptyCell`

- Search a record.

  ```bas
  Dim foundRow As Long

  foundRow = CLng(GetRecords(Sheet1, 1, 1, "Value to look for"))

  MsgBox foundRow
  ```

- Search multiple records.

  ```bas
  Dim aRowsFound() As String
  Dim i As Long

  aRowsFound = Split(GetRecords(Sheet1, 1, 1, "Value to look for", True), " ")

  For i = 0 To UBound(aRowsFound)
    MsgBox CLng(aRowsFound(i))
  Next i
  ```

### ManageRecord

Procedure to add or update a record.

- Adding records.

  ```bas
  Dim aData(2) As String
  Dim emptyCell As Long

  emptyCell = GetEmptyCell(Sheet1, 1, 1)

  aData(0) = "1"
  aData(1) = "Jhon"
  aData(2) = "HQ"

  ManageRecord Sheet1, emptyCell, 1, aData()
  ```

- Updating records.

  ```bas
  Dim aData(0) As String
  Dim userRow As Long

  userRow = CLng(GetRecords(Sheet1, 1, 1, "1"))

  aData(0) = "David"

  ManageRecord Sheet1, userRow, 2, aData()
  ```

### DeleteRecord

Procedure to delete a record.

```bas
Dim rowToDeleted As Long

rowToDeleted = CLng(GetRecords(Sheet1, 1, 1, "1"))

DeleteRecord Sheet1, rowToDeleted
```

### Sha256

Sha256 function to encrypt characters.

```bas
MsgBox Sha256("text to encrypt")
```

### GetEmptyCell

Function to obtain the end of a column.

```bas
Dim emptyCell As Long

emptyCell = GetEmptyCell(Sheet1, 1, 1)

MsgBox emptyCell
```

### ValidateFields

Function to validate the fields of a form, for this in the tag of each control use any of the following options **[ number | date | email ]**, if you want to validate that the field is also mandatory add **required**.

```bas
If Not ValidateFields(Me) Then Exit Sub
```

### GetSerialNumber

Function to obtain the serial number of the device.

```bas
MsgBox GetSerialNumber
```

### GenerateUuid

Function to create a GUID / UUID

```bas
MsgBox GenerateUuid
```

### StrFormat

Function to generate a string.

```bas
Dim aValues() As String

ReDim aValues(1)
aValues(0) = "cat"
aValues(1) = "domestic animal"

MsgBox StrFormat("The {0} is a {1}.", aValues)
```

### CleanControls

Procedure for cleaning controls: TextBox, ComboBox, CheckBox, OptionButton and ListBox.

- Clean all controls.

  ```bas
  CleanControls Me
  ```

- Clean only the controls with the tag you see fit.

  ```bas
  CleanControls Me, "tag"
  ```

### FillCombobox

Procedure to fill a ComboBox.

_Dependencies_: `GetEmptyCell`

```bas
FillComboBox Sheet1, 1, 1, Me.ComboBox1
```

### GeneratePDF

Procedure to generate a PDF.

- Generate PDF of a visible sheet.

  ```bas
  GeneratePDF Sheet1, "generated-pdf"
  ```

- Generate PDF of a hidden sheet.

  ```bas
  Sheet1.Visible = xlSheetVisible
  GeneratePDF Sheet1, "generated-pdf"
  Sheet1.Visible = xlSheetVeryHidden
  ```

### FormDesign

Form design procedure.

```bas
FormDesign Me
```

### ExecuteQuery

Function to execute a query in a DB.

_The examples use a database in Spanish_

- Create a record.

  ```bas
  Dim oData As Object
  Dim sql As String

  'sql = "INSERT INTO Productos VALUES (16, 'ProductoNuevo', 3, '4.2')"
  sql = "INSERT INTO Productos (IdProducto, NomProducto, IdGrupo, Precio) VALUES (16, 'ProductoNuevo', 3, '4.2')"
  Set oData = ExecuteQuery(sql)
  Set oData = Nothing
  ```

- Update a record.

  ```bas
  Dim oData As Object
  Dim sql As String

  sql = "UPDATE Productos SET NomProducto='ProductoNuevoModificado' WHERE IdProducto=16"
  Set oData = ExecuteQuery(sql)
  Set oData = Nothing
  ```

- Delete a record.

  ```bas
  Dim oData As Object
  Dim sql As String

  sql = "DELETE FROM Productos WHERE IdProducto=16"
  Set oData = ExecuteQuery(sql)
  Set oData = Nothing
  ```

- Read a record.

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

- Read multiple records.

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

<p align="right">(<a href="#top"> ↑ back to top </a>)</p>

## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/amazing-feature`)
3. Commit your Changes (`git commit -m 'feat: add some amazing-feature'`)
4. Push to the Branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

<p align="right">(<a href="#top"> ↑ back to top </a>)</p>

## License

Distributed under the MIT License. See [LICENSE](LICENSE) for more information.

<p align="right">(<a href="#top"> ↑ back to top </a>)</p>
