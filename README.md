excel-vba-editable
==================

EdiTable provides an editable table-like view on your data. Unlike PivotTables, you can edit your data directly in the view.

## Demo

### Screenshots

![Screen1](https://github.com/ftrippel/excel-vba-editable/blob/master/screen1.PNG)

![Screen2](https://github.com/ftrippel/excel-vba-editable/blob/master/screen2.PNG)

### Usage

* Click `Build` to build the view.
* Click `Destroy` to destroy the view.
* Click `Push` to push the values to the view.
* Click `Pull` to pull the values from the view.

## Functionality

* Provides VBA class `Table` and auxiliary classes `Cell`, `LinkedCell`
* Works by creating a link (`LinkedCell`) between the data cell and the view cell
* A `Table` can be created as thus
```visualbasic
Dim T As Table
Set T = New Table
T.AddColumn ActiveSheet.Range("A:A"), ColumnType.Normal
T.AddColumn ActiveSheet.Range("B:B"), ColumnType.Normal
T.AddColumn ActiveSheet.Range("C:C"), ColumnType.value
T.AddColumn ActiveSheet.Range("D:D"), ColumnType.data
T.Build "TableView"
T.Protect "password"
```

## Limitations

* Cannot insert or remove data rows
