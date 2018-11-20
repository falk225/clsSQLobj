# clsSQLobj
VBA Class Module that makes working with DAO recordsets and running update/delete queries on your MS Access DB easy.

## Installation
Download the clsSQLobj.bas file to your local computer and import it into your MS Access Database using File ->  Import File in the built-in VBA IDE.

### References
You must also add the following references. Tools -> References
- Microsoft Scripting Runtime (scrrun.dll)
- Microsoft DAO 3.6 Object Library (dao360.dll)

## Usage
The object can be given a complete SQL statement as a single string or it can be given Select, From, Where, OrderBy, and Insert statements as separate strings which it will then compile into complete SQL statement.

### Querying using a complete SQL string
```vba
dim oSQL as new clsSQLobj
    osql.strSql = "SELECT Field1, Field2, Field3 " & _
         "FROM TableA " & _
         "WHERE Field1='X';"
    'resulting recordset can be used immediately via .rs
    debug.print oSQL.rs.RecordCount
```

### Querying using seperate SQL parts
```vba
dim oSQL as new clsSQLobj
    oSQL.Sel = "Field1, Field2, Field3"
    oSQL.From = "TableA"
    oSQL.Where = "Field1='X'"
    debug.print oSQL.makeStr()
```
SELECT Field1, Field2, Field3 FROM TableA WHERE (Field1='X');
```vba
    debug.print oSQL.rs!Field1
```

### Multiple recordsets
```vba
dim oSQL as new clsSQLobj
    oSQL.strSQL = "my sql string here"
    oSQL.addRecordset "rs1"
    'Alternative One-line Usage: oSQL.addRecordset "rs1", "my sql string here"
    oSQL.strSQL = "my 2nd sql string here"
    oSQL.addRecordset "rs2"
    debug.print oSQL!rs1.recordcount & " " & oSQL("rs2").recordcount
    oSQL.rsClose "rs1"
    oSQL.rsClose "rs2"
```

### Other Functions
#### appendToTable()
Append records specified using From and Where properties to the table specified in Insert. strSQL must be empty.

#### deleteRows()
Removes records specified in Where property from table specified in From property. strSQL must be empty.

#### runUpdateQry()
Creates a QueryDef object using strSQL property and executes it. Good for insert, update, or delete queries.

#### refresh()
Requeries and updates the recordset specified by the .rs method.

#### has_rs()
Returns true if .rs refers to a recordset and false if it has not yet been created.