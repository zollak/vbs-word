# Copy Table Properties from the first table to all others in the document

```
Sub ScratchMaco()
Dim oTbl As Word.Table
For Each oTbl In ActiveDocument.Tables
    ' Set the base table, the properties of which is copied
    oTbl.Style = ActiveDocument.Tables(1).Style
Next
End Sub
```