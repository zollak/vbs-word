# Adding Multiple Rows to a Table

One way is to rely on the trusty F4 key. Insert a single row into your table, and then repeatedly press the F4 key until you have the number of rows you want.

If you would like to use a macro to do the trick, this one is particularly helpful.

```
Sub AddTableRows()
    If Selection.Information(wdWithInTable) Then
        Application.Dialogs(wdDialogTableInsertRow).Show
    Else
        MsgBox "Insertion point not in a table!"
    End If
End Sub
```

All you need to do is to make sure the insertion point is within the table and then run the macro. When you do, you'll see the Insert Rows dialog box. Just enter the number of rows you want and when you click OK, Word inserts that number into the table.