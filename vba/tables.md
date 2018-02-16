#### Getting the table
``` vba
Set objTable = Worksheets("[worksheet name]").ListObjects("[table name]")
```

#### Loop throw the table rows
``` vba
For objTableLine = 1 To objTable.ListRows.Count
Do
Next
```

#### Creating a new row in your table
``` vba
'At the begining
Set objRow = objTable.ListRow.Add

'Wherever you want (X is the new row index at the table)
Set objRow = objTable.ListRow.Add X
```

#### Fill the row cells with any value
```vba 
'X is the column index of the cell and YYYYYY is any value
With objRow
	.Range(X) = YYYYYY
End With
```
