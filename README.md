<div align="center">

## Access ADO Bulk Table Copy \(updated\)


</div>

### Description

A valuable tool in your programming toolkit. Use this routine for when an Access table goes bad. Often when this happens conventional INSERT or APPEND queries or cut/paste techniques don't work for backing up your table. This is a routine for copying data from one table to another table, field by field. As the information is processed, bad data fields are skipped over and logged in a log file. Only good data is deposited in the target table. **re-uploaded/Corrected II**
 
### More Info
 
In the Access debug window type the function name and in parenthesis

type your source table name and the target table name and .... away you go!

Make sure you have set a reference to the Microsoft DAO 2.5 or 2.6 Library

Make sure all the "Allow Zero Lengths" table properties in all fields

have been set to YES.

(This will now be done automatically with the recently added "SetZeroLength" routine.)

Make sure the "REQUIRED" property is set to NO.

Create a little log table. And call it "tbl_UpdateLog"

with the following fields:

Name Type Size

ID Number (Long) 4

BadID Number (Long) 4

Comment Text 50

Feelings of relief. Many high fives. A good nights sleep...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ted Calbazana](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ted-calbazana.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ted-calbazana-access-ado-bulk-table-copy-updated__1-27352/archive/master.zip)





### Source Code

```
Public Sub AccessDAOBulkCopy(MySource$, MyTarget$)
'Author: Ted Calbazana
'Date Created: 6/26/2001
'Purpose: A record to record, field to field
'utility for copying only the good data between two tables. (DAO Version)
'Any bad data fields are skipped logged into log table "tbl_UpdateLog".
'This utitlity would be used for worst case scenarios
'such as when one of a tables fields has gotten corrupted.
'(Memo fields are notorious sources of concern)
'When this happens conventional methods of copying or exporting a table will not work.
'TO USE:
'Make sure you have a reference to the Microsoft DAO 2.5 or 2.6 Library
'Make sure all the "Allow Zero Lengths" table properties in all fields have been set to YES.
'(I've added the routine "SetZeroLength" so you don't need to do this manually.
'Make sure you have security permissions)
'Make sure the "REQUIRED" property is set to NO.
'Create a little log table to store error notices. And call it "tbl_UpdateLog"
'with the following fields:
 'Name Type Size
 'ID Number (Long) 4
 'BadID Number (Long) 4
 'Comment Text 50
'In the Access debug window type the function name and in parenthesis
'type your source table name and the target table name
'ie: AccessDAOBulkCopy("MyFavData","MyCleanedData")
'OK - You 're good to go!
 On Error GoTo Err_Handler
 Dim RecordIndex As Long
 Dim FieldIndex As Long
 Dim FieldCount As Long
 Dim RecordCount As Long
 Dim DB As Database
 Dim RS1 As Recordset
 Dim RS2 As Recordset
 Dim MySource As String
 Dim MyTarget As String
 Set DB = DBEngine(0)(0)
 DB.Execute "DELETE * FROM " & MyTarget
 DB.Execute "DELETE * FROM tbl_UpdateLog"
 SetAllowZeroLength (MyTarget) 'It works now.
 'Set the table names right here
 Set RS1 = DB.OpenRecordset(MySource, dbOpenTable)
 Set RS2 = DB.OpenRecordset(MyTarget, dbOpenTable)
 If Not RS1.EOF Then
 FieldCount = RS1.Fields.Count
 RS1.MoveLast
 RecordCount = RS1.RecordCount
 RS1.MoveFirst
 Else
 MsgBox "No Records to Copy", vbInformation
 Exit Sub
 End If
 For RecordIndex = 1 To RecordCount
 RS2.AddNew
 For FieldIndex = 0 To (FieldCount - 1)
 If Not IsNull(RS1.Fields(FieldIndex)) Then
 On Error Resume Next
 If IsDate(RS1.Fields(FieldIndex)) Then
 RS2.Fields(RS1.Fields(FieldIndex).Name) = RS1.Fields(FieldIndex)
 'Log the bad fields
 If Err.Number > 0 Then
 DB.Execute "INSERT INTO tbl_UpdateLog ( BadID, Comment )SELECT " & "'" & Format(RS1.Fields("ID")) & "', '" & "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Debug.Print "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Err.Number = 0
 End If
 ElseIf IsNumeric(RS1.Fields(FieldIndex)) Then
 RS2.Fields(RS1.Fields(FieldIndex).Name) = RS1.Fields(FieldIndex)
 'Log the bad fields
 If Err.Number > 0 Then
 DB.Execute "INSERT INTO tbl_UpdateLog ( BadID, Comment )SELECT " & "'" & Format(RS1.Fields("ID")) & "', '" & "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Debug.Print "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Err.Number = 0
 End If
 Else
 RS2.Fields(RS1.Fields(FieldIndex).Name) = RS1.Fields(FieldIndex) & ""
 'Log the bad fields
 If Err.Number > 0 Then
 DB.Execute "INSERT INTO tbl_UpdateLog ( BadID, Comment )SELECT " & "'" & Format(RS1.Fields("ID")) & "', '" & "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Debug.Print "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Err.Number = 0
 End If
 End If
 DoEvents
 End If
 Next FieldIndex
 Debug.Print "Rec: " & Format(RecordIndex) & " of " & Format(RecordCount)
 RS2.Update
 DoEvents
 RS1.MoveNext
 Next RecordIndex
 Beep
 MsgBox "Processing has been completed.", vbInformation
Quit_Handler:
 Set RS1 = Nothing
 Set RS2 = Nothing
 Set DB = Nothing
 Exit Sub
Err_Handler:
 DB.Execute "INSERT INTO tbl_UpdateLog ( BadID, Comment )SELECT " & "'" & Format(RS1.Fields("ID")) & "', '" & "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Beep
 Debug.Print "Field#" & Format(FieldIndex) & "(" & RS1.Fields(FieldIndex).Name & ") Error or Locked Out" & "'"
 Err = 0
 Resume Quit_Handler
End Sub
Function SetAllowZeroLength(MyTable As String)
'Author: Planet Source Code
'Purpose This function sets the allow zero string to true
'for all Text and Memo fields in all tables in an Access database.
Dim I As Integer, J As Integer
Dim DB As Database, td As TableDef, fld As Field
Set DB = CurrentDb()
'The following line prevents the code from stopping if you do not
'have permissions to modify particular tables, such as system
'tables.
 On Error Resume Next
 For I = 0 To DB.TableDefs.Count - 1
 If DB.TableDefs(I).Name = MyTable Then
 Set td = DB(I)
 For J = 0 To td.Fields.Count - 1
  Set fld = td(J)
  If (fld.Type = DB_TEXT Or fld.Type = DB_MEMO) And Not _
  fld.AllowZeroLength Then
  fld.AllowZeroLength = True
  End If
 Next J
 End If
 Next I
 DB.Close
End Function
```

