<div align="center">

## Execute stored procedure with variables that returns multiple recordset


</div>

### Description

This shows how to two things, first how to pass variables to a stored procedure on a SQL server, second it shows how to handle multiple recordsets being returned from the stored procedure. This allows for a greater speed then trying to pass the entire SQL query to the server especially when you have several. Please Vote if you like it! :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tom Bruinsma](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-bruinsma.md)
**Level**          |Intermediate
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tom-bruinsma-execute-stored-procedure-with-variables-that-returns-multiple-recordset__1-23178/archive/master.zip)





### Source Code

```
Dim SecurityCode As Integer
 Dim LocationCode As Integer
 Dim Engineer As String
 Dim cn As New ADODB.Connection
 Dim rs As ADODB.Recordset
 Dim ConnStr As String
 'Create connection string
 ConnStr = "uid=sa;pwd=;driver={SQL Server};" & _
 "server=<server>;database=<database>;dsn=''"
 'Open the connection the the server
 With cn
 .ConnectionString = ConnStr
 .ConnectionTimeout = 10
 .Properties("Prompt") = adPromptNever
 .Open
 End With
 'Supply the stored procedure and the variables you are going to pass
 'remember to put string and date values in apostophes
 SQLQuery = "sp_WorkList(" & LocationCode & ", '" & Engineer & "', " & SecurityCode & ")"
 'Execute stored procedure
 Set rs = cn.Execute(SQLQuery)
 'If the stored procedure returns any rows of data process the information
 Do While Not rs Is Nothing
 'if we have reached the end of the recordset, get the next recordset that was returned
 Do While Not rs.EOF
 'show the data, i currently use this to populate a treeview... but you can use your imagination
 For Each Field In rs.Fields
 Debug.Print Field.Name & " = " & Field
 Next Field
 Loop
 'get the next recordset
 Set rs = rs.NextRecordset
 Loop
'*******************************
'Example of a SQL stored procedure that returns
'multiple recordsets
'This is copied from my MS SQL Server 7 sp
'I use this to populate a users worklist
'there are 5 fields and the sixth is the name
'of the category with a null in the field value
'I assign the field names to a the name i wish to
'show in the description of the field. You will
'see a 1/2/_ in the field name, these are
'translated to various charchters that SQL server
'permit. Because all the queries are on the server
'all i have to do is modify the stored procedure
'to change the categories in a worklist. It is a
'better then having to recompile!!!
'I hope this helps everyone
'********************************
CREATE PROCEDURE sp_WorkList
	@loccode int,
	@name varchar(100),
	@security int
AS
if @security = 2
	BEGIN
		SELECT [ISR_#] as NSR_#, Institution as Customer, ISR_Rec_d as [Opened], TDate as Due, left(SubProject_Desc,100) as Description, eng_proj_type, '' as Pending_Bid_No_Bid FROM MasterISR WHERE Status = 'Pending Confirm' ORDER BY [ISR_#];
		SELECT [ISR_#] as NSR_#, Institution as Customer, ISR_Rec_d as [Opened], TDate as Due, left(SubProject_Desc,100) as Description, eng_proj_type, '' as Pending_Labor_Assignment FROM MasterISR WHERE Status = 'Pending Assign' ORDER BY [ISR_#];
		SELECT [ISR_#] as NSR_#, Institution as Customer, ISR_Rec_d as [Opened], TDate as Due, left(SubProject_Desc,100) as Description, eng_proj_type, '' as Pending_Submission FROM MasterISR WHERE Status = 'Pending ASR' ORDER BY [ISR_#];
		SELECT [ISR_#] as NSR_#, Institution as Customer, ISR_Rec_d as [Opened], TDate as Due, left(SubProject_Desc,100) as Description, eng_proj_type, '' as Pending_ISR_Number FROM MasterISR WHERE Status = 'Pending ISR' ORDER BY [ISR_#];
	END
if @loccode = 1 GOTO LAN_Eng
if @loccode = 2 GOTO WAN_Eng
LAN_Eng:
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteDueDT as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Pending_KO FROM MasterISR WHERE (LAN_Engineer = @name OR WAN_Engineer = @name) AND (LANCompActDT is null AND Status='Pending KO') ORDER BY ISR_#;
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteDueDT as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Proposal_1_Rework FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (LANCompActDT is null AND Status='Proposal - Rework')) ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteDueDT as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Design FROM MasterISR WHERE (Info_BO is null AND (LAN_Engineer = @name OR WAN_Engineer = @name) AND (LANCompActDT is null AND QuoteCompDT is null AND ((MasterISR.Status)='open' Or (MasterISR.Status)='proposal'))) ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteDueDT as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Design FROM MasterISR WHERE (LAN_Engineer = @name OR WAN_Engineer = @name) AND MasterISR.Status='Design' AND Info_BO is null ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], Info_BO as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Pending_NDP FROM MasterISR WHERE (LAN_Engineer = @name OR WAN_Engineer = @name) AND (Status='Design' OR Status = 'Open') AND Not Info_BO is null ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, QuoteCompDT as Implem2, Network_Target_Date as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Implementation_1_Rework FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (Status='Implementation - Rework' AND LANCompActDT Is Null)) ORDER BY QuoteCompDT;
SELECT [ISR_#] as NSR#, Institution as Customer, QuoteCompDT as Implem2, Network_Target_Date as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Implementation FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (Status='Implementation' AND LANCompActDT Is Null)) ORDER BY QuoteCompDT;
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], Network_Target_Date as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Hold FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (Status='Hold' AND LANCompActDT Is Null)) ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteCompDT as Due, Left(SubProject_Desc, 100) as Description , eng_proj_type, '' as Wait_for_FF FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (Status='Proposal' AND LANCompActDT Is Null)) ORDER BY QuoteCompDT;
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], QuoteCompDT as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Wait_for_FF FROM MasterISR WHERE (LAN_Engineer = @name OR WAN_Engineer = @name) AND (LANCompActDT is null AND Status='Pending FF') ORDER BY [ISR_#];
SELECT [ISR_#] as NSR#, Institution as Customer, ISR_Rec_d as [Open], Network_Target_Date as Due, Left(SubProject_Desc, 100) as Description, eng_proj_type, '' as Engineering_Closed FROM MasterISR WHERE ((LAN_Engineer = @name OR WAN_Engineer = @name) AND (LANCompActDT Is Not Null AND PE_ClosedDT Is Null AND not Status = 'Closed')) ORDER BY Network_Target_Date;
return
```

