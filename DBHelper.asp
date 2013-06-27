<%
'<!-- include file="Collection.asp" -->
' Functions that Depends on Collection.asp are commented out.
' May add support for C_ArraySet.cls.

' DBHelper written by Darren Bates.

Function NewDB(sConnection)
	Set NewDB = (New DBHelper).Init(sConnection)
	'NewDB.ConnectionString = sConnectionString
End Function

Const RS_READ_ONLY = true
Const RS_READ_WRITE = false

Class DBHelper
	private dbConnectionString
	private oConn 
	
	Public Function Init(sConnection)
		If IsObject(sConnection) then ' assume it is an adodb.connection object
			Set oConn = sConnection
		Else ' if it is a string
			dbConnectionString = sConnection
			OpenDBConnection
        End If
		Set Init = me
	End Function
	
	'Sub class_initialize()
	'End Sub
	
	Public Property Let ConnectionString(value)
		dbConnectionString = value
	End Property

	Public Property Get ConnectionString
		ConnectionString = dbConnectionString
	End Property

	Private Function OpenDBConnection()

		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.CursorLocation = adUseClient
		
		on error resume next
		Call oConn.Open(dbConnectionString)
		If err.number <> 0 then
			lp "Our database is temporarily down. Please try again in a few minutes."
			lpe "If you continue to get this message, please contact <a href='mailto:" & application("SupportEmail") & "'>Support</a>."
		end if
		on error goto 0
		
		'Set OpenDBConnection = oConn
	End Function


	Private Function OpenRecordset(ByVal sSQL, ByVal bReadOnly)
		Dim oRS
		Dim nRSOption
		Dim nLockOption

		' session("Debug") is set to true when logged on through /Includes/Debug.asp
		If session("Debug") = true then session("sSQL") = sSQL ' primarily for debug purposes
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
			
		If bReadOnly Then
			nRSOption   = adOpenStatic
			nLockOption = adLockReadOnly
		Else
			nRSOption   = adLockPessimistic
			nLockOption = adLockOptimistic
		End If

		oRS.CursorLocation = adUseClient
		on error resume next
		Call oRS.Open(sSQL, oConn, nRSOption, nLockOption)
		if err.number <> 0 then
			lp "A Data error has occured: " & err.Description
			lp "Please contact " & Session("Organization") & " support so we can fix this issue as quickly as possible." 
			session("sSQL") = sSQL
			'TODO: send an email to support/dev
			'lpe sSQL
			lpe ""
		end if
		on error goto 0
		Set OpenRecordset = oRS
	End Function


	'' Shortcut functions to reduce code duplication
	' Open a recordset in read-only mode
	Public Function GetRS( ByVal sSQL )
		set GetRS = OpenRecordset( sSQL, RS_READ_ONLY )
	End Function

	'Open a recordset in read/write mode
	Public Function GetRSRW( ByVal sSQL )
		set GetRSRW = OpenRecordset( sSQL, RS_READ_WRITE )
	End Function


	' Opens a recordset in read/write mode
	' Provide the Table name, the list of fields needed (default is "*"), 
	Function GetTableRS( ByVal sTable, ByVal sFields, ByVal sFilter )
		Set GetTableRS = GetTableRSoption( sTable, sFields, sFilter, RS_READ_ONLY )
	End Function

	Function GetTableRSRW( ByVal sTable, ByVal sFields, ByVal sFilter )
		Set GetTableRSRW = GetTableRSoption( sTable, sFields, sFilter, RS_READ_WRITE )
	End Function

	' Returns the first value of the first field using GetRS
	Public Function GetScalar( ByVal sSQL )
		Dim rs: Set rs = GetRS( sSQL )
		If rs.EOF then
			GetScalar = null
		Else
			GetScalar = rs.Fields(0).Value
		End If
		DisposeRS rs
	End Function

	' Returns the values in the first row as a fieldname/value dictionary
	Public Function Get1stRecord( ByVal sTable, ByVal sField, ByVal sFilter )
		'Dim dict: Set dict = New Collection
		Dim dict: Set dict = Server.CreateObject("Scripting.Dictionary")
		Dim rs: Set rs = GetRS( sTable, sField, sFilter )
		Dim f
		If not rs.EOF then
			For each f in rs.Fields
				dict.Add f.name, rs(f.name).value
			Next
		End If
		DisposeRS rs
		Set Get1stRecord = dict
	End Function

	' Returns the values in the first row as a fieldname/value dictionary
	Public Function GetRecord( ByVal sSQL )
		'Dim dict: Set dict = New Collection
		Dim dict: Set dict = Server.CreateObject("Scripting.Dictionary")
		Dim rs: Set rs = GetRS( sSQL )
		If not rs.EOF then
			Dim f
			For each f in rs.Fields
				dict.Add f.name, rs(f.name).value
			Next
		End If
		DisposeRS rs
		Set GetRecord = dict
	End Function

	' Returns a dictionary with the first field as a key and the second field as the value
	Public Function GetRecordsDict( ByVal sTable, ByVal sFields, ByVal sFilter )
		Dim d: Set d = Server.CreateObject("Scripting.Dictionary")
		'Dim d: Set d = New Collection
		Dim rs: Set rs = GetRS( sTable, sFields, sFilter )
		Dim nFields: nFields = rs.Fields.Count
		Dim b2nd: b2nd = (nFields > 1)
		Do while not rs.EOF
			Dim sKeyValue : sKeyValue = rs(0).value
			If not d.Exists(sKeyValue) then
				'call d.Add( rs(0).value, rs(IIf(b2nd, 1, 0)).value )
				Select Case nFields
				Case 0
					Exit Function
				Case 1
					call d.Add( sKeyValue, sKeyValue )
				Case 2
					call d.Add( sKeyValue, rs(1).value )
				Case Else ' > 2
					Dim dFields: set dFields = Server.CreateObject("Scripting.Dictionary")
					'Dim dFields: set dFields = New Collection
					Dim f
					For each f in rs.Fields
						If not dfields.Exists( f.Name ) then
							dFields.Add f.Name, f.Value
						End If
					Next
					d.Add sKeyValue, dFields
				End Select
			End If
			rs.MoveNext
		Loop
		Set GetRecordsDict = d
		Set d = nothing
		DisposeRS rs
	End Function

	' Returns a dictionary with the first field as a key and the second field as the value
	'	If there are more than 2 fields, the values will be a sub-dictionary of the fields
	Public Function GetAllRecords( ByVal sSQL )
		Dim d: Set d = Server.CreateObject("Scripting.Dictionary")
		'Dim d: Set d = New Collection
		Dim rs: Set rs = GetRS( sSQL )
		Dim nFields: nFields = rs.Fields.Count
		Dim b2nd: b2nd = (nFields > 1)
		Dim f
		Do while not rs.EOF
			Dim sKeyValue : sKeyValue = rs(0).value
			If not d.Exists(sKeyValue) then
				Select Case nFields
				Case 0
					Exit Function
				Case 1
					call d.Add( sKeyValue, sKeyValue )
				Case 2
					call d.Add( sKeyValue, rs(1).value )
				Case Else ' > 2
					'Dim dFields: set dFields = New Collection
					Dim dFields: set dFields = Server.CreateObject("Scripting.Dictionary")
					For each f in rs.Fields
						If not dfields.Exists( f.Name ) then
							dFields.Add f.Name, f.Value
						End If
					Next
					d.Add sKeyValue, dFields
				End Select
			End If
			rs.MoveNext
		Loop
		Set GetAllRecords = d
		Set d = nothing
		DisposeRS rs
	End Function
	
	Function SelectFrom( sTable, sFields, sFilter)
		If not sFields > "" then sFields = "*"
		If not sFilter > "" then sFilter = "1 = 1"
		
		SelectFrom = " Select " & sFields & vbCRLF _
			&	" From	" & sTable & vbCRLF _
			&	" Where	" & sFilter
		
	End Function

	Function GetTableRSoption( sTable, sFields, sFilter, bReadOnly )
		Dim sSQL
		sSQL =	SelectFrom( sTable, sFields, sFilter )

		Set GetTableRSoption = OpenRecordset( sSQL, bReadOnly )
		'ExecuteSQL "Set Quoted_Identifier On"
	End Function

	'Finds (bookmarks) the record that matches the sFilter, and returns whether found or not
	Public Function FindRS( ByRef rs, ByVal sFilter )
		Call rs.Find(sFilter, 0, adSearchForward, adBookmarkFirst)
		FindRS = not rs.EOF
	End Function

	Public Function UpdateRecords( ByVal sTable, ByVal sFields, ByVal sFilter )
		Dim sSQL
		
		If not sFilter > "" then sFilter = "1 = 0"
		sSQL =	" Update " & sTable & vbCRLF _
			&	" Set " & sFields & vbCRLF _
			&	" Where	" & sFilter
		ExecuteSQL sSQL
	End Function

	Public Function DeleteRecords( ByVal sTable, ByVal sFilter )
		Dim sSQL
		
		If not sFilter > "" then sFilter = "1 = 0"
		sSQL =	" Delete " & sTable & vbCRLF _
			&	" Where	" & sFilter
		ExecuteSQL sSQL
	End Function

	' Assumes cValues is a collection of field name keys and value items
	'TODO: make more robust re: field types, etc.
	Public Function InsertRecords( ByVal sTable, ByVal cValues )
		'If TypeName(cValues) <> "Collection" then exit Function
		If TypeName(cValues) <> "Dictionary" then exit Function
		
		Dim sSQL 
		sSQL _
			= "Insert " & SQLFixUp(sTable) _
			& " ([" & join( cValues.Keys, "], [" ) & "] ) " _
			& "Values ( '" & join( cValues.Items, "', '" ) & "')"

		ExecuteSQL sSQL
	End Function

	Public Function ExecuteSQL( ByVal sSQL )
		' session("Debug") is set to true when logged on through /Includes/Debug.asp
		'If session("Debug") = true then session("sSQL") = sSQL ' primarily for debug purposes
		dim nRowsAffected
		on error resume next
		
		oConn.Execute sSQL, nRowsAffected, adCmdText + adExecuteNoRecords
		
		if err.number <> 0 then
			Response.Write "A Data error has occured: " & err.Description
			Response.Write "Please contact " & Session("Organization") & " support so we can fix this issue as quickly as possible." 
			Session("sSQL") = sSQL
			'TODO: send an email to support/dev
			'lpe sSQL
			Response.End
		end if
		on error goto 0
		ExecuteSQL = nRowsAffected
	End Function


	Private Function CreateCommand(ByVal conDB, ByVal nCommandType)
		Dim oCommand
		
		Set oCommand = Server.CreateObject("ADODB.Command")
		Set oCommand.ActiveConnection = conDB

		oCommand.CommandType = nCommandType
		
		Set CreateCommand = oCommand
	End Function


	Private Function GetADOXCatalog(ByVal oConn)
		Dim oADOXCatalog

		Set oADOXCatalog = CreateObject("ADOX.Catalog")

		oADOXCatalog.ActiveConnection = oConn
		
		Set GetADOXCatalog = oADOXCatalog
	End Function


	Public Function RSFieldsToArray(ByVal rsData, asFields)
		Dim nNumFields
		Dim nField

		nNumFields = rsData.Fields.Count
			
		If nNumFields > 0 Then
			ReDim asFields(nNumFields - 1)

			For nField = 0 To nNumFields - 1
				asFields(nField) = rsData.Fields(nField).Name
			Next
		End If

		RSFieldsToArray = nNumFields
	End Function

	''' Function FieldSet
	''' Takes a recordset and returns a dictionary of the fields and attributes
	Public Function FieldSet( ByRef rs )
		dim fField, dField
		
		'set FieldSet = New Collection
		set FieldSet = Server.CreateObject("Scripting.Dictionary")
		for each fField in rs.Fields
			'set dField = New Collection
			set dField = Server.CreateObject("Scripting.Dictionary")
			dField.Add "name", trim( fField.name )
			dField.Add "type", fField.type
			
			If not FieldSet.Exists( trim( fField.name ) ) then
				FieldSet.Add trim( fField.name ), dField
			End if
		next
	End Function


	Public Function FieldTypeToString(ByVal nType)
		Select Case nType
			Case adTinyInt, adSmallInt, adInteger, adBigInt, _
					adSingle, adDouble, adCurrency, adDecimal, adNumeric
				FieldTypeToString = "Number"

			Case adBoolean
				FieldTypeToString = "Y/N"
			
			Case adDate, adDBDate, adDBTime, adDBTimeStamp
				FieldTypeToString = "Date"

			Case adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
				FieldTypeToString = "Text"

			Case Else
				FieldTypeToString = "UNKNOWN"
		End Select 
	End Function

    Public Function ColumnExists(ByVal sTableName, ByVal sColName)
        ColumnExists = Not IsNull( GetScalar("Select COL_LENGTH('" & SQLFixUp(sTableName) & "', '" & SQLFixUp(sColName) & "')") )
    End Function

End Class 'DBHelper


' Generic database helper functions:
Function SQLFixUp(ByVal sString)
	If IsNull(sString) Then 
		SQLFixUp = ""
		Exit Function
	End If
		
	SQLFixUp = Replace( sString, "'", "''" )
End Function

Function SQLValue(ByVal sValue)
    If IsNull(sValue) then
        SQLValue = "NULL"
    Else
        SQLValue = "'" & SQLFixUp(sValue) & "'"
    End If
End Function

Function IsRS( ByRef rs )
	IsRS = false
	If not IsObject( rs ) then exit function
	If rs is nothing then exit function
	On error resume next
	dim b: b = rs.EOF ' Can this object be treated as a RecordSet?
	if err.number <> 0 then exit function ' ... nope
	On Error goto 0
		
	IsRS = True
End Function

Function DisposeRS( ByRef rs )
	If IsRS( rs ) then
		If rs.State = 1 then ' adStateOpen
			rs.close
		End If
		set rs = nothing
	End If
End Function


%>
