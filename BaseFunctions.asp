<%
Public Function IsBool(ByVal oValue)
	If IfNull(oValue, "") = "" Then 
		IsBool = False
		Exit Function
	End If
	IsBool = IsEquiv( oValue, "True" ) Or IsEquiv( oValue, "False" )
End Function

Public Function StrCompI( string1, string2 )
	' StrComp"I" for Ignore Case
	StrCompI = StrComp( string1, string2, vbTextCompare )
End Function

Public Function IsEquiv( ByVal string1, ByVal string2 )
	if IsNull(string1) or IsNull(string2) then 
		IsEquiv = false
		Exit Function
	end if
	IsEquiv = ( 0 = StrCompI( trim(string1), trim(string2) ) )
End Function

Public Function IsInStr( sSearch, sFind )
	IsInStr = 0 < InStr(1, sSearch, sFind, vbTextCompare) 
End Function

Public Function IsChecked( ByVal string1 )
	IsChecked = IsEquiv( string1, "on" )
End Function

Public Function ReplaceText( sSearch, sFind, sReplace )
	ReplaceText = Replace( sSearch, sFind, sReplace, 1, -1, vbTextCompare )
End Function

Public Function ReplaceTokens( ByVal sText, ByVal dTokens )
	Dim sToken
	for each sToken in dTokens.Keys
		sText = replace( sText, "%" & sToken & "%", dTokens(sToken), 1, -1, vbTextCompare )
	next
	ReplaceTokens = sText
End Function


Public Function URLDecode(ByVal sText)
	Dim oRegExpr
	Dim oMatchCollection
	Dim oMatch
	Dim sDecoded
	
	sDecoded = sText
	
	Set oRegExpr = Server.CreateObject("VBScript.RegExp")
	
	oRegExpr.Pattern = "%[0-9,A-F]{2}"
	oRegExpr.Global = True
	
	Set oMatchCollection = oRegExpr.Execute(sText)
	
	For Each oMatch In oMatchCollection
		sDecoded = Replace(sDecoded,oMatch.value,Chr(CInt("&H" & Right(oMatch.Value,2))))
	Next
	
	URLDecode = sDecoded
End function

Function ResponseRedirect(sURL)
	If gbDebug then
		lp "You are about to be redirected to <a href='" & sURL & "'>" & sURL & "</a>"
	Else
		Response.Redirect sURL
	End If
End Function

Function PCase(ByVal strInput)' As String
'Ref: http://www.asp101.com/samples/pcase.asp
    Dim I 'As Integer
    Dim CurrentChar, PrevChar 'As Char
    Dim strOutput 'As String

    PrevChar = ""
    strOutput = ""

    For I = 1 To Len(strInput)
        CurrentChar = Mid(strInput, I, 1)

        Select Case PrevChar
            Case "", " ", ".", "-", ",", """", "'"
                strOutput = strOutput & UCase(CurrentChar)
            Case Else
                strOutput = strOutput & LCase(CurrentChar)
        End Select

        PrevChar = CurrentChar
    Next 'I

    PCase = strOutput
End Function 
    
Function TruncateString( sText, nLength )
	'TruncateString = IIf( len(sText) > nLength, left(sText, nLength - 1) & "...", sText )
	TruncateString = IIf( len(sText) > nLength, left(sText, nLength - 1) & "&hellip;", sText )
End Function


' Returns the value in QueryString if not "", otherwise returns Def
Public Function QueryStrOrDef( ByVal VarName, ByVal Def )
	QueryStrOrDef = Request.QueryString( VarName )
	If Not Len( QueryStrOrDef ) > 0 then QueryStrOrDef = Def
End Function

'Returns value in QueryString if Integer (as Long), otherwise returns Def (as Long)
Public Function QueryIntOrDef( ByVal VarName, ByVal Def )
	Dim sVal
	sVal = Request.QueryString( VarName )
	If Len( sVal ) > 0 And IsNumeric( sVal ) then 
		QueryIntOrDef = CLng( sVal )
	Else
		QueryIntOrDef = CLng( Def )
	End If
End Function

'Returns value in Request (QueryString, Form, or Cookie) if not "", otherwise returns Def
Public Function RequestOrDef( ByVal VarName, ByVal Def )
	RequestOrDef = Request( VarName )
	If Not Len( RequestOrDef ) > 0 then RequestOrDef = Def
End Function

'Returns value in Request if integer (as Long), otherwise returns Def (as Long)
Public Function RequestIntOrDef( ByVal VarName, ByVal Def )
	RequestIntOrDef = Request( VarName )
	If len( RequestIntOrDef ) > 0 And IsNumeric( RequestIntOrDef ) then
		RequestIntOrDef = cLng( RequestIntOrDef )
	else
		RequestIntOrDef = cLng( Def )
	end if
End Function

' Returns value in QueryString or Form (like Request, but excluding Cookies)
Public Function QueryOrFormStr( ByVal VarName )
	QueryOrFormStr = Request.QueryString( VarName )
	If Not Len( QueryStrOrDef ) > 0 then QueryOrFormStr = Request.Form( VarName )
End Function

Public Function QueryOrSession( ByVal VarName, Def )

End Function


Public Function QueryOrCookie( ByVal VarName )
	QueryOrCookie = Request.QueryString( VarName )
	If QueryOrCookie = "" Then QueryOrCookie = Request.Cookies( VarName )
End Function

' Returns Value if present, Def if not.
Public Function StrOrDef( ByVal Value, ByVal Def )
	Value = IfNull( Value, "" )
	If Value = "" then
		StrOrDef = Def
	Else
		StrOrDef = Value
	End If
End Function

' Returns Value if Integer, Def if not.
Public Function IntOrDef( ByVal Value, ByVal Def )
	Value = IfNull( Value, "" )
	If Value = "" or not IsNumeric( Value ) then 
		IntOrDef = Def
	Else
		IntOrDef = CLng( Value )
	End If
End Function

' Returns Value if Boolean, Def if not
Public Function BoolOrDef( ByVal Value, ByVal Def )
	Value = IfNull( Value, False )
	If Value = "" or not IsBool( Value ) then 
		BoolOrDef = Def
	Else
		BoolOrDef = CBool( Value )
	End If
End Function

Public Function DateOrDef( ByVal Value, ByVal Def )
	If IsDate( Value ) then
		DateOrDef = CDate( Value )
	ElseIf IsDate( Def ) then
		DateOrDef = CDate( Def )
	Else
		DateOrDef = null
	End	If
End Function

Public Function TimeOrDef( ByVal Value, ByVal Def )
	If IsDate( Value ) then
		TimeOrDef = TimeValue( Value )
	ElseIf IsDate( Def ) then
		TimeOrDef = TimeValue( Def )
	Else
		TimeOrDef = null
	End	If
End Function

Public Function FormStrOrDef( ByRef Form, ByVal VarName, ByVal Def )
	FormStrOrDef = FormTypeOrDef( Form, VarName, Def, "String" )
End Function

Public Function FormIntOrDef( ByRef Form, ByVal VarName, ByVal Def )
	FormIntOrDef = FormTypeOrDef( Form, VarName, Def, "Int" )
End Function

Public Function FormTypeOrDef( ByRef Form, ByVal VarName, ByVal Def, ByVal VarType )
	FormTypeOrDef = Form( VarName )
	
	if FormTypeOrDef > "" then
	
		Select Case VarType
		
		Case "Boolean", "Bool"
			If IsBool( FormTypeOrDef ) then
				FormTypeOrDef = cBool( FormtypeOrDef )
			Else
				FormTypeOrDef = Def
			End If
			
		Case "Integer", "Int"
			If IsNumeric( FormTypeOrDef ) then
				FormTypeOrDef = cLng( FormtypeOrDef )
			Else
				FormTypeOrDef = Def
			End If
			
		Case Else ' would be "String", and no checking is necessary.
				
		End Select
	else
		FormTypeOrDef = Def
	end if
End Function

Public Function InMethod( sMethodName )
	if gbDebug then Session("Method") = sMethodName
End Function

Public Function IIf( ByVal bConditional, ByVal sTrue, ByVal sFalse )
'	if cBool( bConditional ) then
on error resume next
	if bConditional then
		IIf = sTrue
	else
		IIf = sFalse
	end if
End Function

'LineFeed
Public Function lf( ByVal sText )
	lf = sText & vbCRLF
End Function

'HTML New Line
Public Function br( ByVal sText )
	br = sText & "<br />" & vbCRLF
End Function

'Print/Write to HTTP stream
Public Function p( ByVal sText )
    Response.Write sText
End Function

'HTML Line Print
Public Function lp( ByVal sText ) 'lp for LinePrint
'TODO: check for printable items (ie not objects)
	on error resume next
    Call p( br( sText ) )
End Function

'New-Line Print
Public Function nlp( ByVal sText ) 'nlp for New-Line Print (without <br/>)
    p lf( sText )
End Function

'Line Print and End
Public Function lpe( ByVal sText ) 
	Call lp( sText )
	Response.End
End Function
   
'alias for lp (debug version)
Public Function debug( ByVal sText ) 
	lp sText
End Function

'alias for lpe (ala PERL)
Public Function die( ByVal sText ) 
	lpe sText
End Function

''' Supplied with a recordset, returns a select list of rs(0) as value and rs(1) as display text
Public Function SelectListRS( rs, sName, sClass, sSelected )
	if not rs.EOF then
		SelectListRS = "<select name=""" & sName & """ id=""" & sName & """ class=""" & sClass & """>" & vbCRLF 
		do while not rs.EOF
			SelectListRS = SelectListRS _
				&	"<option value=""" & rs(0) & """" & iif( IsEquiv( rs(0), sSelected ), " selected=""selected""", "" ) & ">" _
				&	rs(1) _
				&	"</option>" & vbCRLF 
			rs.MoveNext
		loop
		SelectListRS = SelectListRS & "</select>"
	end if
End Function
   
''' Supplied with a dictionary, returns a select list of d.Key as value and d.Item as display text
Public Function SelectListDict( d, sName, sClass, sSelected )
	set d = Server.CreateObject("Scripting.Dictionary")
	dim sValue
	if d.Count > 0 then
		SelectListDict = "<select name=""" & sName & """ id=""" & sName & """ class=""" & sClass & """>" & vbCRLF 
		For each sValue in d.Keys
			SelectListDict = SelectListDict _
				&	"<option value=""" & sValue & """" & iif( IsEquiv( sValue, sSelected ), " selected=""selected""", "" ) & ">" _
				&	d(sValue) _
				&	"</option>" & vbCRLF 
		Next
		SelectListDict = SelectListDict & "</select>"
	end if
End Function

'Takes a QueryString formatted string and returns a dictionary of parameters
Public Function QueryStringToDict( sQS )
	Dim d, asQS, asKey, i
	Set d = Server.CreateObject("Scripting.Dictionary")
	If Left(sQS, 1) = "?" then sQS = Right(sQS, Len(sQS) - 1)
	asQS = Split( sQS, "&" )
	For i = 0 to UBound( asQS )
		asKey = Split( asQS(i), "=" )
		If uBound(asKey) > 0 and not d.Exists(asKey(0)) then
			d.Add asKey(0), asKey(1)
		End If
	Next 
	Set QueryStringToDict = d
End Function

'Takes a dictionary of parameters and creates a QueryString formatted string
'(inverse of QueryStringToDict above)
Public Function DictToQueryString( dQS )
	Dim sQS, q
	For each q in dQS.Keys
		sQS = sQS & IIf( sQS = "", "", "&" ) & q & "=" & dQS(q)
	Next
	DictToQueryString = sQS
End Function
   
''' Returns a dictionary of the querystring
Public Function QueryStringDict()
	Dim qs, v
	Set qs = Server.CreateObject("Scripting.Dictionary")
	For each v in Request.QueryString
		If not qs.Exists(v) then
			qs.Add v, Request.QueryString(v)
		End If
	Next
	Set QueryStringDict = qs
End Function

''' Returns true if the key is in the query string.
''' Helps to determine if the key was provided and value is "" vs key wasn't provided.
Public Function QueryStringExists( sKey )
	Dim k
	For each k in Request.QueryString
		if IsEquiv( k, sKey ) then 
			QueryStringExists = True
			Exit Function
		End IF
	Next
	QueryStringExists = False
End Function

''' Returns true if the key is in the Session collection.
''' Helps to determine if the key is in Session and value is "" vs key isn't in Session.
Public Function SessionExists( sKey )
	Dim k
	For each k in Session.Contents
		if IsEquiv( k, sKey ) then 
			SessionExists = True
			Exit Function
		End If
	Next
	SessionExists = False
End Function

Public Sub EnsureDictionary( ByRef dct )
	Select Case TypeName(dct)
	Case "Dictionary"
	Case "String"
		Set dct = QueryStringToDict(dct)
	Case Else
		Set dct = Server.CreateObject("Scripting.Dictionary")
	End Select
	If TypeName(dct) <> "Dictionary" then set dct = Server.CreateObject("Scripting.Dictionary")
End Sub

' Merge keys and values from dSrc into dDest
Public Sub MergeSet( ByRef dDest, ByRef dSrc )
	Dim k
	For Each k in dSrc.Keys
		If not dDest.Exists( k ) then
			dDest.Add k, dSrc(k)
		End If
	Next
End Sub

Public Function RSTable( rs )
	RecordsetTable rs, "", "datatable" 
End Function

Public Function RecordsetTable( rs, sID, sClass )
	dim sField
	' there're probably better defaults than these:
	'if sID = "" then sID = "TableID"
	if sClass = "" then sClass = "datatable" 

%> 
	<table id="<%= sID %>" class="<%= sClass %>">
		<thead>
		  <tr>
<%
	for each sField in rs.Fields
%>			<th><%= sField.name %></th>
<%			
	next
%>		  </tr>
		</thead>
		<tbody>
<%
	'rs.MoveFirst

	
  If rs.state = 1 then
	while not rs.EOF
%>		  <tr>
<%	
		for each sField in rs.Fields
			sField = iif( isnull(sField), "<null>", sField & "&nbsp;" )
			sField = replace( sField, vbCrlf, "<br/>" )
%>			<td><%=sField  %></td>
<%
		next
%>		  </tr>
<%
		rs.MoveNext
	wend
  Else
	nlp "<tr><td>Closed</td></tr>"
  End If 'rs.State
%>		</tbody>
	</table>
<%
End Function

Public Function RecordsetTableEditEx( rs, sID, sClass, aDontEdit )
	dim fField, sFieldValue
	' there're probably better defaults than these:
	'if sID = "" then sID = "TableID"
	'if sClass = "" then sClass = "TableClass" 
	if TypeName(aDontEdit) = "String" then aDontEdit = split( aDontEdit, "," )
	
	if rs.RecordCount > 0 then rs.MoveFirst

%> 
	<table id="<%= sID %>" class="<%= sClass %>">
		<thead>
		  <tr>
<%
	for each fField in rs.Fields
%>			<th><%= fField.name %></th>
<%			
	next
%>		  </tr>
		</thead>
		<tbody>
<%

	while not rs.EOF
%>		  <tr>
<%	
		for each fField in rs.Fields
			sFieldValue = Server.HTMLEncode( fField.Value )
		  if IsInArray( aDontEdit, fField.Name, vbTextCompare ) then
%>			<td><%=sFieldValue  %></td>
<%		  else
			select case fField.type
			'TODO: Add cases for Date, Time, Boolean, Integer, etc.
				case adVarchar:
					if fField.DefinedSize > 50 then
%>			<td><textarea id="<%=fField.Name %>" name="<%=fField.Name %>" cols="15"  rows="1"><%=sFieldValue  %></textarea></td>
<%					else
%>			<td><input type="text" size="12" maxlength="<%=fField.DefinedSize %>" id="<%=fField.Name %>" name="<%=fField.Name %>" value="<%=sFieldValue  %>" /></td>
<%
					end if
				case adInteger:
%>			<td><input type="text" size="8" maxlength="<%=fField.DefinedSize %>" id="<%=fField.Name %>" name="<%=fField.Name %>" value="<%=sFieldValue  %>" /></td>
<%
				
				case else:
%>			<td><input type="text" size="12" maxlength="<%=fField.DefinedSize %>" id="<%=fField.Name %>" name="<%=fField.Name %>" value="<%=sFieldValue  %>" /></td>
<%
				
			end select
		  end if
		next
%>		  </tr>
<%
		rs.MoveNext
	wend
%>		</tbody>
	</table>
<%
End Function

Public Function Ordinal( n )
	Ordinal = n
	n = IntOrDef( n, null )
	If IsNull( n ) then exit function
	Select Case n mod 10
	Case 1
		Ordinal = n & "st"
	Case 2
		Ordinal = n & "nd"
	Case 3
		Ordinal = n & "rd"
	Case Else
		Ordinal = n & "th"
	End Select
End Function

Public Function CamelSpace( sText )
	Dim sBodyBuild
	Dim sChar, sPrevChar, sAddSpace
	Dim nLength
	Dim i
	
	If IsNull(sText) Then 
		CamelSpace = ""
		Exit Function
	End If
	
	sPrevChar = ""
	
	nLength = Len(sText)
	For i = 1 to nLength
		sChar = Mid(sText, i, 1)
		' Insert a space in front of every capital letter
		' unless the previous character is a space or capital letter
		if ucase( sChar ) = sChar _
		and sPrevChar <> "" _
		and	sPrevChar <> " " _
		and	sPrevChar <> "_" _
		and ucase( sPrevChar ) <> sPrevChar then
			sAddSpace = " "
		else 
			sAddSpace = ""
		end if
		sBodyBuild = sBodyBuild & sAddSpace & sChar
		sPrevChar = sChar
	Next

	CamelSpace = sBodyBuild

End Function


Function StripChars( sText )
	sText = Trim(sText)
	StripChars = StripCharsEx( sText, " _-" )
End Function
	
	
Function StripCharsEx( sText, sGoodChars )
	Dim i, sChar
	
	StripCharsEx = ""
	
	'If "_" is allowed, and " " isn't, replace " " with "_"
	If IsInStr( sGoodChars, "_" ) and not IsInStr( sGoodChars, " " ) then
		sText = replace( sText, " ", "_" )
	end if
	
	For i = 1 to len( sText )
		sChar = mid(sText, i, 1)
		if InStr( 1, sGoodChars, sChar ) _
		or	(sChar >= "a" and sChar <= "z") _
		or	(sChar >= "A" and sChar <= "Z") _
		or	(sChar >= "0" and sChar <= "9") then
			StripCharsEx = StripCharsEx & sChar
		end if
	next
	StripCharsEx = trim( StripCharsEx )
End Function
	
	
Function StripHTML(sHTML)
	Dim oRegExp
	Dim sOutput
	
	Set oRegExp = New Regexp

	oRegExp.IgnoreCase = True
	oRegExp.Global = True
	oRegExp.Pattern = "<(.|\n)+?>"

	'Replace all HTML tag matches with the empty string
	sOutput = oRegExp.Replace(sHTML, "")

	'Replace all < and > with &lt; and &gt;
	sOutput = Replace(sOutput, "<", "&lt;")
	sOutput = Replace(sOutput, ">", "&gt;")

	Set oRegExp = Nothing
	
	StripHTML = sOutput  
End Function

	
Function StripScript( sText )
	' Remove all javascript content from the text 
	'	(or anything else that could be deviously used from user submitted content)
	Dim RegEx
	Set RegEx = New RegExp
	'Set RegEx = Server.CreateObject( "Scripting.RexExp" )
	RegEx.IgnoreCase = True
	RegEx.Pattern = "<\s*script.*/\w*\s*>"
	RegEx.Global = True
	
	StripScript = RegEx.Replace( sText, "" )
End Function

Function ArrayIndex( asText, sFind )
	Dim i
	ArrayIndex = -1
	for i = 0 to ubound(asText)
		if IsEquiv( asText(i), sFind ) then
			ArrayIndex = i
			exit for
		end if
	next
End Function

Function IsInArray( asText, sFind )
	IsInArray = false
	Dim i, sText
	sFind = trim(sFind)
	If sFind = "" then exit function 'I think this needs to be here, but needs to be triple checked...
	If not IsArray(asText) then exit function

	for i = 0 to ubound(asText)
		sText = asText(i)
		if StrComp( trim( asText(i) ), sFind, vbTextCompare ) = 0 then
			IsInArray = true
			exit for
		end if
	next
End Function


''' Takes a one dimensional string array and returns a dictionary
Function StringArrayToDictionary( aArray )
'TODO: make it multidimensional capable
	dim sVal
	set StringArrayToDictionary = Server.CreateObject("Scripting.Dictionary")
	for each sVal in aArray
		StringArrayToDictionary.Add trim(sVal), trim(sVal)
	next
End Function

''' Takes a one dimensional string array and returns a collection
Function StringArrayToCollection( aArray )
'TODO: make it multidimensional capable
	dim sVal
	set StringArrayToCollection = new Collection
	for each sVal in aArray
		StringArrayToCollection.Add trim(sVal), trim(sVal)
	next
End Function

''' Take a JSON string and create a dictionary from it.
'	Currently only makes simple dictionaries from simple JSON strings
'	Example: dict( "{key1:item1,key2:item2}" )
Function JsonToDict( sJSON )
'	Function JsonToCollection( sJSON )
'	Function CollectionToJson( cCol )
End Function


Function IfNull( ByVal InValue, ByVal DefValue )
	IfNull = IIf( IsNull( InValue ), DefValue, InValue )
End Function

Function NullIf( ByVal InValue, ByVal NullValue )
	NullIf = IIf( InValue = NullValue, null, InValue )
End Function

Function InitializeFolder( ByVal asFolderpath )
	Dim objFSO
	Dim i
	Dim sContentPath
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	For i = 0 to UBound(asFolderpath)
		if i = 0 then
			sContentPath = Server.MapPath( asFolderpath(0) )
		else
			sContentPath = objFSO.BuildPath( sContentPath, asFolderpath(i) )
			'sContentPath = sContentPath & "\" & asFolderpath(i)
		end if
		
		If Not objFSO.FolderExists(sContentPath) Then
			objFSO.CreateFolder(sContentPath) 
		End If
	Next
	
	InitializeFolder = sContentPath
	Set objFSO = Nothing
End Function

' Ensures the content folder(s) exist, both in /content and in /contentdeploy.
' Just provide the folderpath beyond /content/
' Returns the MapPath of the folderpath
Function EnsureContentFolder( ByVal asFolderpath )
	Dim objFSO
	Dim i, sFolder
	Dim sContentPath

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	For each sFolder in Array( CONTENT_DIR, STAGING_DIR )
		sContentPath = Server.MapPath( sFolder )
		For i = 0 to UBound(asFolderpath)
			sContentPath = objFSO.BuildPath( sContentPath, asFolderpath(i) )
			
			If Not objFSO.FolderExists(sContentPath) Then
				objFSO.CreateFolder(sContentPath) 
			End If
		Next
		If sFolder = CONTENT_DIR then EnsureContentFolder = sContentPath
	Next
	Set objFSO = Nothing
End Function

Function FileExists( sFilePath )
	'TODO: Is there another similar function?
	Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
	FileExists = fso.FileExists( ServerMapPath( sFilePath ) )
End Function

' Server.MapPath has a known bug(?) where an error is thrown if a file name has a comma or quote/apostrophe in it
' (ref: 
'	http://msdn.microsoft.com/en-us/library/ms524632
'	http://classicasp.aspfaq.com/files/directories-fso/why-do-i-get-an-invalid-path-character-error.html)
Function ServerMapPath( ByVal sFilePath )
	'sFilePath is assumed to be a relative file path
	If IsInStr(sFilePath, "\") then ' it's likely already been mapped
		ServerMapPath = sFilePath
	Else
		ServerMapPath = Server.MapPath("/") & IIf( left(sFilePath, 1) = "/", "", "/" ) & replace(sFilePath, "/", "\" )
	End If
End Function

Function ReportError( sInfo )
	ReportError = "<br/>An error occurred: " & err.Description & "<br/>" & vbCRLF _
		&	IIf( len(sInfo) > 0, sInfo & "<br/>" & vbCRLF, "" ) _
		&	"Please notify Support so we can address this issue as soon as we can.<br/>" & vbCRLF
End Function

Function ToString( var )
	Dim sTypeName
	sTypeName = TypeName( var )
	Select Case sTypeName
	Case "String"
		ToString =  var
	Case "Integer", "Long", "Single", "Double"
		ToString = CStr( var )
	Case Else
		ToString = "(Undetermined)"
	End Select
	ToString = "[" & sTypeName & "] " & ToString
End Function

Function RemoteLink( ByVal sURL )
	If not IsEquiv( Left( sURL, 4 ), "http" ) then
		RemoteLink = "http://" & sURL
	Else	
		RemoteLink = sURL
	End If
End Function

Public Function alert( ByVal sText )
	nlp	"<script type='text/javascript'>"
	nlp	"	alert('" & sText & "');"
	nlp	"</script>"
End Function

''' Ensures an Item is (or Items are) in a comma-space separated string of Items
Public Function AddItem( ByVal sList, ByVal sItem )
	' sItem may be a comma-space separated set of values
	Dim asItems, sSubItem, sListEx
	asItems = Split( sItem, ", " )
	sListEx = ", " & sList & ", "
	For each sSubItem in asItems
		If not IsInStr( sListEx, ", " & sSubItem & ", " ) then
			sList = sList & IIf( sList = "", "", ", " ) & sSubItem
		End If
	Next
	AddItem = sList
End Function

''' Ensures an Item is (or Items are) NOT in a comma-space separated string of Items
Public Function RemoveItem( ByVal sList, ByVal sItem )
	' sItem may be a comma-space separated set of values
	Dim asItems, sListEx, sItemEx, sSubItem
	asItems = Split( sItem, ", " )
	sListEx = ", " & sList & ", "	
	For each sSubItem in asItems
		sItemEx = ", " & sSubItem & ", "		
		If IsInStr( sListEx, sItemEx ) then
			sListEx = Replace( sListEx, sItemEx, ", " ) 
			sList = mid( sListEx, 3, IIf( len(sListEx) < 4, 0, len(sListEx) - 4 ) )
		End If
	Next
	RemoveItem = sList
End Function

''' Returns true if there is a common (comma-space separated) item between lists
Function IsInCommaList( ByVal sList1, ByVal sList2 )
	IsInCommaList = false
	If trim(sList1) = "" or trim(sList2) = "" then exit function

	Dim aList1 : aList1 = split( sList1, ", ")
	Dim aList2 : aList2 = split( sList2, ", ")
	Dim sItem
	For each sItem in aList2
		If IsTextInArray(aList1, sItem) then
			IsInCommaList = true
			Exit For
		End If
	Next
End Function

Function IsValidEmailAddress( sEmail )
	'Ref: http://www.linuxjournal.com/article/9585?page=0,1
	'1. An e-mail address consists of local part and domain separated by an at sign (@) character (RFC 2822 3.4.1).
	'2. The local part may consist of alphabetic and numeric characters, 
	'	and the following characters: !, #, $, %, &, ', *, +, -, /, =, ?, ^, _, `, {, |, } and ~, 
	'	possibly with dot separators (.), inside, but not at the start, end or next to another dot separator (RFC 2822 3.2.4).
	'3. The local part may consist of a quoted string—that is, anything within quotes ("), including spaces (RFC 2822 3.2.5).
	'4. Quoted pairs (such as \@) are valid components of a local part, though an obsolete form from RFC 822 (RFC 2822 4.4).
	'5. The maximum length of a local part is 64 characters (RFC 2821 4.5.3.1).
	'6. A domain consists of labels separated by dot separators (RFC1035 2.3.1).
	'7. Domain labels start with an alphabetic character followed by zero or more alphabetic characters, 
	'	numeric characters or the hyphen (-), ending with an alphabetic or numeric character (RFC 1035 2.3.1).
	'8. The maximum length of a label is 63 characters (RFC 1035 2.3.1).
	'9. The maximum length of a domain is 255 characters (RFC 2821 4.5.3.1).
	'10 The domain must be fully qualified and resolvable to a type A or type MX DNS address record (RFC 2821 3.6).	
	'	Domain labels start with an alphabetic character followed by zero or more alphabetic characters, 
	'	numeric characters or the hyphen (-), ending with an alphabetic or numeric character (RFC 1035 2.3.1)

    Dim regEx	: Set regEx = New RegExp 
    regEx.IgnoreCase = true
	IsValidEmailAddress = False
	Dim aParts	: aParts = split(sEmail, "@")
	Dim nParts	: nParts = UBound(aParts)
	If nParts < 1 then Exit Function ' no @ sign
	
	'Checking the Domain
	Dim sDomain: sDomain = aParts(nParts) 'get the last part
	If sDomain = "" or len(sDomain) > 255 then Exit Function ' can't be longer than 255 chars
    regEx.Pattern = "[^A-Z0-9\-\.]" 
	If regEx.Test(sDomain) then Exit Function 'Only alpha-numeric and dots allowed in Domain (not allowing comments)
	If left(sDomain, 1) = "." or right(sDomain,1) = "." then Exit Function 'can't start or end with a dot
	If IsInStr(sDomain, "..") then Exit Function 'can't have two adjacent dots
	If not IsInStr(sDomain, ".") then Exit Function 'needs at least one dot
	Dim aDomain: aDomain = split(sDomain, ".")
	If UBound(aDomain) = 0 then Exit Function
	
	'Checking the Local Address
	ReDim Preserve aParts(nParts - 1)
	Dim sLocal: sLocal = join(aParts,"@") ' reconstruct the entire Local part
	If sLocal = "" or len( sLocal ) > 63 then Exit Function
	If left(sLocal, 1) = "." or right(sLocal,1) = "." then Exit Function 'can't start or end with a dot
	If IsInStr(sLocal, "..") then Exit Function 'can't have two adjacent dots
	If left(sLocal, 1) = """" and right(slocal, 1) = """" then 'quoted local part
		If IsInStr(mid(sLocal, 2, len(sLocal) - 2), """") then Exit Function 'embedded quotes not allowed unless escaped
	Else
		If IsInStr(sLocal, """") then Exit Function 'Quotes otherwise not allowed unless escaped
		If IsInStr(sLocal, " ") then Exit Function 'spaces not allowed unless escaped
		If IsInStr(sLocal, "@") and not IsInStr(sLocal, "\@") then Exit Function ' @ needs to be escaped
	End If
	
	IsValidEmailAddress = true
End Function

'Basically a specialized Query String wrapper for Collection
Class QueryString
	private m_qs

	Sub Class_Initialize
		Dim key
		set m_qs = new Collection
		For each key in Request.QueryString
			Call m_qs.Add(key, Request.QueryString(key))
		Next
	End Sub

	Function IncludeForm()
		Dim key
		For each key in Request.Form
			Call m_qs.Add(key, Request.Form(key))
		Next
		Set IncludeForm = me
	End Function

	Function Add(sKey, sValue)
		Call m_qs.Add(sKey, sValue)
		Set Add = me
	End Function

	Function Remove(sKey)
		Call m_qs.Remove(sKey)
		Set Remove = me
	End Function

	Function Clear()
		Call m_qs.Clear()
		Set Clear = me
	End Function

	Function ToString()
		'ToString = m_qs.ToStringWith("&")
		Dim sQS, sKey
		For each sKey in m_QS.Keys
			sQS = sQS & IIf( sQS = "", "", "&" ) & sKey & "=" & Server.URLEncode( m_QS(sKey) )
		Next
		ToString = sQS
	End Function
End Class

%>

