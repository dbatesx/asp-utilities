<%
Class StringBuilder
	'Ref: http://mital.dk/index.asp?page=3&vpath=3|&snippetCat=2
	'Alt Ref: http://www.u229.no/stuff/snippets/ASPString.asp
	Private strArray()
	Private intGrowRate
	Private intItemCount
	Private Sub Class_Initialize()		
		intGrowRate = 50
		intItemCount = 0
	End Sub
	
	Public Property Get GrowRate
		GrowRate = intGrowRate
	End Property

	Public Property Let GrowRate(value)
		 intGrowRate = value
	End Property

	Public Property Get Length
		Length = len(ToString())
	End Property

	Public Property Get IsEmpty
		IsEmpty = (intItemCount = 0)
	End Property
	
	Public Sub Reset()
		Redim strArray(intGrowRate)
	End Sub

	Public Sub Clear()
		Call Reset()
	End Sub
	
	Public Sub Append(str)
			
		If intItemCount = 0 Then
			Call Reset
		ElseIf intItemCount > UBound(strArray) Then			
			Redim Preserve strArray(Ubound(strArray) + intGrowRate)
		End If
		strArray(intItemCount) = str
		intItemCount = intItemCount + 1
	End Sub	
	
	Public Sub AppendLine(str)
		Call Append(str + vbCRLF)
	End Sub

	Public Function FindString(str)
		Dim x,mx
		mx = intItemCount - 1
		For x = 0 To mx
			If strArray(x) = str Then
				FindString = x
				Exit Function
			End If
		Next
		FindString = -1
	End Function
	
	Public Default Function ToString()
		ToString = ToStringWith("")
	End Function

	Public Function ToString2( ByVal sep )
		ToString2 = ToStringWith(sep)
	End Function
		
	Public Function ToStringWith( ByVal sDelim )
		If intItemCount = 0 Then
			ToStringWith = ""
		Else
			Redim Preserve strArray(intItemCount)
			ToStringWith = Join(strArray, sDelim)
		End If		
	End Function

End Class
%>