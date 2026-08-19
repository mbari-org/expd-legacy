<%
Function z2Null(varValue)
	' Replace empty strings with
	' Nulls to make database happy.
	If (Len(varValue)=0 or IsNull(varValue)) Then
		Z2Null = "Null"
	Else
		Z2Null = varValue
	End If
End Function

Function String2Null(varValue)
	' Replace empty strings with
	' Nulls to make database happy.
	If (Len(varValue)=0 or IsNull(varValue)) Then
		String2Null = "Null"
	Else
		String2Null = "'" & varValue & "'"
	End If
End Function

Function Number2Null(varValue)
	' Replace empty strings with
	' Nulls to make database happy.
	If (Len(varValue)=0 or IsNull(varValue)) Then
		Number2Null = "Null"
	Else
		Number2Null = varValue
	End If
End Function

Function Boolean2Null(varValue)
	' Replace empty strings with
	' Nulls to make database happy.
	If (Len(varValue)=0 or IsNull(varValue)) Then
		Boolean2Null = "Null"
	Else
	   If varValue = "False" Then
              Boolean2Null = 0
           Else
              Boolean2Null = 1
           End If
	End If
End Function
%>