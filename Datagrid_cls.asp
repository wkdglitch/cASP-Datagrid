<%
'***********************************************************************************************
' Class: DataGrid (see ReadMe.txt)
' Author: Fernando Herrera 
'			Based on the DataGrid Class of Nick DelMedico (www.pixel420.com)
'
' Properties:
'		.Connection			- The connection object to use if one exists (default is none)
'		.Command			- The SQL query to build the grid from (Required)
'		.Recordset			- The Recordset to build the grid from if not using .Command (Required)
'		.AutoGenerateCols	- Boolean specifying if the grid creates the columns (default is "True")
'		.GridClassName		- Name of a CSS Class to apply to the grid (default is none)
'		.GridStyle			- CSS inline style to apply to the grid (default is none)
'		.GridAlign			- Alignment of grid ("center", "left", "right" - default is none)
'		.HeaderClassName	- Name of a CSS Class to apply to the grid header (default is none)
'		.HeaderStyle		- CSS inline style to apply to the grid header (default is none)
'		.ItemClassName		- Name of a CSS Class to apply to the records/rows (default is none)
'		.ItemStyle			- CSS inline style to apply to the records/rows (default is none)
'		.AltItemClassName	- Name of a CSS Class to apply to every other record/row (default is none)
'		.AltItemStyle		- CSS inline style to apply to the every other record/row (default is none)
'		.LinkClassName		- Name of a CSS Class to apply to links added by .LinkColumn() (default is none)
'		.LinkStyle			- CSS inline style to apply to links added by .LinkColumn() (default is none)
'		.FooterClassName	- Name of a CSS Class to apply to the grid paging footer (default is none)
'		.FooterStyle		- CSS inline style to apply to the grid paging footer (default is none)
'		.PagingLinkClass	- Name of a CSS Class to apply to Next X/Previous X links (default is none)
'		.PageResults		- Boolean whether to page the results (default is "False")
'		.CurrentPage		- Int that represents which page to start from when paging (default is "1")
'		.PageSize			- Int that represents how many records shown per page when paging (default is "10")
'		.SortColumn			- The name of the db field to sort by.
'		.SortType			- The way to sort the datagrid (ASC/DESC).
'		.SortImgUp			- Image file displayed for sort type ASC
'		.SortImgDn			- Image file displayed for sort type DESC
'		.NavBar				- Boolean specifying if the navigation bar is displayed (default is "True")
'		.FullRowSelect		- Boolean enabling full row selection. (requires .FullRowLink to be set. default is "False")
'		.FullRowSelectColor	- Color of hilight row.
'		.FullRowLink		- URL string for full row select.
'       .SumColumns         - Boolean specifying if totals are kept for the datagrid (default is "false")
'       .SumColumnsName     - List of Columns to keep a Total On.  (column1,column2,etc)
'		.DateFormat			- Format for date fields (0 = vbGeneralDate - Default. Returns date: mm/dd/yy and time if specified: hh:mm:ss PM/AM. 1 = vbLongDate - Returns date: weekday, monthname, year 2 = vbShortDate - Returns date: mm/dd/yy 3 = vbLongTime - Returns time: hh:mm:ss PM/AM 4 = vbShortTime - Return time: hh:mm)
'
' Methods:
'		.CreateConnection	- Creates a connection. (Args: Connection string)
'		.SetTableOptions	- Overrides default grid settings. (Args: Width, Padding, Spacing, Border)
'		.AddColumn			- Replaces the db column name w/ one you specify and creates "sort by" headers
'								-->(Args: [DB Field], [Header Text], [Boolean is sort by header], CSS class name for sort by links)
'		.ReturnDataGrid		- Outputs DataGrid as a Text String. (Args: None)
'		.Bind				- Outputs DataGrid. (Args: None)
'***********************************************************************************************

Class DataGrid

'===[ Private Variables ]=====================================================
Private p_objConn
Private p_objRS
Private p_strQS
Private p_strCommand

Private p_blnPrivConn
Private p_blnAutoCols
Private p_blnAllowPaging
Private p_blnLinkCol
Private p_blnSortCol
Private p_blnNavBar
Private p_blnNavForm
Private p_blnSumColumns

Private p_arrBoundCols
Private p_arrDisplayCols
Private p_arrRecords
Private p_arrLinks
Private p_arrTemp

Private p_strSortName
Private p_strSortType
Private p_strSortImgUp
Private p_strSortImgDn
Private p_intAbsPage
Private p_intPageSize

Private p_strOutput
Private p_strBoundCols
Private p_strDisplayCols
Private p_strLinks
Private p_strScriptName
Private p_intCurRow
Private p_intEndRec
Private p_intColCount
Private p_i
Private p_x

Private p_strHeader
Private p_intPadding
Private p_intSpacing
Private p_intGridBorder
Private p_strGridAlign
Private p_strGridStyle
Private p_strGridClass
Private p_strGridWidth
Private p_strHeadStyle
Private p_strHeadClass
Private p_strItemStyle
Private p_strItemClass
Private p_strAltItemStyle
Private p_strAltItemClass
Private p_strThisStyle
Private p_strThisClass
Private p_strPagingClass
Private p_strFootStyle
Private p_strFootClass
Private p_strLinkStyle
Private p_strLinkClass

Private p_blnFullRowSelect
Private p_strFullRowSelectColor
Private p_strFullRowLink

Private p_strSumColumnsName
Private p_strDateFormat



'===[ Properties ]============================================================
Public Property Set Connection(oConn)
	Set p_objConn = oConn
End Property

Public Property Let Command(sSql)
	p_strCommand = sSql
End Property

Public Property Let Recordset(oRs)
	Set p_objRS = Server.CreateObject("ADODB.Recordset")
	Set p_objRS = oRs
End Property

Public Property Let AutoGenerateCols(bAutoCols)
	If bAutoCols = True Or bAutoCols = False Then
		p_blnAutoCols = bAutoCols
	End If
End Property

Public Property Let GridClassName(sClass)
	p_strGridClass = " class=""" & sClass & """"
End Property

Public Property Let GridStyle(sStyle)
	p_strGridStyle = " style=""" & sStyle & """"
End Property

Public Property Let GridAlign(sAlign)
	Select Case LCase(sAlign)
		Case "center", "left", "right"
		p_strGridAlign = " align=""" & sAlign & """"
	End Select
End Property

Public Property Let HeaderClassName(sClass)
	p_strHeadClass = " class=""" & sClass & """"
End Property

Public Property Let HeaderStyle(sStyle)
	p_strHeadStyle = " style=""" & sStyle & """"
End Property

Public Property Let ItemClassName(sClass)
	p_strItemClass = " class=""" & sClass & """"
End Property

Public Property Let ItemStyle(sStyle)
	p_strItemStyle = " style=""" & sStyle & """"
End Property

Public Property Let AltItemClassName(sClass)
	p_strAltItemClass = " class=""" & sClass & """"
End Property

Public Property Let AltItemStyle(sStyle)
	p_strAltItemStyle = " style=""" & sStyle & """"
End Property

Public Property Let LinkClassName(sClass)
	p_strLinkClass = " class=""" & sClass & """"
End Property

Public Property Let LinkStyle(sStyle)
	p_strLinkStyle = " style=""" & sStyle & """"
End Property

Public Property Let FooterClassName(sClass)
	p_strFootClass = " class=""" & sClass & """"
End Property

Public Property Let FooterStyle(sStyle)
	p_strFootStyle = " style=""" & sStyle & """"
End Property

Public Property Let PagingLinkClass(sClass)
	p_strPagingClass = " class=""" & sClass & """"
End Property

Public Property Let PageResults(bPaging)
	If bPaging = True Or bPaging = False Then
		p_blnAllowPaging = bPaging
	End If
End Property

Public Property Let CurrentPage(iAbsPage)
	If IsNumeric(iAbsPage) And (iAbsPage <> 0) Then p_intAbsPage = CInt(iAbsPage)
End Property

Public Property Let PageSize(iPageSize)
	If IsNumeric(iPageSize) And (iPageSize <> 0) Then p_intPageSize = CInt(iPageSize)
End Property

Public Property Let SortColumn(sSortName)
	p_strSortName = sSortName
End Property

Public Property Let SortType(sSortType)
	p_strSortType = sSortType
End Property

Public Property Let SortImgUp(sSortImg)
	p_strSortImgUp = sSortImg
End Property

Public Property Let SortImgDn(sSortImg)
	p_strSortImgDn = sSortImg
End Property

Public Property Let NavBar(bNavBar)
	If bNavBar = True Or bNavBar = False Then
		p_blnNavBar = bNavBar
	End If
End Property

Public Property Let NavForm(bNavForm)
	If bNavForm = True or bNavForm = False Then
		p_blnNavForm = bNavForm
	End If
End Property

Public Property Let FullRowSelect(bFullRowSelect)
	If bFullRowSelect = True or bFullRowSelect = False Then
		p_blnFullRowSelect = bFullRowSelect
	End If
End Property

Public Property Let FullRowLink(sFullRowLink)
	p_strFullRowLink = sFullRowLink
End Property

Public Property Let FullRowSelectColor(sColor)
	p_strFullRowSelectColor = sColor
End Property

Public Property Let SumColumns(blnSumColumns)
    p_blnSumColumns = blnSumColumns
End Property

Public Property Let SumColumnsNames(sSumColumnsName)
    p_strSumColumnsName = sSumcolumnsName
End Property

Public Property Let DateFormat(sFormat)
    p_strDateFormat = sFormat
End Property


'===[ Private Functions ]=====================================================
'*****************************************************************************
' This fires when an instance of the class is created
'*****************************************************************************
Private Sub Class_Initialize()
	p_strScriptName = Request.ServerVariables("SCRIPT_NAME")
	p_strQS = Request("QUERY_STRING")
	p_strGridWidth = "100%" 
	p_strGridAlign = ""
	p_strGridClass = "" 
	p_strGridStyle = ""
	p_strHeadClass = "" 
	p_strHeadStyle = ""
	p_strItemClass = "" 
	p_strItemStyle = ""
	p_strFootClass = "" 
	p_strFootStyle = ""
	p_strLinkClass = "" 
	p_strLinkStyle = ""
	p_strAltItemClass = "" 
	p_strAltItemStyle = ""
	p_strPagingClass = "" 
	p_strLinks = ""
	p_intGridBorder = 1 
	p_intCurRow = 1
	p_intPadding = 2 
	p_intSpacing = 2
	p_intPageSize = 10 
	p_intAbsPage = 1
	p_blnAutoCols = True 
	p_blnPrivConn = False
	p_blnAllowPaging = False 
	p_blnLinkCol = False
	p_blnSortCol = False 
	p_strSortType = ""
	p_strSortImgUp = "" 
	p_strSortImgDn = ""
	p_intColCount = 0 
	p_blnNavBar = True
	p_blnNavForm = False
	p_blnFullRowSelect = False
	p_strFullRowLink = ""
	p_strFullRowSelectColor = "#dcfac9"
	p_blnSumColumns = False
	p_strDateFormat = 1
End Sub


'*****************************************************************************
' This fires when the instance of the class is destroyed
'*****************************************************************************
Private Sub Class_Terminate()
	p_strHeader = Empty
	p_strOutput = Empty
	If p_blnPrivConn Then 
		'p_objRS.Close
		'p_objConn.Close
		'Set p_objRS = Nothing
		'Set p_objConn = Nothing
	End If
End Sub


Private Function ConvertToString(ByVal vVar)
	If Not (IsNull(vVar) Or IsEmpty(vVar) Or vVar = "") Then		
		ConvertToString = CStr(vVar)
	Else
		ConvertToString = ""
	End If	
End Function


Private Function DeleteFromQueryString(sQueryString, sParam)
	Dim iStartIndex, iEndIndex
	Dim sString
	iStartIndex = InStr(1, sQueryString, sParam, vbTextCompare)
	If iStartIndex <> 0 Then
		iEndIndex = InStr(iStartIndex, sQueryString, "&", vbTextCompare)
		If iEndIndex = 0 Then iEndIndex = Len(sQueryString)
		sQueryString = Replace(sQueryString, Mid(sQueryString, iStartIndex, (iEndIndex - iStartIndex + 1)), "", 1, -1, vbTextCompare)
	End If
	DeleteFromQueryString = sQueryString
End Function


Private Function Nav_Post(iPageSize, iPageCount, iAbsPage, [iRecordCount], [sNav])
	Dim p_sTemp, p_iCounter, p_iCounterStart, p_iCounterEnd, p_sNav
	If Len(Trim(sNav)) > 0 Then
		p_sNav = DeleteFromQueryString(sNav, "page")
		p_sNav = DeleteFromQueryString(p_sNav, "list")
	End If
	p_sTemp = "<tr><td" & p_strFootClass & p_strFootStyle & " colspan=""" & p_intColCount & """>"
	p_sTemp = p_sTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr><td align=""left"">"
	p_sTemp = p_sTemp & "&nbsp;<b>" & iRecordCount & "</b> Records Found - Displaying Page "& iAbsPage &" of "& iPageCount &"</td>"
				
	If p_blnAllowPaging Then
		p_sTemp = p_sTemp & "<td align=""right"">"
					
		'# Write out the "|< First"
		If iAbsPage <> 1 Then
			p_sTemp = p_sTemp & "<button name=""firstpage"" onclick=""submitForm(1, "& iPageSize &",'"& p_strSortName &"','"& p_strSortType &"');"">|&lt; First</button>"
		End If
					
		'# Write out the "<< Prev"
		If iAbsPage <> 1 Then
			p_sTemp = p_sTemp & "<button name=""prevpage"" onclick=""submitForm("& (iAbsPage - 1) &", "& iPageSize &",'"& p_strSortName &"','"& p_strSortType &"');"">&lt;&lt; Prev</button>"
		End If
					
		'# Write out the "Next >>"
		If Cint(iAbsPage) <> Cint(iPageCount) Then
			p_sTemp = p_sTemp & "<button name=""nextpage"" onclick=""submitForm("& (iAbsPage + 1) &", "& iPageSize &",'"& p_strSortName &"','"& p_strSortType &"');"">Next &gt;&gt;</button>"
		End If
					
		'# Write out the "Last >|"
		If Cint(iAbsPage) <> Cint(iPageCount) Then
			p_sTemp = p_sTemp & "<button name=""lastpage"" onclick=""submitForm("& iPageCount &", "& iPageSize &",'"& p_strSortName &"','"& p_strSortType &"');"">Last &gt;|</button>"
		End If
		
		'# Write out the Numbers
		p_sTemp = p_sTemp & "&nbsp; Go to: <select name=""spage"" onchange=""javascript:this.submit();"">"
		For p_iCounter = 1 To iPageCount
			If Cint(p_iCounter) <> Cint(iAbsPage) Then
				p_sTemp = p_sTemp & "<option value=""" & p_iCounter &""">Page " & p_iCounter & "</option>"
			Else
				p_sTemp = p_sTemp & "<option value=""" & p_iCounter &""" selected>Page " & p_iCounter & "</option>"
			End If
		Next
		p_sTemp = p_sTemp & "</select>"		
					
		p_sTemp = p_sTemp & "</td>"
	End If
	p_sTemp = p_sTemp & "</tr></table></td></tr>"
	Nav_Post = p_sTemp
End Function


Private Function Nav_Get(iPageSize, iPageCount, iAbsPage, [iRecordCount], [sNav])
	Dim p_sTemp, p_iCounter, p_iCounterStart, p_iCounterEnd, p_sNav
	If Len(Trim(sNav)) > 0 Then
		p_sNav = DeleteFromQueryString(sNav, "page")
		p_sNav = DeleteFromQueryString(p_sNav, "list")
	End If
	p_sTemp = "<tr><td" & p_strFootClass & p_strFootStyle & " colspan=""" & p_intColCount & """>"
	p_sTemp = p_sTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr><td align=""left"">"
	p_sTemp = p_sTemp & "&nbsp;<b>" & iRecordCount & "</b> Records Found - Displaying Page "& iAbsPage &" of "& iPageCount &"</td>"
					
	If p_blnAllowPaging Then
		p_sTemp = p_sTemp & "<td align=""right"">"
						
		'Begin Numbering System
		If (iAbsPage Mod 10) = 0 Then
			p_iCounterStart = iAbsPage - 9
		Else
			p_iCounterStart = iAbsPage - (iAbsPage Mod 10) + 1
		End If
		p_iCounterEnd = p_iCounterStart + 9
		If p_iCounterEnd > iPageCount Then p_iCounterEnd = iPageCount
						
		'# Write out the "|< First"
		If iAbsPage <> 1 Then
			p_sTemp = p_sTemp & "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & _
								"?page=1&list=" & iPageSize _
								& "&" & p_sNav & """" & p_strPagingClass & ">[|&lt;]</a>&nbsp;"
		End If
						
		'# Write out the << Previous Code
		If iAbsPage <> 1 Then
			p_sTemp = p_sTemp & "<a href='" & Request.ServerVariables("SCRIPT_NAME") & _
								"?page=" & (iAbsPage - 1) & "&list=" & iPageSize _
								& "&" & p_sNav & "'" & p_strPagingClass & ">[&lt;&lt;]</a>&nbsp;"
		End if
						
		'# Write out the Numbers
		For p_iCounter = p_iCounterStart To p_iCounterEnd
			If Cint(p_iCounter) <> Cint(iAbsPage) Then
				p_sTemp = p_sTemp & "<a href='" & Request.ServerVariables("SCRIPT_NAME") _
									& "?page=" & p_iCounter & "&list=" & iPageSize & "&" & p_sNav & "'" & p_strPagingClass & ">" & " " & p_iCounter & "</a>"
			Else
				p_sTemp = p_sTemp & "<b>" & " " & p_iCounter & "</b>"
			End If
			If p_iCounter <> p_iCounterEnd Then p_sTemp = p_sTemp & " "
		Next
						
		'# Write out the >> Pages            
		If Cint(iAbsPage) <> Cint(iPageCount) Then
		p_sTemp = p_sTemp & "&nbsp;<a href='" & Request.ServerVariables("SCRIPT_NAME") _
							& "?page=" & (iAbsPage + 1) & "&list=" & iPageSize & "&" & p_sNav _
							& "'" & p_strPagingClass & ">[&gt;&gt;]</a>"               
		End if

		'# Write out the "Last >|"
		If Cint(iAbsPage) <> Cint(iPageCount) Then
		p_sTemp = p_sTemp & "&nbsp;<a href=""" & Request.ServerVariables("SCRIPT_NAME") _
							& "?page=" & iPageCount & "&list=" & iPageSize & "&" & p_sNav _
							& """" & p_strPagingClass & ">[&gt;|]</a>"               
		End if
						
		p_sTemp = p_sTemp & "</td>"
	End If
	p_sTemp = p_sTemp & "</tr></table></td></tr>"
	Nav_Get = p_sTemp
End Function


Private Function CreateNav(iPageSize, iPageCount, iAbsPage, [iRecordCount], [sNav])
	Dim p_sTemp
	If p_blnNavBar Then
		If p_blnNavForm Then
			p_sTemp = Nav_Post(iPageSize, iPageCount, iAbsPage, iRecordCount, sNav)
		Else
			p_sTemp = Nav_Get(iPageSize, iPageCount, iAbsPage, iRecordCount, sNav)
		End If
	End If
	CreateNav = p_sTemp
End Function


Private Function ReturnNextStyle(iCurRow)
	If iCurRow Mod 2 = 0 Then
		If p_strAltItemStyle = "" Then
			ReturnNextStyle = p_strItemStyle
		Else
			ReturnNextStyle = p_strAltItemStyle
		End If
	Else
		ReturnNextStyle = p_strItemStyle
	End If
End Function


Private Function ReturnNextClass(iCurRow)
	If iCurRow Mod 2 = 0 Then
		If p_strAltItemClass = "" Then
			ReturnNextClass = p_strItemClass
		Else
			ReturnNextClass = p_strAltItemClass
		End If
	Else
		ReturnNextClass = p_strItemClass
	End If
End Function


Private Function FormatTemplate(ByRef objFields, ByVal sTemplate)
	Dim objField, sTemp
	sTemp = sTemplate
	For Each objField In objFields
		Select Case objField.Type
		   Case adDBTimeStamp
				If IsEmpty(ObjFields(objField.Name).Value) Or IsNull(ObjFields(ObjField.Name).Value) Or ObjFields(ObjField.Name).Value = vbNullString Then
					sTemp = Replace(sTemp, "#" & objField.Name & "#", ConvertToString(objFields(objField.Name).Value))
				 Else
					sTemp = Replace(sTemp, "#" & objField.Name & "#", FormatDateTime(objFields(objField.Name).Value, p_strDateFormat))
				End If
					
			Case adCurrency
				If IsEmpty(ObjFields(objField.Name).Value) Or IsNull(ObjFields(ObjField.Name).Value) Or ObjFields(ObjField.Name).Value = vbNullString Then
					sTemp = Replace(sTemp, "#" & objField.Name & "#", ConvertToString(objFields(objField.Name).Value))
				 Else
					sTemp = Replace(sTemp, "#" & objField.Name & "#", FormatCurrency(objFields(objField.Name).Value))
				End If
				
				
			Case Else
				sTemp = Replace(sTemp, "#" & objField.Name & "#", ConvertToString(Trim(objFields(objField.Name).Value)))
		End Select
		If ConvertToString(sTemp) = "" Then sTemp = "&nbsp;"
	Next
	FormatTemplate = sTemp
End Function


Private Sub SetFilter(ByRef objRS, ByVal lPageSize, ByVal iAbsolutePage, ByVal sSortName, ByVal sSortType)	
	If sSortName <> "" Then objRS.Sort = "[" & sSortName & "] " & sSortType
	If CLng(lPageSize) <> -1 Then
		objRS.PageSize = CLng(lPageSize)
		objRS.CacheSize = objRS.PageSize
		objRS.AbsolutePage = iAbsolutePage
	Else
		objRS.MoveFirst
	End If	
End Sub


Private Function CreateFullRowSelectScript()
	Dim sTemp
	If p_blnFullRowSelect Then
		sTemp = sTemp & "<script type=""text/javascript"">" & vbCrLf & _
			"function ChangeColor(tableRow, highLight){" & vbCrLf & _
			"	if (highLight){" & vbCrLf & _
			"		tableRow.style.backgroundColor = '" & p_strFullRowSelectColor & "';" & vbCrLf & _
			"	} else {" & vbCrLf & _
			"		tableRow.style.backgroundColor = '';" & vbCrLf & _
			"	}" & vbCrLf & _
			"}" & vbCrLf & _
			vbCrlf & _
			"function DoNav(theUrl){" & vbCrLf & _
			"	document.location.href = theUrl;" & vbCrLf & _
			"}" & vbCrLf & _
			"</script>" & vbCrLf
		CreateFullRowSelectScript = sTemp
	End If
End Function



'===[ Methods ]===============================================================
	Public Sub CreateConnection(sConn)
		Set p_objConn = Server.CreateObject("ADODB.Connection")
		p_objConn.Open sConn
		p_blnPrivConn = True
	End Sub


	Public Sub SetTableOptions(sWidth, iSpace, iPad, iBorder)
		If IsNumeric(sWidth) Or Right(sWidth, 1) = "%" Then p_strGridWidth = sWidth
		If IsNumeric(iSpace) Then p_intPadding = iSpace
		If IsNumeric(iPad) Then p_intSpacing = iPad
		If IsNumeric(iBorder) Then p_intGridBorder = iBorder
	End Sub


	Public Sub AddColumn(sColumn, sHeader, bSort, sTemplate)
		Dim p_sColumn, p_sSort, p_sSortType, p_sNav, p_sHClass
		If Len(Trim(p_strQS)) > 0 Then
			p_sNav = p_strQS
			p_sNav = DeleteFromQueryString(p_sNav, "sort")
			p_sNav = DeleteFromQueryString(p_sNav, "stype")
		End If
		If bSort Then
			If p_blnNavForm Then
				If (p_strSortType = "desc") or (p_strSortType = "")Then
					p_sSortType = "onclick=""submitForm("& p_intAbsPage &", "& p_intPageSize &",'"& sColumn &"','asc');"""
				Else
					p_sSortType = "onclick=""submitForm("& p_intAbsPage &", "& p_intPageSize &",'"& sColumn &"','desc');"""
				End If
				p_sSort = "<a href=""#"" "& p_sSortType &" title=""Sort By " & sHeader & """" & p_sHClass & ">" & sHeader & "</a>"
				If (Len(Trim(p_strSortImgUp)) > 0) And (Len(Trim(p_strSortImgDn)) > 0) And (p_strSortName = sColumn) Then					If p_strSortType = "asc" Then
						p_sSort = p_sSort & "&nbsp;<img src=""" & p_strSortImgUp & """ alt=""Sort Ascending"">&nbsp;"					Else
						p_sSort = p_sSort & "&nbsp;<img src=""" & p_strSortImgDn & """ alt=""Sort Descending"">&nbsp;"					End If
				End If
				p_blnSortCol = True
			Else
				If (p_strSortType = "desc") or (p_strSortType = "") Then					p_sSortType = "&stype=asc&" & p_sNav				Else					p_sSortType = "&stype=desc&" & p_sNav				End If
				p_sSort = "<a href=""" & p_strScriptName & "?sort=" & sColumn & p_sSortType &""" title=""Sort By " & sHeader & """" & p_sHClass & ">" & sHeader & "</a>"
				If (Len(Trim(p_strSortImgUp)) > 0) And (Len(Trim(p_strSortImgDn)) > 0) And (p_strSortName = sColumn) Then					If p_strSortType = "asc" Then
						p_sSort = p_sSort & "&nbsp;<img src=""" & p_strSortImgUp & """ alt=""Sort Ascending"">&nbsp;"					Else
						p_sSort = p_sSort & "&nbsp;<img src=""" & p_strSortImgDn & """ alt=""Sort Descending"">&nbsp;"					End If
				End If				p_blnSortCol = True
			End If
		Else
			p_sSort = sHeader			p_blnSortCol = False
		End If
		If Len(Trim(sTemplate)) > 0 Then
			p_sColumn = sTemplate
		Else
			p_sColumn = "&nbsp;"
		End If
		p_strDisplayCols = p_strDisplayCols & sColumn & "|=|" & sHeader & "|=|" & p_sSort & "|=|" & p_sColumn & "|||"
		p_blnAutoCols = False
	End Sub


	Public Function ReturnDataGrid()
		Dim p_objField
		If IsObject(p_objConn) And Not IsEmpty(p_strCommand) Then
			Set p_objRS = Server.CreateObject("ADODB.Recordset")
			Set p_objRS = p_objConn.Execute(p_strCommand)
			p_blnPrivConn = True
		End If
		If IsObject(p_objRS) And Not p_objRS.EOF Then
			If p_blnAllowPaging Then
				SetFilter p_objRS, p_intPageSize, p_intAbsPage, p_strSortName, p_strSortType
			End If
			p_strOutput = "<table width=""" & p_strGridWidth & """ border=""" & p_intGridBorder & """ cellpadding=""" & p_intPadding & """ cellspacing=""" & p_intSpacing & """" & p_strGridAlign & p_strGridClass & p_strGridStyle & ">" & vbCrLf & "<tr>"
			If Not p_blnAutoCols Then
				p_arrDisplayCols = Split(p_strDisplayCols, "|||")
				For p_x = 0 To UBound(p_arrDisplayCols) - 1
					p_arrTemp = Split(p_arrDisplayCols(p_x), "|=|")
					For p_i = 0 To p_objRS.Fields.Count - 1
						If p_arrTemp(0) = p_objRS.Fields(p_i).Name Then
							p_strHeader = p_strHeader & "<td nowrap align=""center""" & p_strHeadClass & p_strHeadStyle & ">" & FormatTemplate(p_objRS.Fields, p_arrTemp(2)) & "</td>"
							p_intColCount = p_intColCount + 1
						End If
					Next
				Next
			Else
				For Each p_objField In p_objRS.Fields
					p_strHeader = p_strHeader & "<td nowrap align=""center""" & p_strHeadClass & p_strHeadStyle & ">" & FormatTemplate(p_objRS.Fields, p_objField.Name) & "</td>"
					p_intColCount = p_intColCount + 1
				Next
			End If
			p_strOutput = p_strOutput & p_strHeader & "</tr>" & vbCrLf
			If Not p_blnAutoCols Then
				Do Until p_objRS.EOF
					p_strThisClass = ReturnNextClass(p_intCurRow)
					p_strThisStyle = ReturnNextStyle(p_intCurRow)
					p_strOutput = p_strOutput & "<tr>"
					For p_x = 0 To UBound(p_arrDisplayCols) - 1
						p_arrTemp = Split(p_arrDisplayCols(p_x), "|=|")
						For p_i = 0 To p_objRS.Fields.Count - 1
							If p_arrTemp(0) = p_objRS.Fields(p_i).Name Then
							' Check if Column is to keep total running
								p_strOutput = p_strOutput & "<td" & p_strThisClass & p_strThisStyle & ">" & FormatTemplate(p_objRS.Fields, p_arrTemp(3)) & "</td>"
							End If
						Next
					Next
					p_strOutput = p_strOutput & "</tr>" & vbCrLf
					p_intCurRow = p_intCurRow + 1
					If p_blnAllowPaging Then
						If CLng(p_intPageSize) <> -1 And p_intCurRow > CLng(p_intPageSize) Then Exit Do	
					End If
					p_objRS.MoveNext
				Loop
			Else
				Do Until p_objRS.EOF
					p_strThisClass = ReturnNextClass(p_intCurRow)
					p_strThisStyle = ReturnNextStyle(p_intCurRow)
					p_strOutput = p_strOutput & "<tr>"
					For Each p_objField In p_objRS.Fields
						p_strOutput = p_strOutput & "<td" & p_strThisClass & p_strThisStyle & ">" & FormatTemplate(p_objRS.Fields, "#" & p_objField.Name & "#") & "</td>"
					Next
					p_strOutput = p_strOutput & "</tr>" & vbCrLf
					p_intCurRow = p_intCurRow + 1
					If p_blnAllowPaging Then
						If CLng(p_intPageSize) <> -1 And p_intCurRow > CLng(p_intPageSize) Then Exit Do	
					End If
					p_objRS.MoveNext
				Loop
			End If
			p_strOutput = p_strOutput & CreateNav(p_intPageSize, p_objRS.PageCount, p_intAbsPage, p_objRS.RecordCount, p_strQS) & "</table>" & vbCrLf
			ReturnDataGrid = p_strOutput
			If p_blnPrivConn Then
				p_objRS.Close
				p_objConn.Close
				Set p_objRS = Nothing
				Set p_objConn = Nothing
			End If
		Else
			ReturnDataGrid = "<div align=""center""><B>No Records Found.</B></div>"
		End If
	End Function


    Public Function CreateTotalsColumns(ByRef oRs)
    Dim strTotalRow
    
      If p_blnSumColumns Then
        oRs.MoveFirst
        
        Dim uColumns
        uColumns = Split(p_strSumColumnsName,",")
        ReDim uTotal(UBOUND(uColumns))            
        Dim x         
                
            Do Until oRs.EOF
                for x = 0 to uBound(uColumns)
                    uTotal(x) = uTotal(x) + oRs(uColumns(x))
                Next
            oRs.MoveNext
            Loop
            
            for x = 0 to uBound(uColumns)
                strTotalRow = strTotalRow & " Sum: " & uColumns(x) & " " & FormatNumber(uTotal(x),0) & "<br/>"
            Next
        
                
        End If        
        CreateTotalsColumns = strTotalRow
    End Function

	Public Sub Bind()		
		If IsObject(p_objConn) And Not IsEmpty(p_strCommand) Then
			Set p_objRS = Server.CreateObject("ADODB.Recordset")
			Set p_objRS = p_objConn.Execute(p_strCommand)
			p_blnPrivConn = True
		End If
		If IsObject(p_objRS) And Not p_objRS.EOF Then
			If p_blnAllowPaging Then
				SetFilter p_objRS, p_intPageSize, p_intAbsPage, p_strSortName, p_strSortType
			End If
			If p_blnFullRowSelect Then
				Response.Write CreateFullRowSelectScript()
			End If
			Response.Write "<table width=""" & p_strGridWidth & """ border=""" & p_intGridBorder & """ cellpadding=""" & p_intPadding & """ cellspacing=""" & p_intSpacing & """" & p_strGridAlign & p_strGridClass & p_strGridStyle & ">" & vbCrLf & "<tr>"
			If Not p_blnAutoCols Then
				p_arrDisplayCols = Split(p_strDisplayCols, "|||")
				For p_x = 0 To UBound(p_arrDisplayCols) - 1
					p_arrTemp = Split(p_arrDisplayCols(p_x), "|=|")
					For p_i = 0 To p_objRS.Fields.Count - 1
						If p_arrTemp(0) = p_objRS.Fields(p_i).Name Then
							Response.Write "<td nowrap align=""center""" & p_strHeadClass & p_strHeadStyle & ">" & FormatTemplate(p_objRS.Fields, p_arrTemp(2)) & "</td>"
							p_intColCount = p_intColCount + 1
						End If
					Next
				Next
			Else
				For Each p_objField In p_objRS.Fields
					Response.Write "<td nowrap align=""center""" & p_strHeadClass & p_strHeadStyle & ">" & FormatTemplate(p_objRS.Fields, p_objField.Name) & "</td>"
					p_intColCount = p_intColCount + 1
				Next
			End If
			Response.Write "</tr>" & vbCrLf
			If Not p_blnAutoCols Then
				Do Until p_objRS.EOF
					p_strThisClass = ReturnNextClass(p_intCurRow)
					p_strThisStyle = ReturnNextStyle(p_intCurRow)
					
					If p_blnFullRowSelect Then
						Response.Write "<tr " & p_strThisClass & p_strThisStyle & " onmouseover=""ChangeColor(this, true);"" onmouseout=""ChangeColor(this, false);"" onclick=""DoNav('" & FormatTemplate(p_objRS.Fields, p_strFullRowLink) & "');"">"
					Else
						Response.Write "<tr" & p_strThisClass & p_strThisStyle & ">"
					End If
					
					For p_x = 0 To UBound(p_arrDisplayCols) - 1
						p_arrTemp = Split(p_arrDisplayCols(p_x), "|=|")
						For p_i = 0 To p_objRS.Fields.Count - 1
							If p_arrTemp(0) = p_objRS.Fields(p_i).Name Then
								Response.Write "<td>" & FormatTemplate(p_objRS.Fields, p_arrTemp(3)) & "</td>"
							End If
						Next
					Next
					Response.Write "</tr>" & vbCrLf
					p_intCurRow = p_intCurRow + 1
					If p_blnAllowPaging Then
						If CLng(p_intPageSize) <> -1 And p_intCurRow > CLng(p_intPageSize) Then Exit Do	
					End If
					p_objRS.MoveNext
				Loop
			Else
				Do Until p_objRS.EOF
					p_strThisClass = ReturnNextClass(p_intCurRow)
					p_strThisStyle = ReturnNextStyle(p_intCurRow)
					Response.Write "<tr>"
					For Each p_objField In p_objRS.Fields
						Response.Write "<td" & p_strThisClass & p_strThisStyle & ">" & FormatTemplate(p_objRS.Fields, "#" & p_objField.Name & "#") & "</td>"
					Next
					Response.Write "</tr>" & vbCrLf
					p_intCurRow = p_intCurRow + 1
					If p_blnAllowPaging Then
						If CLng(p_intPageSize) <> -1 And p_intCurRow > CLng(p_intPageSize) Then Exit Do	
					End If
					p_objRS.MoveNext
				Loop
			End If
			Response.Write CreateNav(p_intPageSize, p_objRS.PageCount, p_intAbsPage, p_objRS.RecordCount, p_strQS) & "</table>" & vbCrLf
			if(p_blnSumColumns) Then			
			    Response.Write(CreateTotalsColumns(p_objRs))			
			End If
			If p_blnPrivConn Then
				p_objRS.Close
				p_objConn.Close
				Set p_objRS = Nothing
				Set p_objConn = Nothing
			End If
		Else
			Response.Write "<div align=""center""><B>No Records Found.</B></div>"
		End If
	End Sub


End Class
%>