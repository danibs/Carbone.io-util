Option Explicit

' Highlight Carbone.io placeholders
' 
' @author danibs
' @version 1.0
' @see https://github.com/danibs/Carbone.io-util/
'
' @see https://ask.libreoffice.org/t/writer-text-position/93296/
' @See https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Search_and_Replace
' @see https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRangeCompare.html
' @see https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Working_with_Documents
' @see https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Searching_for_Text_Portions
' @see https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html

' *********************************************************
' Style "CarbonPlaceholder"'s values
Private Const FONT_NAME = "Courier New"
Private Const FONT_COLOR = &H000000
Private Const FONT_BACKCOLOR = &HEEEEEE
Private Const FONT_HEIGHT = 8

' new character style, do not change!
Private Const CHARACTER_STYLE = "CarbonPlaceholder"
' *********************************************************

' *********************************************************
' Open / Close placeholder
Private Const OPEN_PLACEHOLDER = "{"
Private Const CLOSE_PLACEHOLDER = "}"
' *********************************************************

' *********************************************************
' Valid search sections
Private Const SEARCH_ON_BODY = True
Private Const SEARCH_ON_HEADER_FOOTER = False
Private Const SEARCH_ON_CELL = True
' *********************************************************


' *********************************************************
' Last error count: ERR05
' *********************************************************

' Main function, the ONLY ONE to execute
Sub highlightCarbonePlaceholder()
	Dim oDoc As Object ' com.sun.star.text.GenericTextDocument, com.sun.star.util.XSearchable
	Dim oFoundZone As Object ' com.sun.star.uno.XInterface
	Dim oCursVisible As Object ' com.sun.star.text.text.XTextViewCursor
	Dim nullObject As Object
	
	oDoc = ThisComponent ' com.sun.star.util.XSearchable compliant document
 	createCharacterStyle(oDoc)
 	oCursVisible = createVisibleCursor(oDoc)

  	oFoundZone = highlightTheCouple(oDoc, oCursVisible, nullObject)
  	Do While Not(IsNull(oFoundZone))
 		oFoundZone = highlightTheCouple(oDoc, oCursVisible, oFoundZone)
 	Loop
End Sub


Private Function highlightTheCouple(ByRef oDoc As Object,  ByRef oCursVisible As Object, ByRef oZoneStartFrom As Object) As Object
	Dim oSearchZoneOpen As Object  ' com.sun.star.util.SearchDescriptor
	Dim oSearchZoneClose As Object  ' com.sun.star.util.SearchDescriptor
	Dim oFoundZoneOpen As Object
	DIm oFoundZoneNextOpen As Object
	Dim oFoundZoneClose As Object
	
	
	oSearchZoneOpen = createSearchDescriptor(oDoc, OPEN_PLACEHOLDER)
	oSearchZoneClose = createSearchDescriptor(oDoc, CLOSE_PLACEHOLDER)

	' find 1st {
	If IsNull(oZoneStartFrom) Then
		oFoundZoneOpen = oDoc.findFirst(oSearchZoneOpen)
	Else
		oFoundZoneOpen = oDoc.findNext(oZoneStartFrom.End, oSearchZoneOpen)
	End If

	If Not(IsNull(oFoundZoneOpen)) Then
		' find next {
		oFoundZoneNextOpen = oDoc.findNext(oFoundZoneOpen.End, oSearchZoneOpen)
		
		' find } after 1st {
		oFoundZoneClose = oDoc.findNext(oFoundZoneOpen.End, oSearchZoneClose)
	End If

	' found {, not found }
	If Not(IsNull(oFoundZoneOpen)) And IsNull(oFoundZoneClose) Then
		With oCursVisible
			.goToRange(oFoundZoneOpen.Start, False)
			.goToRange(oFoundZoneOpen.End, True)
		End With
		
		MsgBox "ERR01 - Cannot find closing parenthesis", MB_ICONEXCLAMATION
		Exit Function
	End If
	
	' found next {
	If Not(IsNull(oFoundZoneNextOpen)) Then
		Dim openingSection As String
		Dim closingSection As String
		
		openingSection = oFoundZoneOpen.Text.ImplementationName
		closingSection = oFoundZoneClose.Text.ImplementationName
		
		If openingSection <> closingSection Then
			With oCursVisible
				.gotoRange(oFoundZoneOpen.Start, False)
				.gotoRange(oFoundZoneOpen.End, True)
			End With

			MsgBox "ERR05 - Cannot find closing parenthesis", MB_ICONEXCLAMATION
			Exit Function
		End If
	
	
		With oCursVisible
			.gotoRange(oFoundZoneOpen.End, False)
			.gotoRange(oFoundZoneClose.Start, True)
	
			If InStr( .String, OPEN_PLACEHOLDER) > 0 Then
				MsgBox "ERR02 - Found the opening parenthesis before the closing parenthesis"
				Exit Function
			End If
		End With
	End If
	
	
	If Not(IsNull(oFoundZoneOpen)) And Not(IsNull(oFoundZoneClose)) Then
		With oCursVisible
			.gotoRange(oFoundZoneOpen.End, False)
			.gotoRange(oFoundZoneClose.Start, True)
	
			If InStr( .String, OPEN_PLACEHOLDER) > 0 Then
				MsgBox "ERR03 - Found the opening parenthesis before the closing parenthesis"
				Exit Function
			End If
		End With
	
		If isValidSection(oFoundZoneOpen) Then
			' https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XController.html
			oDoc.CurrentController.Select(oCursVisible)
			oCursVisible.CharStyleName = CHARACTER_STYLE
		End If
	
		highlightTheCouple = oFoundZoneClose
	End If
	
	Exit Function


ErrorHandler:
	' highlight error
	With oCursVisible
		.goToRange(oFoundZoneNextOpen.Start, False)
		.goToRange(oFoundZoneNextOpen.End, True)
	End With
	
    MsgBox "ERR04 - Something went wrong", MB_ICONEXCLAMATION
End Function


Private Function createSearchDescriptor(ByRef oDoc As Object, str As String) As Object
	Dim oSearchZone As Object  ' com.sun.star.util.SearchDescriptor

	oSearchZone = oDoc.createSearchDescriptor()
	oSearchZone.SearchString = str
	
	createSearchDescriptor = oSearchZone
End Function


Private Function createVisibleCursor(ByRef oDoc As Object) As Object
	Dim oCursVisible As Object ' com.sun.star.text.text.XTextViewCursor

	oCursVisible = oDoc.CurrentController.ViewCursor  ' visible cursor
	oCursVisible.goToStart(false)
	
	createVisibleCursor = oCursVisible
End Function

Private Function isValidSection(ByRef oFoundZone As Object) As Boolean
	' SwXBodyText for "body"
	' SwXHeadFootText for Header/Footer
	' SwXCell for table's cell
	Dim sectionName As String

	sectionName = oFoundZone.Text.ImplementationName
	
	isValidSection = (sectionName = "SwXBodyText" And SEARCH_ON_BODY = True) _
							Or (sectionName = "SwXHeadFootText" And SEARCH_ON_HEADER_FOOTER = True) _
							Or (sectionName = "SwXCell" And SEARCH_ON_CELL = True)
End Function

' create new character style if not exists
' once created, it remains in memory until LibreOffice is restarted
Private Sub createCharacterStyle(ByRef oDoc As Object)
	Dim oStyles As Object
	Dim oStyle As Object
	
	oStyles = oDoc.StyleFamilies.getByName("CharacterStyles")
	
	If Not(oStyles.hasByName(CHARACTER_STYLE)) Then
		oStyle = oDoc.createInstance("com.sun.star.style.CharacterStyle")
	
		With oStyle
			.CharFontName = FONT_NAME
			.CharColor = FONT_COLOR
			.CharBackColor = FONT_BACKCOLOR
			.CharHeight = FONT_HEIGHT
			
			.CharWeight = com.sun.star.awt.FontWeight.NORMAL ' remove Bold
			.CharPosture = com.sun.star.awt.FontSlant.NONE ' remove Italic
			.CharUnderline = com.sun.star.awt.FontUnderline.NONE ' remove Underline
			.CharStrikeout = com.sun.star.awt.FontStrikeout.NONE ' remove Strikeout
		End With
		
		oStyles.insertByName(CHARACTER_STYLE, oStyle)
	End If
End Sub
