VERSION 5.00
Begin VB.UserControl PostalAddress 
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4110
   ToolboxBitmap   =   "PostalAddress.ctx":0000
   Begin VB.TextBox txtContents 
      Height          =   1635
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "PostalAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################
' Postal Address Data Entry Control
' - Created 16 May 2000 by James Vincent Carnicelli
' - Updated 18 May 2000 by James Vincent Carnicelli
'
' Notes:
' Standardized postal addresses are used almost everywhere in
' typical business applications.  This control is designed to
' allow users to enter postal addresses validly while hiding
' the validation details.  This control assumes a format typical
' for most U.S. postal address that fits the following pattern:
'   <street address line1>
'   [<street address line 2>]
'   <city>, <state> <zip>[-<4-digit zip extension>]
'
' Here's a typical usage:
'
' Private Sub Form_Load()
'     PostalAddress1.Address = "123 XYZ Plaza" & vbCrLf & "Hereton, MN 12345"
' End Sub
'
' Private Sub PostalAddress1_Change()
'     Text1.Text = PostalAddress1.IsValid & ": " & PostalAddress1.ValidationError & vbCrLf & vbCrLf _
'       & PostalAddress1.Address & vbCrLf & vbCrLf _
'       & PostalAddress1.StateName(PostalAddress1.State)
' End Sub
'
' Private Sub PostalAddress1_LostFocus()
'     If Not PostalAddress1.IsValid Then
'         MsgBox "Not valid: " & PostalAddress1.ValidationError
'     Else
'         Text1.Text = PostalAddress1.Address & vbCrLf & vbCrLf _
'           & PostalAddress1.StateName(PostalAddress1.State)
'     End If
' End Sub
'
' While this control comes initialized with names and
' abbreviations for the 50 United States and pseudo-states the
' U.S. Post Office recognizes (e.g., VI = Virgin Islands), your
' code may prefer to clear this out and repopulate it with your
' own list -- perhaps a region of the country.  To do so, just
' call .ClearStates and .AddState(<abbrev>, <name>) with each
' state you want to add.
'################################################################

Option Explicit


'######## Private Declarations ########################

'Internal storage for public properties
Private msStreetLine1 As String
Private msStreetLine2 As String
Private msCity As String
Private msState As String
Private msZip As String
Private msZip4 As String
Private msValidationError
Private mnValidationErrorCode

'Variant array of 2-character abbreviations for states
Private maStateAbbreviations As Variant

'Variant array of full names of states associated with abbreviations
Private maStateNames As Variant


'######## Public Declarations ########################

'This is an array containing the complete list of error codes and their 
'messages, which are used to populate .ValidationError.  
'.ValidationErrorCode. will contain an index to this array.  Those list 
'items that include ???? in them will have error-specific information 
'in place of ???? when they are used to populate .ValidationError.
'User code may alter the contents of this array to provide custom error 
'messages.
Public ValidationErrors As Variant

Public Event Change()
Public Event Click() 'MappingInfo=txtContents,txtContents,-1,Click
Public Event DblClick() 'MappingInfo=txtContents,txtContents,-1,DblClick
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtContents,txtContents,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtContents,txtContents,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtContents,txtContents,-1,KeyUp
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtContents,txtContents,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtContents,txtContents,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtContents,txtContents,-1,MouseUp


'######## Public Properties ########################

'This indicates whether or not the control considers the address to
'have a valid format and a recognized state code.  Since one's own
'program may have other validation rules (such as maximum length for
'the city name), these might best be implemented by checking the
'relevant data properties (e.g., Len(PostalAddress1.City) <= 20)
Public Property Get IsValid() As Boolean
    IsValid = (mnValidationErrorCode = -1)
End Property

'This address may be optional in your application.  This property
'identifies whether or not this control has any data in it at all.
'It does this by simply seeing if all the publicly-available data
'properties are blank.
Public Property Get IsEmpty() As Boolean
    IsEmpty = (msStreetLine1 = "" _
           And msStreetLine2 = "" _
           And msCity = "" _
           And msState = "" _
           And msZip = "" _
           And msZip4 = "")
End Property

'The parsed and reassembled address.  While it may look like the
'same thing that the user typed into the text box, it is not
'necessarily so.  It is instead constructed using the parsed-out
'basic parts of the address.  Setting this property causes the
'text box to be populated verbatim and that to then be parsed.
'New-line character combinations (<CR><LF>) separate the two or
'three lines.
Public Property Get Address() As String
Attribute Address.VB_ProcData.VB_Invoke_Property = "StandardDataFormat;Data"
Attribute Address.VB_UserMemId = 0
Attribute Address.VB_MemberFlags = "400"
    Address = Me.Street & vbCrLf & msCity & ", " & msState & " " & msZip
    If msZip4 <> "" Then
        Address = Address & "-" & msZip4
    End If
End Property
Public Property Let Address(newAddress As String)
    txtContents.Text = newAddress
    'Parsing happens as a result
End Property

'The first of at least one and at most two lines representing
'the building number, street name, and any addition stuff like
'suite number.
Public Property Get StreetLine1() As String
Attribute StreetLine1.VB_MemberFlags = "400"
    StreetLine1 = msStreetLine1
End Property
Public Property Let StreetLine1(ByVal newStreetLine1 As String)
    msStreetLine1 = Trim(newStreetLine1)
    msRefresh
End Property

'The optional line two of the street address.  If there is no
'line two, this will be blank.  See .StreetLine1 for more.
Public Property Get StreetLine2() As String
Attribute StreetLine2.VB_MemberFlags = "400"
    StreetLine2 = msStreetLine2
End Property
Public Property Let StreetLine2(ByVal newStreetLine2 As String)
    msStreetLine2 = Trim(newStreetLine2)
    msRefresh
End Property

'The simple combination of street address lines one and two,
'with a new-line character combination between them if there
'is a second line.
Public Property Get Street() As String
Attribute Street.VB_MemberFlags = "400"
    If msStreetLine2 = "" Then
        Street = msStreetLine1
    Else
        Street = msStreetLine1 & vbCrLf & msStreetLine2
    End If
End Property
Public Property Let Street(ByVal newStreet As String)
    Dim lPos As Long
    lPos = InStr(1, newStreet, vbCrLf)
    If lPos = 0 Then
        msStreetLine1 = Trim(newStreet)
        msStreetLine2 = ""
    Else
        msStreetLine1 = Trim(Left(newStreet, lPos - 1))
        msStreetLine2 = Trim(Mid(newStreet, lPos + 2))
    End If
    msRefresh
End Property

'The city, which is basically everything in the text box before
'the state (usually, but not necessarily, followed by a comma
'before the state).
Public Property Get City() As String
Attribute City.VB_MemberFlags = "400"
    City = msCity
End Property
Public Property Let City(ByVal newCity As String)
    msCity = Trim(newCity)
    msRefresh
End Property

'The two-character state code (e.g., "AK").
Public Property Get State() As String
Attribute State.VB_MemberFlags = "400"
    State = msState
End Property
Public Property Let State(ByVal newState As String)
    msState = Trim(newState)
    msRefresh
End Property

'The first (or only) five digits of the ZIP code.
Public Property Get Zip() As String
Attribute Zip.VB_MemberFlags = "400"
    Zip = msZip
End Property
Public Property Let Zip(ByVal newZip As String)
    msZip = Trim(newZip)
    msRefresh
End Property

'If the zip code is an extended one, this will be the four
'digits following the first and the dash (e.g., "62542-7311").
Public Property Get Zip4() As String
Attribute Zip4.VB_MemberFlags = "400"
    Zip4 = msZip4
End Property
Public Property Let Zip4(ByVal newZip4 As String)
    msZip4 = Trim(newZip4)
    msRefresh
End Property

'If .IsValid = False, this contains a terse explanation of
'what's most obviously wrong with the address.  See the
'ValidationErrors array declaration for the list of errors.
Public Property Get ValidationError() As String
    ValidationError = msValidationError
End Property
Public Property Get ValidationErrorCode() As Integer
    ValidationErrorCode = mnValidationErrorCode
End Property

'While .Address provides the "real" address, this allows the
'containing program to see what's actually in the text box.
Public Property Get RawText() As String
Attribute RawText.VB_MemberFlags = "400"
    RawText = txtContents.Text
End Property
Public Property Let RawText(New_RawText As String)
    txtContents.Text = New_RawText
    'Parsing happens as a result
End Property

'The number of states in the internal list.
Public Property Get StateCount() As Integer
Attribute StateCount.VB_MemberFlags = "400"
    StateCount = UBound(maStateAbbreviations) + 1
End Property



'The following properties are basically mirrors of what's built
'into the UserControl or text box; e.g., the background color.
'There's no significant processing associated with them.

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtContents.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    txtContents.SelStart = New_SelStart
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtContents.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    txtContents.SelLength = New_SelLength
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtContents.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
    txtContents.SelText = New_SelText
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtContents,txtContents,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtContents.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtContents.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtContents,txtContents,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtContents.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtContents.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtContents,txtContents,-1,Font
Public Property Get Font() As Font
    Set Font = txtContents.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set txtContents.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtContents,txtContents,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = txtContents.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txtContents.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


'######## Public Methods ########################

'Clear the text box and thus all basic data properties (e.g.,
'City).  Since this forces validation like any other changing
'of the text box's contents, .IsValid will be False and the
'validation error will have an appropriate message.
Public Sub Clear()
    txtContents.Text = ""
    'Parsing happens as a result
End Sub

'Populate all the data properties in one step.  This is simple, but
'it also helps to avoid strange problems that can happen during
'the refresh that follows setting one of the properties when the
'address is invalid.
Public Sub Populate(ByVal StreetLine1 As String, ByVal StreetLine2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Zip4 As String)
    msStreetLine1 = Trim(StreetLine1)
    msStreetLine2 = Trim(StreetLine2)
    msCity = Trim(City)
    msState = Trim(State)
    msZip = Trim(Zip)
    msZip4 = Trim(Zip4)
    msRefresh
End Sub

'Clear out the lists of state abbreviations and names to make way
'for new ones using .AddState().
Public Sub ClearStates()
    maStateAbbreviations = Array()
    maStateNames = Array()
End Sub

'Add a state to the end of the internal list.
Public Sub AddState(ByVal Abbreviation As String, Name As String)
    ReDim Preserve maStateAbbreviations(UBound(maStateAbbreviations) + 1)
    ReDim Preserve maStateNames(UBound(maStateNames) + 1)
    maStateAbbreviations(UBound(maStateAbbreviations)) = Abbreviation
    maStateNames(UBound(maStateNames)) = Name
End Sub

'Identify the state abbreviation associated with the specified
'list index.  Since the list is zero-based, the index must be
'in the range 0 to .StateCount - 1.
Public Function StateAbbreviation(ByVal Index As Integer) As String
    Dim nPos As Integer
    StateAbbreviation = maStateAbbreviations(Index)
End Function

'Identify the name of the state associated with the abbreviation.
'A typical use of this is:
'   X = PostalAddress.StateName(PostalAddress.State)
Public Function StateName(ByVal Abbreviation As String) As String
    Dim nPos As Integer
    Abbreviation = UCase(Abbreviation)
    nPos = FindStateByAbbreviation(Abbreviation)
    If nPos = -1 Then Exit Function
    StateName = maStateNames(nPos)
End Function

'Returns the zero-based list index of the specified state, if
'found.  Returns -1 if not found.
Public Function FindStateByAbbreviation(ByVal Abbreviation As String) As Integer
    Dim i As Integer
    Abbreviation = UCase(Abbreviation)
    For i = 0 To UBound(maStateAbbreviations)
        If maStateAbbreviations(i) = Abbreviation Then
            FindStateByAbbreviation = i
            Exit Function
        End If
    Next
    FindStateByAbbreviation = -1
End Function

'Returns the zero-based list index of the specified state, if
'found.  Returns -1 if not found.
Public Function FindStateByName(ByVal Name As String) As Integer
    Dim i As Integer
    Name = UCase(Name)
    For i = 0 To UBound(maStateNames)
        If UCase(maStateNames(i)) = Name Then
            FindStateByName = i
            Exit Function
        End If
    Next
    FindStateByName = -1
End Function


'######## Hidden Members ########################


'######## Private Methods ########################

'Identifies "where" the cursor is in the text box's representation of
'the postal address, including the number of recognized lines so
'far constructed, the line (starting with 1) the cursor is on,
'the number of characters in from the start of the line (starting
'with 1) it is, and the contents of the line the cursor is
'currently on.  This is currently not used within this control,
'but may have use in your own application.
Private Sub msWhereAmI(ByRef prnLineCount As Integer, ByRef prnLine As Integer, ByRef prnPosOnLine As Integer, ByRef prsLineContents As String)
    Dim nPos As Integer, nNewPos As Integer, sContents As String
    
    sContents = txtContents.Text
    While Right(sContents, 2) = vbCrLf
        sContents = Left(sContents, Len(sContents) - 2)
    Wend
    If Len(sContents) < txtContents.SelStart Then sContents = sContents & vbCrLf
    
    nPos = 1
    prnLine = 1
    Do
        nNewPos = InStr(nPos, sContents, vbCrLf)
        If nNewPos = 0 Then
            prsLineContents = Mid(sContents, nPos)
            prnPosOnLine = txtContents.SelStart - nPos + 1
            prnLineCount = prnLine
            Exit Do
        End If
        prsLineContents = mfSafeMid(sContents, nPos, nNewPos - nPos)
        prnPosOnLine = txtContents.SelStart - nPos + 2
        If nNewPos > txtContents.SelStart Then
            nPos = nNewPos + 2
            prnLineCount = prnLine + 1
            Do
                nNewPos = InStr(nPos, sContents, vbCrLf)
                If nNewPos = 0 Then Exit Do
                prnLineCount = prnLineCount + 1
                nPos = nNewPos + 2
            Loop
            Exit Do
        End If
        prnLine = prnLine + 1
        nPos = nNewPos + 2
    Loop
End Sub

'Replaces the text box's contents with the "real" address
'assembled from the data properties.
Private Sub msRefresh()
    txtContents.Text = Me.Address
End Sub

'Here's the heart of this control: the parser.  This function
'reads the text box to pick out the data properties, generates
'an appropriate validation error message as needed, and returns
'True only if it considers the address to be completely valid.
Private Sub msParse()
    Dim aLines As Variant, i As Long, sLine As String
    Dim sToken As String
    
    'Set up for parse
    mnValidationErrorCode = -1
    msValidationError = ""
    
    'Extract all lines of text and remove blank ones
    aLines = Split(txtContents.Text, vbCrLf)
    i = 0
    Do
        If i > UBound(aLines) Then Exit Do
        sLine = Trim(aLines(i))
        If sLine = "" Then
            'Eliminate this blank line
            msRemoveFromArray aLines, i
        Else
            aLines(i) = Trim(sLine)
            i = i + 1
        End If
    Loop
    If UBound(aLines) = -1 Then
        msSelectError 0  'Missing entire address
        Exit Sub
    End If
    
    'Now let's see if there's a City/State/ZIP line
    sLine = aLines(UBound(aLines))
    
    '---- Parse out ZIP code ----
    i = mfInStrLast(sLine, " ")
    If i = 0 Then
        sLine = " " & sLine
        i = 1
    End If
    If i <> 0 Then
        sToken = Replace(mfSafeMid(sLine, i + 1), "-", "") 'Whack out dashes
        If mfAreDigits(sToken) Then
            If Len(sToken) = 5 Then
                msZip = sToken
                msZip4 = ""
                sLine = Trim(Left(sLine, i - 1))
            ElseIf Len(sToken) = 9 Then
                msZip = Left(sToken, 5)
                msZip4 = Right(sToken, 4)
                sLine = Trim(Left(sLine, i - 1))
            Else
                msZip = ""
                msZip4 = ""
                sLine = Trim(Left(sLine, i - 1))
                msSelectError 3, Len(sToken)  'Found ZIP code with ???? digits instead of 5 or 9
            End If
        Else
            msZip = ""
            msZip4 = ""
            msSelectError 2  'Couldn't find a ZIP code
        End If
    End If
    While mfSafeRight(sLine, 1) = ","
        sLine = Trim(Left(sLine, Len(sLine) - 1))
    Wend
    
    '---- Parse out state ----
    If Len(sLine) = 2 Then
        sLine = " " & sLine
        sToken = sLine
    Else
        sToken = mfSafeMid(sLine, Len(sLine) - 2)
    End If
    If Len(sToken) = 3 _
      And (Left(sToken, 1) = " " Or Left(sToken, 1) = ",") _
      And mfAreLetters(mfSafeMid(sToken, 2)) Then
        msState = UCase(mfSafeRight(sLine, 2))
        sLine = Trim(Left(sLine, Len(sLine) - 3))
    Else  'Ok, no 2-letter state code
        'Search for a full state name
        For i = 0 To UBound(maStateNames) + 1
            If i > UBound(maStateNames) Then  'No more to try
                msState = ""
                msSelectError 4  'Couldn't find a state
                Exit For
            End If
            If mfMatchAtEnd(sLine, maStateNames(i), msState) Then
                msState = maStateAbbreviations(i)
                Exit For
            End If
        Next
    End If
    sLine = Trim(sLine)
    While mfSafeRight(sLine, 1) = ","
        sLine = Trim(Left(sLine, Len(sLine) - 1))
    Wend
    
    'Let's see if this is a "real" state
    i = FindStateByAbbreviation(msState)
    If i = -1 Then
        msSelectError 8, msState  'Can't recognize "????" as a state
    Else
        msState = maStateAbbreviations(i)
    End If
    
    '---- Parse out city ----
    If msZip <> "" Then
        While mfSafeRight(sLine, 1) = ","
            sLine = Trim(Left(sLine, Len(sLine) - 1))
        Wend
        msCity = Trim(sLine)
        
        If msCity = "" Then
            msSelectError 5  'Couldn't find a city
        End If
    Else
        msCity = ""
        msSelectError 5  'Couldn't find a city
    End If
    
    'Remove city/state/ZIP line from consideration
    If msCity <> "" Or msState <> "" Or msZip <> "" Then
        msRemoveFromArray aLines, UBound(aLines)
    End If
    
    '---- Parse out street address ----
    msStreetLine1 = "":  msStreetLine2 = ""
    If UBound(aLines) >= 0 Then msStreetLine1 = aLines(0)
    If UBound(aLines) >= 1 Then msStreetLine2 = aLines(1)
    
    'Too many lines
    If UBound(aLines) > 1 Then
        msSelectError 7  'Can't have more than two street-address lines
        Exit Sub
    End If
End Sub

'Select the error code and message to report to the user code.
'Note that only the first call of this during parsing will be
'respected.
Private Sub msSelectError(ByVal pvnCode As Integer, Optional ByVal pvsExtraParameter As String = "")
    Dim sError As String
    
    'Already called once?
    If mnValidationErrorCode <> -1 Then Exit Sub
    
    sError = ValidationErrors(pvnCode)
    If pvsExtraParameter <> "" Then
        sError = Replace(sError, "????", pvsExtraParameter)
    End If
    mnValidationErrorCode = pvnCode
    msValidationError = sError
End Sub

'Determine if pvsToFind can be found at the end of prsLine,
'regardless of case or intermittent spaces.  Returns actual
'match in prsFound and trims this off of prsLine.
Private Function mfMatchAtEnd(ByRef prsLine As String, ByVal pvsToFind As String, ByRef prsFound) As Boolean
    Dim nPosLine As Long, nPosToFind As Long
    Dim sCharLine As String, sCharToFind As String
    nPosLine = Len(prsLine)
    nPosToFind = Len(pvsToFind)
    Do
        If nPosLine < 1 Then
            If nPosToFind < 1 Then
                Exit Do  'Done matching
            Else
                Exit Function  'Not a match
            End If
        End If
        If nPosToFind < 1 Then Exit Do  'Done matching
        sCharLine = UCase(Mid(prsLine, nPosLine, 1))
        sCharToFind = UCase(Mid(pvsToFind, nPosToFind, 1))
        If sCharLine = " " And sCharToFind <> " " Then
            'Ignore extra spaces
            nPosLine = nPosLine - 1
        ElseIf sCharToFind = " " And sCharLine = " " Then
            'Whoa; expecting at least one space here
            nPosToFind = nPosToFind - 1
        ElseIf sCharLine = sCharToFind Then
            'This character matches
            nPosLine = nPosLine - 1
            nPosToFind = nPosToFind - 1
        Else
            'Not a match
            Exit Function
        End If
    Loop
    If nPosLine > 0 Then
        If Not mfAreLetters(Mid(prsLine, nPosLine, 1)) Then
            'Ok, at beginning of word
        Else
            Exit Function  'Must be beginning of word
        End If
    End If
    prsFound = Mid(prsLine, nPosLine + 1)
    prsLine = Left(prsLine, nPosLine)
    mfMatchAtEnd = True
End Function

'Are all the characters in the input string digits?
Private Function mfAreDigits(ByVal pvsText As String) As Boolean
    Dim lPos As Long, nChar As Integer
    If pvsText = "" Then Exit Function
    For lPos = 1 To Len(pvsText)
        nChar = Asc(Mid(pvsText, lPos, 1))
        If nChar < vbKey0 Then Exit Function
        If nChar > vbKey9 Then Exit Function
    Next
    mfAreDigits = True
End Function

'Are all the characters in the input string letters?
Private Function mfAreLetters(ByVal pvsText As String) As Boolean
    Dim lPos As Long, nChar As Integer
    If pvsText = "" Then Exit Function
    pvsText = UCase(pvsText)
    For lPos = 1 To Len(pvsText)
        nChar = Asc(Mid(pvsText, lPos, 1))
        If nChar < vbKeyA Then Exit Function
        If nChar > vbKeyZ Then Exit Function
    Next
    mfAreLetters = True
End Function

'The built-in Left(), Mid(), and Right() functions bomb under
'certain conditions -- especially when the start positions or
'lengths have unexpected values.  These three functions don't.
'Any bad input is acceptable.
Private Function mfSafeLeft(ByRef prsText As String, ByVal pvlLength As Long) As String
    On Error Resume Next
    If pvlLength < 0 Then Exit Function
    mfSafeLeft = Left(prsText, pvlLength)
End Function
Private Function mfSafeMid(ByRef prsText As String, ByVal pvlStart, Optional ByVal pvlLength) As String
    On Error Resume Next
    If pvlStart < 1 Then Exit Function
    If IsMissing(pvlLength) Then
        mfSafeMid = Mid(prsText, pvlStart)
    Else
        If pvlLength < 0 Then Exit Function
        mfSafeMid = Mid(prsText, pvlStart, pvlLength)
    End If
End Function
Private Function mfSafeRight(ByRef prsText As String, ByVal pvlLength As Long) As String
    On Error Resume Next
    If pvlLength < 1 Then Exit Function
    mfSafeRight = Right(prsText, pvlLength)
End Function

'This finds the last occurrance of the pattern in the text using
'InStr() until there are no more matches.  Returns 0 if the pattern
'is not found at all.
Private Function mfInStrLast(ByRef prsText As String, ByVal pvsPattern As String) As Long
    Dim lPos As Long
    lPos = 1
    Do
       lPos = InStr(lPos, prsText, pvsPattern)
       If lPos = 0 Then Exit Function
       mfInStrLast = lPos
       lPos = lPos + Len(pvsPattern)
    Loop
End Function

Private Sub msAppendToArray(ByRef praData As Variant, ByRef prvItem As Variant)
    ReDim Preserve praData(UBound(praData) + 1)
    If IsObject(prvItem) Then
        Set praData(UBound(praData)) = prvItem
    Else
        praData(UBound(praData)) = prvItem
    End If
End Sub

Private Sub msRemoveFromArray(ByRef praData As Variant, ByVal pviIndex As Integer)
    Dim i As Integer
    If UBound(praData) = 0 Then
        praData = Array()
    Else
        For i = pviIndex To UBound(praData) - 1
            If IsObject(praData(i + 1)) Then
                Set praData(i) = praData(i + 1)
            Else
                praData(i) = praData(i + 1)
            End If
        Next
        ReDim Preserve praData(UBound(praData) - 1)
    End If
End Sub


'######## Private Event Handlers ########################

'Clear out the data and provide the default population of states.
Private Sub UserControl_Initialize()
    Clear
    ClearStates
    AddState "AL", "Alabama"
    AddState "AK", "Alaska"
    AddState "AS", "American Samoa"
    AddState "AZ", "Arizona"
    AddState "AR", "Arkansas"
    AddState "CA", "California"
    AddState "CO", "Colorado"
    AddState "CT", "Conneticut"
    AddState "DE", "Delaware"
    AddState "DC", "District of Columbia"
    AddState "FM", "Federated States of Micronesia"
    AddState "FL", "Florida"
    AddState "GA", "Georgia"
    AddState "GU", "Guam"
    AddState "HI", "Hawaii"
    AddState "ID", "Idaho"
    AddState "IL", "Illinois"
    AddState "IN", "Indiana"
    AddState "IA", "Iowa"
    AddState "KS", "Kansas"
    AddState "KY", "Kentucky"
    AddState "LA", "Louisiana"
    AddState "ME", "Maine"
    AddState "MH", "Marshall Islands"
    AddState "MD", "Maryland"
    AddState "MA", "Massachusetts"
    AddState "MI", "Michigan"
    AddState "MN", "Minnesota"
    AddState "MS", "Mississippi"
    AddState "MO", "Missouri"
    AddState "MT", "Montana"
    AddState "NE", "Nebraska"
    AddState "NV", "Nevada"
    AddState "NH", "New Hampshire"
    AddState "NJ", "New Jersey"
    AddState "NM", "New Mexico"
    AddState "NY", "New York"
    AddState "NC", "North Carolina"
    AddState "ND", "North Dakota"
    AddState "MP", "Northern Mariana Islands"
    AddState "OH", "Ohio"
    AddState "OK", "Oklahoma"
    AddState "OR", "Oregon"
    AddState "PW", "Palau"
    AddState "PA", "Pensylvania"
    AddState "PR", "Puerto Rico"
    AddState "RI", "Rhode Island"
    AddState "SC", "South Carolina"
    AddState "SD", "South Dakota"
    AddState "TN", "Tennessee"
    AddState "TX", "Texas"
    AddState "UT", "Utah"
    AddState "VT", "Vermont"
    AddState "VI", "Virgin Islands"
    AddState "VA", "Virginia"
    AddState "WA", "Washington"
    AddState "WV", "West Virginia"
    AddState "WI", "Wisconsin"
    AddState "WY", "Wyoming"
    
    ReDim ValidationErrors(8)
    ValidationErrors(0) = "Missing entire address"
    ValidationErrors(1) = "Missing a valid city/state/ZIP code line"
    ValidationErrors(2) = "Couldn't find a ZIP code"
    ValidationErrors(3) = "Found ZIP code with ???? digits instead of 5 or 9"
    ValidationErrors(4) = "Couldn't find a state"
    ValidationErrors(5) = "Couldn't find a city"
    ValidationErrors(6) = "Missing street address"
    ValidationErrors(7) = "Can't have more than two street-address lines"
    ValidationErrors(8) = "Can't recognize ""????"" as a state"
End Sub

'Since the user is leaving, let's display the "real" address if
'we have it.
Private Sub UserControl_LostFocus()
    'An error we can overlook
    If msValidationError = "Can't have a blank first line for the street address" Then
        txtContents.Text = Me.Address
    End If
    
    If IsValid Then
        txtContents.Text = Me.Address
    End If
End Sub

'Make the text box the same size as this control.
Private Sub UserControl_Resize()
    On Error Resume Next
    
    If UserControl.Width = 0 Then UserControl.Width = 60 * Screen.TwipsPerPixelX
    If UserControl.Height = 0 Then UserControl.Height = 30 * Screen.TwipsPerPixelY
    
    txtContents.Width = UserControl.Width
    txtContents.Height = UserControl.Height
End Sub

'Load design-time property values from storage.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    txtContents.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtContents.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtContents.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtContents.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

'Write design-time property values to storage.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Call PropBag.WriteProperty("BackColor", txtContents.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtContents.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", txtContents.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", txtContents.BorderStyle, 1)
End Sub

'Make any final display initialization preparations.
Private Sub UserControl_Show()
    'Unfortunately, there does not appear to be any elegant way to
    'dynamically update txtContents.ToolTipText when this control's
    'Extender.ToolTipText property is changed by outside code.
    'This only applies the .ToolTipText property for the text box
    'that's specified at design time.
    
    On Error Resume Next  'Can't assume .ToolTipText will always be there
    txtContents.ToolTipText = UserControl.Extender.ToolTipText
    On Error GoTo 0
End Sub

'Since the text box has changed, we need to parse it again to
'keep up to date.
Private Sub txtContents_Change()
    msParse
    RaiseEvent Change
End Sub


'The following event handlers simply raise their same-named
'counterparts to the user code.  There's no significant processing
'associated with them.

Private Sub txtContents_Click()
    RaiseEvent Click
End Sub

Private Sub txtContents_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtContents_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtContents_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtContents_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtContents_LostFocus()
    UserControl_LostFocus
End Sub

Private Sub txtContents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtContents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtContents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
