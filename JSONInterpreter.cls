VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'JSONObject, dynamic JSON encoder and decoder
'Erik von Asmuth, 2018

'Segments (Ctrl + F):
'COMMON PROPERTIES PART
'COMMON METHODS PART
'PARENT-CHILD COMMUNICATION PART
'COMMON JSON FUNCTIONS PART
'JSON VALUE DECODERS PART
'JSON VALUE SIZE PART
'JSON MEMBERS PART
'JSON ENCODER PART

Private JSONString As String
Private ParentObject As JSONInterpreter
Private memberPointers As Object 'Scripting.Dictionary
Private MemberIdentifier As Long
Private endPointer As Long
Private Enum valueType
    vObject = 1
    vArray = 2
    vString = 3
    vNumber = 4
    vBoolean = 5
    vNull = 6
    vInvalid = -1
End Enum

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------   COMMON PROPERTIES PART -------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'Common properties to instantiate and get variables

Public Property Get JSON() As String
    JSON = JSONString
End Property

Public Property Let JSON(InpJSON As String)
    JSONString = InpJSON
    If ParentObject Is Nothing Then
        'Instantiante new object from nothing
        InvalidateMembers
    Else
        ParentObject.ReplaceValue MemberIdentifier, InpJSON
    End If
End Property

Public Property Get VBAVariant() As Variant
    Dim ret As Variant
    On Error GoTo ErrHandler
    parseValue 1, ret   'Start parsing at start of string
    On Error GoTo 0
    If IsObject(ret) Then
        Set VBAVariant = ret
    Else
        'If not object, return collection containing return value
        VBAVariant = ret
    End If
    Exit Property
ErrHandler:
    Err.Raise vbObjectError - 1, Description:="Error when parsing JSON! Make sure to only pass valid JSON"
End Property

Public Property Let VBAVariant(InputVariable As Variant)
    Dim sb As New clsStringBuilder
    JSONAppendValue InputVariable, sb
    JSON = sb.ToString
End Property

Public Property Set VBAVariant(InputVariable As Variant)
    Dim sb As New clsStringBuilder
    JSONAppendValue InputVariable, sb
    JSON = sb.ToString
End Property

Public Property Get valueType() As String
    Select Case myType
        Case vObject
            valueType = "Object"
        Case vString
            valueType = "String"
        Case vArray
            valueType = "Array"
        Case vBoolean
            valueType = "Boolean"
        Case vNull
            valueType = "Null"
        Case vNumber
            valueType = "Number"
        Case vInvalid
            valueType = "Invalid"
    End Select
End Property

Public Property Get IsComplete() As Boolean
    'If I end at the end of my string, true, else, false
    Dim EOV As Long
    EOV = EndOfValue(1)
    If EOV > 0 And EOV <= Len(JSONString) + 1 Then
        IsComplete = True
    End If
End Property

Public Property Get IsEmpty() As Boolean
    'If I'm an array or object that contains 0 items, return true, else false
    'Note: a 0-element unclosed array/object is not empty since it might contain values
    If myType = vObject Then
        IsEmpty = Mid$(JSONString, MoveNextChar(MoveNextChar(1) + 1), 1) = "}"
    ElseIf myType = vArray Then
        IsEmpty = Mid$(JSONString, MoveNextChar(MoveNextChar(1) + 1), 1) = "]"
    End If
End Property


Public Property Get item(ParamArray Position() As Variant) As JSONInterpreter 'Get a certain item
    Dim l As Long
    Dim curPointer As Long
    curPointer = 1
    Dim v As Variant
    For Each v In Position
        If IsObject(v) Then v = v 'If object then default property
        Dim sp As Long
        If (VarType(v) = vbString) And getType(curPointer) = vObject Then 'Key lookup
            curPointer = findStartpointerByKey(CStr(v), curPointer)
            If Not (curPointer > 0 And curPointer <= Len(JSONString)) Then Exit Property
        ElseIf IsNumeric(v) Then
            curPointer = memberStartPointer(CLng(v), curPointer)
            If Not (curPointer > 0 And curPointer <= Len(JSONString)) Then Exit Property
        Else
            Err.Raise vbObjectError, Description:="Invalid input in member lookup!"
        End If
    Next
    If curPointer > 0 And curPointer <= Len(JSONString) Then
        Set item = NewMember(curPointer)
    End If
End Property

Public Property Get Key(Position As Long) As String 'Get the key for a certain item
    If myType <> vObject Then
        Err.Raise vbObjectError - 1, Description:="Attempted to access key of non-object"
    End If
    Dim sp As Long
    sp = keyStartPointer(Position, 1)
    If sp > 0 And sp <= Len(JSONString) Then
        Key = parseString(sp)
    End If
End Property

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------   COMMON METHODS PART ----------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Public methods for general use

Public Function Exists(ParamArray Position() As Variant) As Boolean 'Test if a certain item exists
    Dim l As Long
    Dim curPointer As Long
    curPointer = 1
    Dim v As Variant
    For Each v In Position
        If IsObject(v) Then v = v 'If object then default property
        Dim sp As Long
        If (VarType(v) = vbString) And getType(curPointer) = vObject Then 'Key lookup
            curPointer = findStartpointerByKey(CStr(v), curPointer)
            If Not (curPointer > 0 And curPointer <= Len(JSONString)) Then Exit Function
        ElseIf IsNumeric(v) Then
            curPointer = memberStartPointer(CLng(v), curPointer)
            If Not (curPointer > 0 And curPointer <= Len(JSONString)) Then Exit Function
        Else
            Err.Raise vbObjectError, Description:="Invalid input in member lookup!"
        End If
    Next
    Exists = True
End Function

Public Sub Refresh() 'Get current JSON from parent
    If ParentObject Is Nothing Then Exit Sub
    JSONString = ParentObject.GetMemberJSON(MemberIdentifier)
End Sub

Public Sub Unlink() 'Break link from parent object
    Set ParentObject = Nothing
End Sub

Public Property Get Count() As Long 'Get the number of items
    If myType = vArray Then
        Count = AmountOfItemsInArray(1)
    ElseIf myType = vObject Then
        Count = AmountOfItemsInObject(1)
    Else
        Count = 0
    End If
End Property

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------   PARENT-CHILD COMMUNICATION PART -----------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Methods in here shouldn't get called

Public Property Set Parent(objParent As JSONInterpreter)
    If ParentObject Is Nothing And JSONString = "" Then
        Set ParentObject = objParent
    End If
End Property

Public Sub ReplaceValue(valueIdentifier As Long, newJSON As String)
    Dim startString As Long
    startString = GetMemberPointer(valueIdentifier)
    Dim endString As Long
    endString = EndOfValue(startString)
    Dim offset As Long
    offset = Len(newJSON) - (endString - startString)
    JSONString = Left(JSONString, startString - 1) & newJSON & Mid$(JSONString, endString)
    Dim l As Long
    For l = 1 To memberPointers.Count
        If memberPointers(l) <= startString Then
            'Do nothing
        ElseIf memberPointers(l) < endString Then
            memberPointers(l) = -1 'Pointer to value that got replaced, invalidate
        Else
            memberPointers(l) = (memberPointers(l) + offset) 'Calculate new pointer
        End If
    Next
End Sub

Private Sub InvalidateMembers()
    If Not memberPointers Is Nothing Then
        Dim l As Long
        For l = 1 To memberPointers.Count
            memberPointers(l) = -1
        Next
    End If
End Sub

Public Function GetMemberPointer(id As Long) As Long
    GetMemberPointer = memberPointers(id)
End Function

Public Function GetMemberJSON(id As Long) As String
    Dim p As Long
    p = memberPointers(id)
    Dim e As Long
    e = EndOfValue(p)
    If p = -1 Or e = -1 Or e < p Then Exit Function
    GetMemberJSON = Mid$(JSONString, p, e - p)
End Function

Private Property Get MyStartPointer() As Long
    If ParentObject Is Nothing Then
        MyStartPointer = 1
    Else
        MyStartPointer = ParentObject.GetMemberPointer(MemberIdentifier)
    End If
End Property

Public Property Let Identifier(inp As Long)
    If MemberIdentifier = 0 And Not ParentObject Is Nothing Then 'If I'm a member and don't have an identifier
        MemberIdentifier = inp 'Get my identifier
    Else
        Err.Raise vbObjectError - 1, Description:="Attempt to set member identifier of parent or attempt to overwrite member identifier" 'Else, error
    End If
End Property

Public Function NewMember(startPointer As Long) As JSONInterpreter
    If ParentObject Is Nothing Then 'If I'm the parent
        'Create member dictionary if not exists
        If memberPointers Is Nothing Then Set memberPointers = CreateObject("Scripting.Dictionary")
        Set NewMember = New JSONInterpreter 'Create new member
        Dim id As Long
        id = memberPointers.Count + 1
        memberPointers.Add Key:=id, item:=startPointer 'Keep track of the pointer
        Set NewMember.Parent = Me 'Set the parent
        NewMember.Identifier = id 'Assign identifier to member, must come after Set NewMember = Parent
        NewMember.Refresh 'Let member get their JSON themselves using their identifier
    Else
        Set NewMember = ParentObject.NewMember(startPointer + MyStartPointer - 1) 'Ask parent to create member, calculate pointer position
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------   COMMON JSON FUNCTIONS PART ----------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Moving pointers, getting types, applicable to every JSON object


Private Function myType() As valueType
    If Len(JSONString) = 0 Then
        myType = vInvalid
    Else
        myType = getType(MoveNextChar(1))
    End If
End Function

Private Function getType(ByVal startPointer As Long) As valueType
    If Len(JSONString) < startPointer Then
        getType = vInvalid
        Exit Function
    End If
    Select Case AscW(Mid$(JSONString, startPointer, 1))
        Case 123 '{ -> Object
            getType = vObject
        Case 34 'Double quote -> String
            getType = vString
        Case 91 'Square bracket -> Array
            getType = vArray
        Case 84, 116, 70, 102 't T f F - > Boolean
            getType = vBoolean
        Case 78, 110  'n N -> Null
            getType = vNull
        Case 43, 45, 46, 48 To 57, 69, 101 'Valid JSON numerical characters: +, -, ., 0-9, E, e -> Number
            getType = vNumber
        Case Else
            getType = vInvalid
    End Select
End Function


Private Function MoveNextChar(ByVal Pointer As Long) As Long
    'Moves startPointer to next non-whitespace character
    Dim LenJSON As Long
    LenJSON = Len(JSONString)
    Do While IsWhitespace(Pointer) And Pointer < LenJSON
        Pointer = Pointer + 1
    Loop
    MoveNextChar = Pointer
End Function

Private Function MovePreviousChar(ByVal Pointer As Long) As Long
    'Moves endPointer to the previous character
    Do While IsWhitespace(Pointer) And Not Pointer < 0
        Pointer = Pointer - 1
    Loop
    MovePreviousChar = Pointer
End Function

Private Function IsWhitespace(ByVal Pointer As Long) As Boolean
    If Pointer = 0 Then Stop
    'Is the character the pointer points to a whitespace character?
    Select Case AscW(Mid$(JSONString, Pointer, 1))
        Case 9, 10, 13, 32 'Valid ECMA-404 whitespace characters
            IsWhitespace = True
        Case Else
            IsWhitespace = False
    End Select
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------   JSON VALUE DECODERS PART ------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Functions that take a pointer to a JSON value, and return a VBA variant

Private Function parseValue(ByVal startPointer As Long, ByRef ReturnValue As Variant)
    startPointer = MoveNextChar(startPointer)
    Select Case AscW(Mid$(JSONString, startPointer, 1))
        Case 123 '{ -> Object
            Set ReturnValue = parseObject(startPointer)
        Case 34 'Double quote -> String
            ReturnValue = parseString(startPointer)
        Case 91 'Square bracket -> Array
            ReturnValue = parseArray(startPointer)
        Case 84, 116, 70, 102, 78, 110 't T f F n N -> True, False, Null
            ReturnValue = parseTFN(startPointer)
        Case 43, 45, 46, 48 To 57, 69, 101 'Valid JSON numerical characters: +, -, ., 0-9, E, e -> Number
            ReturnValue = parseNumber(startPointer)
        Case Else
            'Unknown JSON value!
            Err.Raise vbObjectError - 1, Description:="Error parsing: invalid JSON value. Start of value: " & Mid$(JSONString, startPointer, 10)
            endPointer = startPointer + 1
    End Select
End Function

Private Function parseObject(ByVal startPointer As Long) As Object
    Dim ret As Object
    Set ret = CreateObject("Scripting.Dictionary")
    startPointer = MoveNextChar(startPointer + 1) 'Move pointer past starting curly brace to first opening quote
    Do While Mid$(JSONString, startPointer, 1) <> "}" 'While is a bit redundant for valid JSON, could only be evaluated once
        'Parse key:
        Dim keyReturn As String
        keyReturn = parseString(startPointer)
        startPointer = MoveNextChar(endPointer)
        If Not Mid$(JSONString, startPointer, 1) = ":" Then
            Err.Raise vbObjectError - 1, Description:="Missing colon character, key-value pair expected in JSON object but not supplied"
        End If
        startPointer = MoveNextChar(startPointer + 1)
        Dim valReturn As Variant
        parseValue startPointer, valReturn
        ret.Add Key:=keyReturn, item:=valReturn
        startPointer = MoveNextChar(endPointer) 'Move pointer to next character
        If Mid$(JSONString, startPointer, 1) <> "," Then Exit Do 'If no comma then we have the whole object
        startPointer = MoveNextChar(startPointer + 1) 'Skip over comma, next character
    Loop
    Set parseObject = ret
    endPointer = startPointer + 1 'Skip over ending curly brace
End Function

Private Function memberStartPointer(MemberNumber As Long, startPointer As Long) As Long
    If getType(startPointer) = vArray Then
        memberStartPointer = arrayMemberStartPointer(MemberNumber, startPointer)
    ElseIf getType(startPointer) = vObject Then
        memberStartPointer = objectMemberStartPointer(MemberNumber, startPointer)
    Else 'Invalid pointer
        memberStartPointer = Len(JSONString) + 1 'Pointer outside string bound
    End If
End Function

Private Function parseString(ByVal startPointer As Long) As Variant
    'Parses a string. StartPointer = start of string
    startPointer = startPointer + 1 'No first quote
    Dim nextRevSolidus As Long
    Dim nextQuote As Long
    Dim escapeSequenceReturn As String
    Do
        nextRevSolidus = InStr(startPointer, JSONString, "\") 'Locate next escape sequence
        nextQuote = InStr(startPointer, JSONString, """") 'Locate next quote
        If nextRevSolidus <> 0 And nextRevSolidus < nextQuote Then 'Escape sequence in string, append part before, then append escape sequence
            escapeSequenceReturn = parseEscapeSequence(nextRevSolidus) 'Get result of escape sequenct
            parseString = parseString & Mid$(JSONString, startPointer, nextRevSolidus - startPointer) & escapeSequenceReturn 'Append part before escape sequence and escape sequence
            startPointer = endPointer 'Move pointer
        End If
    Loop While nextRevSolidus > 0 And nextQuote > nextRevSolidus
    parseString = parseString & Mid$(JSONString, startPointer, nextQuote - startPointer)
    endPointer = nextQuote + 1 'Move out of string
End Function


Private Function parseTFN(ByVal startPointer As Long) As Variant
    Select Case True
        Case Mid$(JSONString, startPointer, 4) = "true"
            endPointer = startPointer + 4
            parseTFN = True
        Case Mid$(JSONString, startPointer, 5) = "false"
            endPointer = startPointer + 5
            parseTFN = False
        Case Mid$(JSONString, startPointer, 4) = "null"
            endPointer = startPointer + 4
            parseTFN = Null
    End Select
End Function

Private Function parseArray(ByVal startPointer As Long) As Variant
    'Future consideration: half this code is dedicated to reliably determining amount of items in array, nearly all LengthOf functions are only used for this
    'If we either redim preserve or use a collection this code can be simplified immensely
    Dim arrNumberOfItems As Long
    arrNumberOfItems = AmountOfItemsInArray(startPointer) 'Get array size
    Dim returnArray() As Variant 'Declare return var
    If arrNumberOfItems = 0 Then 'If 0 items, return empty array
        parseArray = returnArray
        startPointer = MoveNextChar(startPointer)
        endPointer = startPointer + 1
        Exit Function
    End If
    startPointer = startPointer + 1 'Move past opening square bracket
    ReDim returnArray(arrNumberOfItems - 1) 'Init array
    Dim arrayItem As Variant
    Dim itemPointer As Long
    Do While Not Mid$(JSONString, MoveNextChar(startPointer), 1) = "]"
        parseValue startPointer, arrayItem  'Retrieve value
        If IsObject(arrayItem) Then 'Add to array
            Set returnArray(itemPointer) = arrayItem
        Else
            returnArray(itemPointer) = arrayItem
        End If
        itemPointer = itemPointer + 1 'Move array pointer 1 position
        startPointer = MoveNextChar(endPointer) 'Move pointer past end of last value
        If Mid$(JSONString, startPointer, 1) <> "," Then Exit Do 'If not comma, we're done, pointer should point to closing square bracket
        startPointer = startPointer + 1 'Move pointer past comma
    Loop
    'Return values
    parseArray = returnArray
    endPointer = startPointer + 1
End Function


Private Function parseNumber(ByVal startPointer As Long) As Double
    'Assumes start of pointer points to first character of a number, returns that number as a return value + double
    'Warning: this function can return an endpointer that's equal to the start pointer in case of invalid input
    Dim lenNumber As Long
    endPointer = EndOfNumber(startPointer)
    lenNumber = endPointer - startPointer
    parseNumber = val(Mid$(JSONString, startPointer, lenNumber))
End Function

Private Function parseEscapeSequence(ByVal startPointer As Long) As String
    'Input: ByRef pointer to a backslash inside a string starting an escape sequence
    'Result: pointer gets moved to the end of the escape sequence
    'First character is always a slash, second is the escape char or sequence
    startPointer = startPointer + 1
    Select Case Mid$(JSONString, startPointer, 1)
        Case "/", "\", """"
            parseEscapeSequence = Mid$(JSONString, startPointer, 1)
            endPointer = startPointer + 1
        Case "b"
            parseEscapeSequence = Chr$(8) 'Backspace
            endPointer = startPointer + 1
        Case "f"
            parseEscapeSequence = Chr$(12) 'Form feed
            endPointer = startPointer + 1
        Case "n"
            parseEscapeSequence = vbLf 'Line feed
            endPointer = startPointer + 1
        Case "r"
            parseEscapeSequence = vbCr 'Carriage return
            endPointer = startPointer + 1
        Case "t"
            parseEscapeSequence = vbTab 'Tab character
            endPointer = startPointer + 1
        Case "u"
            parseEscapeSequence = ChrW$(HexToDec(Mid$(JSONString, startPointer + 1, 4)))
            endPointer = startPointer + 5
        Case Else
            'Invalid escape sequence! Return backslach + first character following it
            parseEscapeSequence = Mid$(JSONString, startPointer - 1, 2)
            endPointer = startPointer + 1
            Err.Raise vbObjectError - 1, Description:="Invalid escape sequence in JSON string"
    End Select
End Function

Private Function HexToDec(str As String) As Long
    HexToDec = CLng("&H" & str) 'Quick convert hex to decimal
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------   JSON VALUE SIZE PART ------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Functions that determine the size of a JSON Value based on the start of the value

Private Function EndOfValue(ByVal startPointer As Long) As Long
    Select Case AscW(Mid$(JSONString, startPointer, 1))
        Case 123 '{ -> Object
            EndOfValue = EndOfObject(startPointer)
        Case 34 'Double quote -> String
            EndOfValue = EndOfString(startPointer)
        Case 91 'Square bracket -> Array
            EndOfValue = EndOfArray(startPointer)
        Case 84, 116, 70, 102, 78, 110 't T f F - > Boolean
            EndOfValue = EndOfLiteral(startPointer)
        Case 43, 45, 46, 48 To 57, 69, 101 'Valid JSON numerical characters: +, -, ., 0-9, E, e -> Number
            EndOfValue = EndOfNumber(startPointer)
        Case Else
            EndOfValue = -1
    End Select
End Function


Private Function EndOfString(ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of string variable, returns length of string
    startPointer = startPointer + 1 'Step into string (over opening quote)
    Dim nextQuote As Long
    Dim backslashCount As Long
    nextQuote = InStr(startPointer, JSONString, """") 'Locate next quote
    Do
        If nextQuote = 0 Then
            'Incomplete string, end is beyond EOS
            EndOfString = Len(JSONString) + 2
        End If
        backslashCount = 0
        Do While Mid$(JSONString, nextQuote - backslashCount - 1, 1) = "\" 'If there's a backslash in front of last quote
            backslashCount = backslashCount + 1 'Backtrack, if backslash before quote check if backslash is escaped
        Loop
        If backslashCount Mod 2 = 1 Then 'If uneven amount of backslashes before quote, quote is escaped
            nextQuote = InStr(nextQuote + 1, JSONString, """") 'Move to next quote
        Else
            Exit Do 'Next quote = end of string
        End If
    Loop
    EndOfString = nextQuote + 1 'Skip over closing quote
End Function

Private Function EndOfArray(ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of array variable, returns end of array
    'Start by handling zero-length object
    'Implementation very similar to EndOfObject
    Dim nextOpeningBracket As Long
    Dim nextClosingBracket As Long
    Dim nextDoubleQuote As Long
    startPointer = startPointer + 1 'Step over opening bracket
    Do
        If nextClosingBracket <= startPointer Then nextClosingBracket = InStr(startPointer, JSONString, "]")
        If nextClosingBracket = 0 Then
            EndOfArray = Len(JSONString) + 2 'End of array lies outside of JSON
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextDoubleQuote < nextClosingBracket And nextDoubleQuote < nextClosingBracket Then 'Next closing curly is possibly in string
            'Set pointer at closing quote of string
            startPointer = EndOfString(nextDoubleQuote)
        ElseIf nextOpeningBracket < nextClosingBracket Then 'Next closing bracket closes nested object
            startPointer = EndOfArray(nextOpeningBracket)
        Else
            Exit Do
        End If
    Loop
    EndOfArray = nextClosingBracket + 1 'Step over closing bracket
End Function

Private Function EndOfObject(ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of object variable, returns length of object
    startPointer = startPointer + 1 'Step into object
    Dim nextOpeningCurly As Long
    Dim nextClosingCurly As Long
    Dim nextDoubleQuote As Long
    Do
        If nextClosingCurly <= startPointer Then nextClosingCurly = InStr(startPointer, JSONString, "}")
        If nextClosingCurly = 0 Then
            EndOfObject = Len(JSONString) + 2 'End of object lies outside of JSON
            Exit Function
        End If
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextClosingCurly Then 'Next closing curly is possibly in string
            'Set pointer at closing quote of string
            startPointer = EndOfString(nextDoubleQuote) 'Set pointer at closing quote
        ElseIf nextOpeningCurly < nextClosingCurly Then 'Next closing curly closes nested object
            startPointer = EndOfObject(nextOpeningCurly)
        Else
            Exit Do
        End If
    Loop
    EndOfObject = nextClosingCurly + 1
End Function

Private Function EndOfNumber(ByVal startPointer As Long) As Long
    'Assumes start of pointer points to first character of a number, returns number of characters of that number
    Dim LenJSON As Long
    LenJSON = Len(JSONString)
    Do While startPointer <= LenJSON
        Select Case AscW(Mid$(JSONString, startPointer, 1))
            Case 43, 45, 46, 48 To 57, 69, 101 'Valid JSON numerical characters: +, -, ., 0-9, E, e
            Case Else
                Exit Do
        End Select
        startPointer = startPointer + 1
    Loop
    EndOfNumber = startPointer
End Function

Private Function EndOfLiteral(ByVal startPointer As Long) As Long
    Select Case Mid$(JSONString, startPointer, 1)
        Case "n", "t"
            EndOfLiteral = startPointer + 4
        Case "f"
            EndOfLiteral = startPointer + 5
    End Select
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------   JSON MEMBERS PART ---------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Functions that get the amount of members and the start pointer of members by position or key

Private Function AmountOfItemsInArray(ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of array variable (opening square bracket), returns amount of items in array
    Dim nextOpeningBracket As Long
    Dim nextClosingBracket As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    'Handle 0-item object first
    startPointer = startPointer + 1 'Step into array
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "]" Then
        Exit Function
    End If
    AmountOfItemsInArray = 1 'Amount of items = 1 + amount of comma's that are not in strings/arrays/objects
    Do
        'Initialize []{", pointers
        If nextClosingBracket <= startPointer Then nextClosingBracket = InStr(startPointer, JSONString, "]")
        If nextClosingBracket = 0 Then
            'Incomplete array
            AmountOfItemsInArray = -1
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        If nextComma <= startPointer Then nextComma = InStr(startPointer, JSONString, ",")
        If nextComma = 0 Then nextComma = Len(JSONString) + 1
        'Handle for lowest pointer
        'If " is lowest
        If nextClosingBracket < nextOpeningBracket And nextClosingBracket < nextDoubleQuote And nextClosingBracket < nextOpeningCurly And nextClosingBracket < nextComma Then
            Exit Function 'End of array, return amount
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly And nextOpeningBracket < nextComma Then 'Next item = nested array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextComma Then  'Next item = string
            startPointer = EndOfString(nextDoubleQuote) 'Skip over
        ElseIf nextOpeningCurly < nextComma Then 'Next item = object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        Else 'Next item = comma
            startPointer = nextComma + 1 'Move pointer over comma
            AmountOfItemsInArray = AmountOfItemsInArray + 1 'Increment amount of items by 1
        End If
    Loop
End Function

Private Function AmountOfItemsInObject(ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of array variable (opening curly brace), returns amount of items in array
    Dim nextOpeningBracket As Long
    Dim nextClosingCurly As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    'Handle 0-item array first
    startPointer = startPointer + 1 'Step into array
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "}" Then
        Exit Function
    End If
    AmountOfItemsInObject = 1 'Amount of items = 1 + amount of comma's that are not in strings/arrays/objects
    Do
        'Initialize [{}", pointers
        If nextClosingCurly <= startPointer Then nextClosingCurly = InStr(startPointer, JSONString, "}")
        If nextClosingCurly = 0 Then
            'Incomplete object
            AmountOfItemsInObject = -1
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        If nextComma <= startPointer Then nextComma = InStr(startPointer, JSONString, ",")
        If nextComma = 0 Then nextComma = Len(JSONString) + 1
        'Handle for lowest pointer
        'If " is lowest
        If nextClosingCurly < nextOpeningBracket And nextClosingCurly < nextDoubleQuote And nextClosingCurly < nextOpeningCurly And nextClosingCurly < nextComma Then
            Exit Function 'End of array, return amount
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly And nextOpeningBracket < nextComma Then 'Next item = array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextComma Then  'Next item = string
            startPointer = MoveNextChar(EndOfString(nextDoubleQuote)) 'Skip over, move next char to check :
            'Check if string is followed by :
            If Mid$(JSONString, startPointer, 1) = ":" Then startPointer = startPointer + 1
        ElseIf nextOpeningCurly < nextComma Then 'Next item = nested object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        Else 'Next item = comma
            startPointer = nextComma + 1 'Move pointer over comma
            AmountOfItemsInObject = AmountOfItemsInObject + 1 'Increment amount of items by 1
        End If
    Loop
End Function

Private Function objectMemberStartPointer(MemberNumber As Long, startPointer As Long) As Long
    'Assumes start pointer points to start of array variable (opening square bracket), returns amount of items in array
    Dim nextOpeningBracket As Long
    Dim nextClosingCurly As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    'Handle 0-item array first
    startPointer = startPointer + 1 'Step into object
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "}" Or MemberNumber <= 0 Then
        objectMemberStartPointer = Len(JSONString) + 1 'Invalid pointer
        Exit Function
    End If
    Dim itemCount As Long
    itemCount = 1 'Amount of items = 1 + amount of comma's that are not in strings/arrays/objects
    Do Until itemCount = MemberNumber
        'Initialize [{}", pointers
        If nextClosingCurly <= startPointer Then nextClosingCurly = InStr(startPointer, JSONString, "}")
        If nextClosingCurly = 0 Then nextClosingCurly = Len(JSONString) + 1
        If nextComma <= startPointer Then nextComma = InStr(startPointer, JSONString, ",")
        If nextComma = 0 Then nextComma = Len(JSONString) + 1
        If nextComma = nextClosingCurly Then 'No more comma's or closing curlies
            'Incomplete object
            objectMemberStartPointer = Len(JSONString) + 1 'Invalid pointer
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1

        'Handle for lowest pointer
        'If " is lowest
        If nextClosingCurly < nextOpeningBracket And nextClosingCurly < nextDoubleQuote And nextClosingCurly < nextOpeningCurly And nextClosingCurly < nextComma Then
            objectMemberStartPointer = Len(JSONString) + 1 'Invalid pointer
            Exit Function 'End of array, item count not hit yet
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly And nextOpeningBracket < nextComma Then 'Next item = array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextComma Then  'Next item = string
            startPointer = MoveNextChar(EndOfString(nextDoubleQuote)) 'Skip over, move next char to check :
            'Check if string is followed by :
            If Mid$(JSONString, startPointer, 1) = ":" Then startPointer = startPointer + 1
        ElseIf nextOpeningCurly < nextComma Then 'Next item = nested object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        Else 'Next item = comma
            startPointer = nextComma + 1 'Move pointer over comma
            itemCount = itemCount + 1 'Increment amount of items by 1
        End If
    Loop
    'Step over value in key-value pair
    startPointer = MoveNextChar(startPointer)
    'Assume valid pointer to string, skip over
    startPointer = MoveNextChar(MoveNextChar(EndOfString(startPointer)) + 1) 'Move to end of string pointer, then move to :, then to next char = start of member
    objectMemberStartPointer = startPointer 'Return current position
End Function

Private Function findStartpointerByKey(Key As String, ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of array variable (opening square bracket), returns amount of items in array
    Dim nextOpeningBracket As Long
    Dim nextClosingCurly As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    'Handle 0-item array first
    startPointer = startPointer + 1 'Step into object
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "}" Then
        findStartpointerByKey = Len(JSONString) + 1 'Invalid pointer
        Exit Function
    End If
    Do
        'Initialize [{}", pointers
        If nextClosingCurly <= startPointer Then nextClosingCurly = InStr(startPointer, JSONString, "}")
        If nextClosingCurly = 0 Then nextClosingCurly = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextDoubleQuote = nextClosingCurly Then 'No more quotes or closing curlies
            'Incomplete object, key not found before end, return invalid pointer
            Exit Do
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        'Handle for lowest pointer
        If nextClosingCurly < nextOpeningBracket And nextClosingCurly < nextDoubleQuote And nextClosingCurly < nextOpeningCurly Then 'At end of object
            Exit Do 'Key not found before end of object
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly Then 'Next item = array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly Then   'Next item = string
            Dim previousChar As Long 'Premature optimization, don't allocate string on heap
            previousChar = AscW(Mid$(JSONString, MovePreviousChar(nextDoubleQuote - 1), 1)) 'Get character code for previous character
            If previousChar = 44 Or previousChar = 123 Then 'String following { or , = key
                If parseString(nextDoubleQuote) = Key Then 'Key found!
                    findStartpointerByKey = MoveNextChar(MoveNextChar(endPointer) + 1) 'Skip over :, next char
                    Exit Function
                End If
            End If
            startPointer = MoveNextChar(EndOfString(nextDoubleQuote)) 'Skip over, move next char to check :
            'Check if string is followed by :
            If Mid$(JSONString, startPointer, 1) = ":" Then startPointer = startPointer + 1
        Else 'Next item = nested object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        End If
    Loop
    findStartpointerByKey = Len(JSONString) + 1 'Invalid pointer
End Function

Private Function keyStartPointer(MemberNumber As Long, ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of object variable (opening curly brace), returns key
    Dim nextOpeningBracket As Long
    Dim nextClosingCurly As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    'Handle 0-item array first
    startPointer = startPointer + 1 'Step into object
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "}" Or MemberNumber <= 0 Then
        keyStartPointer = Len(JSONString) + 1 'Invalid pointer
        Exit Function
    End If
    Dim itemCount As Long
    itemCount = 1 'Amount of items = 1 + amount of comma's that are not in strings/arrays/objects
    Do Until itemCount = MemberNumber
        'Initialize [{}", pointers
        If nextClosingCurly <= startPointer Then nextClosingCurly = InStr(startPointer, JSONString, "]")
        If nextClosingCurly = 0 Then nextClosingCurly = Len(JSONString) + 1
        If nextComma <= startPointer Then nextComma = InStr(startPointer, JSONString, ",")
        If nextComma = 0 Then nextComma = Len(JSONString) + 1
        If nextClosingCurly = nextComma Then
            'Incomplete object
            keyStartPointer = Len(JSONString) + 1 'Invalid pointer
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        
        'Handle for lowest pointer
        'If " is lowest
        If nextClosingCurly < nextOpeningBracket And nextClosingCurly < nextDoubleQuote And nextClosingCurly < nextOpeningCurly And nextClosingCurly < nextComma Then
            keyStartPointer = Len(JSONString) + 1 'Invalid pointer
            Exit Function 'End of array, item count not hit yet
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly And nextOpeningBracket < nextComma Then 'Next item = array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextComma Then  'Next item = string
            startPointer = MoveNextChar(EndOfString(nextDoubleQuote)) 'Skip over, move next char to check :
            'Check if string is followed by :
            If Mid$(JSONString, startPointer, 1) = ":" Then startPointer = startPointer + 1
        ElseIf nextOpeningCurly < nextComma Then 'Next item = nested object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        Else 'Next item = comma
            startPointer = nextComma + 1 'Move pointer over comma
            itemCount = itemCount + 1 'Increment amount of items by 1
        End If
    Loop
    'First object is key
    keyStartPointer = MoveNextChar(startPointer) 'Return current position
End Function

Private Function arrayMemberStartPointer(MemberNumber As Long, ByVal startPointer As Long) As Long
    'Assumes start pointer points to start of array variable (opening square bracket)
    Dim nextOpeningBracket As Long
    Dim nextClosingBracket As Long
    Dim nextDoubleQuote As Long
    Dim nextOpeningCurly As Long
    Dim nextComma As Long
    Dim itemCount As Long
    'Handle 0-item array first
    startPointer = startPointer + 1 'Step into array
    If Mid$(JSONString, MoveNextChar(startPointer), 1) = "]" Or MemberNumber <= 0 Then
        arrayMemberStartPointer = Len(JSONString) + 1 'Invalid pointer
        Exit Function
    End If
    itemCount = 1 'Amount of items = 1 + amount of comma's that are not in strings/arrays/objects
    Do Until itemCount = MemberNumber
        'Initialize []{", pointers
        If nextClosingBracket <= startPointer Then nextClosingBracket = InStr(startPointer, JSONString, "]")
        If nextClosingBracket = 0 Then nextClosingBracket = Len(JSONString) + 1
        If nextComma <= startPointer Then nextComma = InStr(startPointer, JSONString, ",")
        If nextComma = 0 Then nextComma = Len(JSONString) + 1
        If nextClosingBracket = nextComma Then
            'Incomplete array, beyond last comma
            arrayMemberStartPointer = Len(JSONString) + 1 'Invalid pointer
            Exit Function
        End If
        If nextOpeningBracket <= startPointer Then nextOpeningBracket = InStr(startPointer, JSONString, "[")
        If nextOpeningBracket = 0 Then nextOpeningBracket = Len(JSONString) + 1
        If nextDoubleQuote <= startPointer Then nextDoubleQuote = InStr(startPointer, JSONString, """")
        If nextDoubleQuote = 0 Then nextDoubleQuote = Len(JSONString) + 1
        If nextOpeningCurly <= startPointer Then nextOpeningCurly = InStr(startPointer, JSONString, "{")
        If nextOpeningCurly = 0 Then nextOpeningCurly = Len(JSONString) + 1
        'Handle for lowest pointer
        'If " is lowest
        If nextClosingBracket < nextOpeningBracket And nextClosingBracket < nextDoubleQuote And nextClosingBracket < nextOpeningCurly And nextClosingBracket < nextComma Then
            Exit Function 'End of array, return amount
        ElseIf nextOpeningBracket < nextDoubleQuote And nextOpeningBracket < nextOpeningCurly And nextOpeningBracket < nextComma Then 'Next item = nested array
            startPointer = EndOfArray(nextOpeningBracket) 'Skip over
        ElseIf nextDoubleQuote < nextOpeningCurly And nextDoubleQuote < nextComma Then  'Next item = string
            startPointer = EndOfString(nextDoubleQuote) 'Skip over
        ElseIf nextOpeningCurly < nextComma Then 'Next item = object
            startPointer = EndOfObject(nextOpeningCurly) 'Skip over
        Else 'Next item = comma
            startPointer = nextComma + 1 'Move pointer over comma
            itemCount = itemCount + 1 'Increment amount of items by 1
        End If
    Loop
    arrayMemberStartPointer = MoveNextChar(startPointer) 'Return current position
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------   JSON ENCODER PART ------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'The JSON encoder should be easily adjustable


'Encoding to string disabled
'Public Function JSONEncode(InputVariable As Variant) As String
'    Dim sb As New clsStringBuilder
'    JSONAppendValue InputVariable, sb
'    JSONEncode = sb.ToString
'End Function

Private Sub JSONAppendValue(InputVariable As Variant, StringBuilder As clsStringBuilder)
    If IsObject(InputVariable) Then
        Select Case TypeName(InputVariable)
            Case "Collection" 'Collections are appended as general enumerables -> arrays
                JSONAppendEnumerable InputVariable, StringBuilder
            Case "Dictionary"
                JSONAppendDictionary InputVariable, StringBuilder
            Case "Recordset", "Recordset2"
                JSONAppendRecordset InputVariable, StringBuilder
            Case Else
                JSONAppendUnidentifiedObject InputVariable, StringBuilder
        End Select
    ElseIf IsArray(InputVariable) Then 'Arrays are appended as general enumerables
        JSONAppendEnumerable InputVariable, StringBuilder
    Else 'Scalar type
        Select Case VarType(InputVariable)
            Case vbBoolean
                JSONApppendBoolean InputVariable, StringBuilder
            Case vbNull
                JSONAppendNull StringBuilder
            Case vbDate
                JSONAppendDate InputVariable, StringBuilder
            Case vbString
                JSONAppendString InputVariable, StringBuilder
            Case vbSingle, vbDouble, vbByte, vbInteger, vbLong, vbCurrency, 20 '20 = vbLongLong
                JSONAppendNumber InputVariable, StringBuilder
            Case Else
                JSONAppendUnidentifiedVariable InputVariable, StringBuilder
        End Select
    End If
End Sub

Private Sub JSONAppendRecordset(InputVariable As Variant, StringBuilder As clsStringBuilder)
    StringBuilder.Append "["
    If InputVariable.EOF And InputVariable.BOF Then
        StringBuilder.Append "}"
        Exit Sub
    End If
    With InputVariable
        .MoveFirst
        Do While Not .EOF
            Dim f As Variant
            StringBuilder.Append "{"
            For Each f In .Fields
                JSONAppendString f.Name, StringBuilder
                StringBuilder.Append ":"
                JSONAppendValue f.Value, StringBuilder
                StringBuilder.Append ","
            Next
            If StringBuilder.right(1) = "," Then
                StringBuilder.ReplaceLast "}"
            Else
                StringBuilder.Append "}"
            End If
            StringBuilder.Append ","
            .MoveNext
        Loop
    End With
    If StringBuilder.right(1) = "," Then
        StringBuilder.ReplaceLast "]"
    Else
        StringBuilder.Append "]"
    End If
End Sub

Private Sub JSONAppendString(InputVariable As Variant, StringBuilder As clsStringBuilder)
    Dim l As Long
    StringBuilder.Append """"
    For l = 1 To Len(InputVariable)
        Select Case AscW(Mid$(InputVariable, l, 1))
            'First: ASCII exemptions
            Case 8 'Backspace
                StringBuilder.Append "\b"
            Case 9 'Tab
                StringBuilder.Append "\t"
            Case 10 'Line feed
                StringBuilder.Append "\n"
            Case 12 'Form feed
                StringBuilder.Append "\f"
            Case 13 'Carriage return
                StringBuilder.Append "\r"
            Case 34 'Double quote
                StringBuilder.Append "\"""
            Case 92 'Reverse solidus
                StringBuilder.Append "\\"
            Case 0 To 127 'Remaining ASCII
                StringBuilder.Append Mid$(InputVariable, l, 1)
            Case Else 'Unicode escape
                StringBuilder.Append "\u" & right("000" & Hex(AscW(Mid$(InputVariable, l, 1))), 4)
        End Select
    Next
    StringBuilder.Append """"
End Sub

Private Sub JSONAppendDate(InputVariable As Variant, StringBuilder As clsStringBuilder)
    If InputVariable - Int(InputVariable) = 0 Then 'Date, no time
        JSONAppendString Format(InputVariable, "yyyy-mm-dd"), StringBuilder
    Else
        JSONAppendString Format(InputVariable, "yyyy-mm-ddThh:nn:ss"), StringBuilder
    End If
End Sub

Private Sub JSONAppendUnidentifiedVariable(InputVariable As Variant, StringBuilder As clsStringBuilder)
    'Try casting to string and appending as such
    Dim str As String
    On Error Resume Next
    str = InputVariable
    If Err.Number <> 0 Then
        str = "[variable " & TypeName(InputVariable) & "]"
    End If
    On Error GoTo 0
    JSONAppendString str, StringBuilder
End Sub

Private Sub JSONAppendDictionary(InputVariable As Variant, StringBuilder As clsStringBuilder)
    Dim l As Long
    StringBuilder.Append "{"
    For l = 0 To InputVariable.Count - 1
        JSONAppendString InputVariable.keys()(l), StringBuilder
        StringBuilder.Append ":"
        JSONAppendValue InputVariable.Items()(l), StringBuilder
        StringBuilder.Append ","
    Next
    If StringBuilder.right(1) = "," Then
        StringBuilder.ReplaceLast "}"
    Else
        StringBuilder.Append "}"
    End If
End Sub


Private Sub JSONAppendUnidentifiedObject(InputVariable As Variant, StringBuilder As clsStringBuilder)
    If IsEnumerable(InputVariable) Then
        JSONAppendEnumerable InputVariable, StringBuilder
    Else
        'Try casting to string and appending as such
        Dim str As String
        On Error Resume Next
        str = InputVariable
        If Err.Number <> 0 Then
            str = "[object " & TypeName(InputVariable) & "]"
        End If
        On Error GoTo 0
        JSONAppendString str, StringBuilder
    End If
End Sub

Private Sub JSONAppendEnumerable(InputVariable As Variant, StringBuilder As clsStringBuilder)
    StringBuilder.Append "["
    Dim v As Variant
    For Each v In InputVariable
        JSONAppendValue v, StringBuilder
        StringBuilder.Append ","
    Next
    If StringBuilder.right(1) = "," Then
        StringBuilder.ReplaceLast "]"
    Else
        StringBuilder.Append "]"
    End If
End Sub

Private Sub JSONAppendNumber(InputVariable As Variant, StringBuilder As clsStringBuilder)
    StringBuilder.Append Replace(InputVariable, GetDecimalSeparator, ".")
End Sub

Private Sub JSONApppendBoolean(InputVariable As Variant, StringBuilder As clsStringBuilder)
    If InputVariable Then
        StringBuilder.Append "true"
    Else
       StringBuilder.Append "false"
    End If
End Sub

Private Sub JSONAppendNull(StringBuilder As clsStringBuilder)
    StringBuilder.Append "null"
End Sub

Private Function IsEnumerable(var As Variant) As Boolean
    Dim v As Variant
    IsEnumerable = False
    On Error GoTo EndOfFunction
    For Each v In var
        Exit For
    Next
    IsEnumerable = True
EndOfFunction:
End Function

Public Function GetDecimalSeparator() As String
    GetDecimalSeparator = Format(0, ".")
End Function