VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private stringBuffer As String
Private ptr As Long

Public ExtendSize As Long

Public Sub Append(str As String)
    If ptr + Len(str) > Len(stringBuffer) Then
        Extend ExtendSize + Len(str) 'Extend one unit + string size
    End If
    Mid$(stringBuffer, ptr, Len(str)) = str
    ptr = ptr + Len(str)
End Sub
Public Sub Replace(str As String, start As Long)
    'Replace part of the string being built with a different string, starting at start
    Dim slen As Long
    slen = Len(str)
    If slen + start < Len(stringBuffer) Then
        Extend slen + ExtendSize
    End If
    If slen + start < ptr Then
        ptr = slen + start
    End If
    Mid$(stringBuffer, start, slen) = str
End Sub

Public Sub ReplaceLast(str As String) 'Replace the last character of the string being built
    Mid$(stringBuffer, ptr - 1, 1) = str
End Sub

Public Sub Extend(NumberOfChars As Long) 'Extend the string being built
    stringBuffer = stringBuffer & String(NumberOfChars, vbNullChar)
End Sub


Public Function ToString() As String
    'Return the string being built
    ToString = Left(stringBuffer, ptr - 1)
End Function

Private Sub Class_Initialize()
    'Initialize defaults
    ExtendSize = 1024
    ptr = 1
End Sub

Public Property Get StringLength() As Long
    StringLength = ptr - 1
End Property

Public Function Mid(start As Long, Optional length As Long = -1) As String
    'Get a substring from the string builder
    If length = -1 Then length = ptr - start - 1
    Mid = VBA.Mid$(stringBuffer, length, length)
End Function

Public Function right(length As Long) As String
    'Get the rightmost part from the string in the stringbuilder
    right = VBA.Mid$(stringBuffer, ptr - length, length)
End Function