VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Author: Michele Locati <michele@locati.it>
' Source: https://github.com/mlocati/vba-json
' License: MIT

Private myDataType As Long
Private myScalarValue As Variant
Private myDictionary As Object
Private myArrayLength As Long
Private myObjectHasKeys As Boolean
Private myObjectKeys() As String
Private Sub Class_Initialize()
    myDataType = JSONDATATYPE_INVALID
    myArrayLength = -1&
End Sub
Public Property Get IsValid() As Boolean
    IsValid = Me.DataType <> JSONDATATYPE_INVALID
End Property
Public Property Get DataType() As Long
    DataType = myDataType
End Property
Public Sub InitializeAsScalar(ByVal value As Variant)
    myDataType = JSONDATATYPE_SCALAR
    myScalarValue = value
End Sub
Public Property Get IsScalar() As Boolean
    IsScalar = Me.DataType = JSONDATATYPE_SCALAR
End Property
Public Property Get ScalarValue() As Variant
    If Me.IsScalar Then
        ScalarValue = myScalarValue
    End If
End Property
Public Sub InitializeAsArray(ByRef dictionary As Object)
    myDataType = JSONDATATYPE_ARRAY
    Set myDictionary = dictionary
    myArrayLength = UBound(dictionary.Keys) + 1&
End Sub
Public Property Get IsArray() As Boolean
    IsArray = Me.DataType = JSONDATATYPE_ARRAY
End Property
Public Property Get ArrayLength() As Long
    ArrayLength = myArrayLength
End Property
Public Function GetArrayItem(ByVal index As Long) As JsonData
    Set GetArrayItem = New JsonData
    If Me.IsArray Then
        Dim key As String
        key = CStr(index)
        If myDictionary.Exists(key) Then
            Set GetArrayItem = myDictionary.Item(key)
        End If
    End If
End Function
Public Sub InitializeAsObject(ByRef dictionary As Object)
    Dim Keys As Variant, numKeys As Long, minIndex As Long, maxIndex As Long
    myDataType = JSONDATATYPE_OBJECT
    Set myDictionary = dictionary
    Keys = dictionary.Keys
    minIndex = LBound(Keys)
    maxIndex = UBound(Keys)
    numKeys = maxIndex - minIndex + 1&
    If numKeys > 0& Then
        Dim index As Long
        myObjectHasKeys = True
        ReDim myObjectKeys(numKeys - 1)
        For index = minIndex To maxIndex
            myObjectKeys(index - minIndex) = Keys(index)
        Next
    Else
        myObjectHasKeys = False
    End If
End Sub
Public Property Get IsObject() As Boolean
    IsObject = Me.DataType = JSONDATATYPE_OBJECT
End Property
Public Property Get ObjectHasKeys() As Boolean
    ObjectHasKeys = myObjectHasKeys
End Property
Public Property Get ObjectKeys() As String()
    If Me.IsObject Then
        ObjectKeys = myObjectKeys
    End If
End Property
Public Function GetObjectItem(ByVal key As String) As JsonData
    Set GetObjectItem = New JsonData
    If Me.IsObject Then
        If myDictionary.Exists(key) Then
            Set GetObjectItem = myDictionary.Item(key)
        End If
    End If
End Function
Public Function GetChildByPath(ByVal path As String) As JsonData
    Dim result As JsonData
    If Not Me.IsValid Then
        Set GetChildByPath = New JsonData
        Exit Function
    End If
    path = "" & path
    If path = "" Then
        Set GetChildByPath = Me
        Exit Function
    End If
    If Me.IsScalar Then
        Set GetChildByPath = New JsonData
        Exit Function
    End If
    Dim p As Long
    Dim myKey As String, subKey As String
    Dim child As JsonData
    p = InStr(1&, path, ".", vbBinaryCompare)
    If p = 0& Then
        myKey = path
        subKey = ""
    Else
        myKey = Left$(path, p - 1&)
        subKey = Mid$(path, p + 1&)
    End If
    If myKey = "" Then
        Set child = Me
    ElseIf Me.IsObject Then
        Set child = Me.GetObjectItem(myKey)
    ElseIf Me.IsArray Then
        If IsNumeric(myKey) Then
            Dim myIndex As Long
            myIndex = CLng(myKey)
            Set child = Me.GetArrayItem(myIndex)
        Else
            Set child = New JsonData
        End If
    Else
        Set child = New JsonData
    End If
    If child.IsValid Then
        If subKey = "" Then
            Set result = child
        Else
            Set result = child.GetChildByPath(subKey)
        End If
    Else
        Set result = child
    End If
    Set GetChildByPath = result
End Function
Public Function ToJSON(Optional ByVal IndentWith As String = vbTab, Optional ByVal Depth As Long = 0&, Optional ByVal NewLineSequence As String = vbNewLine) As String
    Dim prefix As String, index As Long, num As Long, key As Variant, isFirst As Boolean
    prefix = ""
    For index = 1 To Depth
        prefix = prefix & IndentWith
    Next
    Select Case Me.DataType
        Case JSONDATATYPE_INVALID
            ToJSON = ""
        Case JSONDATATYPE_SCALAR
            Select Case VarType(Me.ScalarValue)
                Case vbNull
                    ToJSON = "null"
                Case vbBoolean
                    If Me.ScalarValue Then
                        ToJSON = "true"
                    Else
                        ToJSON = "false"
                    End If
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                    ToJSON = CStr(Me.ScalarValue)
                Case vbString
                    ToJSON = StringToJSON(Me.ScalarValue)
                Case Else
                    Err.Raise 1001, "Unrecognized variable type"
            End Select
        Case JSONDATATYPE_ARRAY
            If Me.ArrayLength = 0& Then
                ToJSON = "[]"
            Else
                ToJSON = "[" & NewLineSequence
                For index = 0 To Me.ArrayLength - 1&
                    ToJSON = ToJSON & prefix & IndentWith & Me.GetArrayItem(index).ToJSON(IndentWith, Depth + 1&, NewLineSequence)
                    If index < Me.ArrayLength - 1& Then
                        ToJSON = ToJSON & ","
                    End If
                    ToJSON = ToJSON & NewLineSequence
                Next
                ToJSON = ToJSON & prefix & "]"
            End If
        Case JSONDATATYPE_OBJECT
            If Not Me.ObjectHasKeys Then
                ToJSON = "{}"
            Else
                ToJSON = "{" & NewLineSequence
                isFirst = True
                For Each key In Me.ObjectKeys
                    If isFirst Then
                        isFirst = False
                    Else
                        ToJSON = ToJSON & "," & NewLineSequence
                    End If
                    ToJSON = ToJSON & prefix & IndentWith & StringToJSON(key) & ": " & Me.GetObjectItem(key).ToJSON(IndentWith, Depth + 1&, NewLineSequence)
                Next
                ToJSON = ToJSON & NewLineSequence & prefix & "}"
            End If
        Case Else
            Err.Raise 1001, "Unrecognized data type"
    End Select
End Function
Private Function StringToJSON(ByVal str As String) As String
    Dim l As Long, index As Long, c As String * 1, code As Long
    ' See https://www.ietf.org/rfc/rfc4627.txt
    StringToJSON = """"
    l = Len("" & str)
    For index = 1& To l
        c = Mid$(str, index, 1&)
        Select Case c
            Case "\"
                StringToJSON = StringToJSON & "\\"
            Case """"
                StringToJSON = StringToJSON & "\"""
            Case vbBack '\x08
                StringToJSON = StringToJSON & "\b"
            Case vbTab '\x09
                StringToJSON = StringToJSON & "\t"
            Case vbLf '\x0A
                StringToJSON = StringToJSON & "\n"
            Case vbFormFeed '\x0C
                StringToJSON = StringToJSON & "\f"
            Case vbCr '\x0D
                StringToJSON = StringToJSON & "\r"
            Case Else
                code = Asc(c)
                If code <= 31 Then
                    StringToJSON = StringToJSON & "\u" & Right$("000" & Hex$(c), 4&)
                Else
                    StringToJSON = StringToJSON & c
                End If
        End Select
    Next
    StringToJSON = StringToJSON & """"
End Function
