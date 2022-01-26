Attribute VB_Name = "Json"
Option Compare Database
Option Explicit

' Author: Michele Locati <michele@locati.it>
' Source: https://github.com/mlocati/vba-json
' License: MIT

Public Const JSONDATATYPE_INVALID = 0&
Public Const JSONDATATYPE_SCALAR = 1&
Public Const JSONDATATYPE_ARRAY = 2&
Public Const JSONDATATYPE_OBJECT = 3&

Private Const JSVARTYPE_SCALAR = "scalar"
Private Const JSVARTYPE_ARRAY = "array"
Private Const JSVARTYPE_OBJECT = "object"
Public Function ParseJSON(ByVal json As String) As JsonData
    Dim jsEngine As Object
    Dim rawObject As Object, rawScalar As Variant
    Dim result As JsonData

    On Error GoTo 0

    Set jsEngine = CreateObject("ScriptControl")
    jsEngine.Language = "JScript"
    jsEngine.AddCode "function getVariableType(v, key) { " & _
        "var undef;" & _
        "if (key !== undef) v = v[key];" & _
        "if (v === undef) return '" & JSVARTYPE_SCALAR & "';" & _
        "if (v === null) return '" & JSVARTYPE_SCALAR & "';" & _
        "if (v instanceof Array) return '" & JSVARTYPE_ARRAY & "';" & _
        "if ('|boolean|number|string|'.indexOf('|' + (typeof v) + '|') >= 0) return '" & JSVARTYPE_SCALAR & "';" & _
        "if (typeof v === 'object') return '" & JSVARTYPE_OBJECT & "';" & _
        "return v;" & _
    "}"
    jsEngine.AddCode "function getObjectElement(v, key) { " & _
        "return v[key];" & _
    "}"
    jsEngine.AddCode "function getArrayLength(v) { " & _
        "return v instanceof Array ? v.length : -1;" & _
    "}"
    jsEngine.AddCode "function getObjectKeys(v) { " & _
        "var result = [];" & _
        "for (var k in v) result.push(k);" & _
        "return result.join('\x00');" & _
    "}"
    On Error Resume Next
    Err.Clear
    Set rawObject = jsEngine.Eval("(" + json + ")")
    If Err.Number = 0& Then
        On Error GoTo 0
        Set ParseJSON = ParseObject(jsEngine, rawObject)
    Else
        Err.Clear
        rawScalar = jsEngine.Eval("(" + json + ")")
        If Err.Number = 0& Then
            On Error GoTo 0
            Set result = New JsonData
            result.InitializeAsScalar rawScalar
            Set ParseJSON = result
        Else
            Err.Clear
            On Error GoTo 0
            Set ParseJSON = New JsonData
        End If
    End If
End Function
Private Function ParseObject(ByRef jsEngine As Object, ByRef rawObject As Object) As JsonData
    Dim rawObjectType As String, index As Long, Keys() As String, HasKeys As Boolean, subType As String, dictionary As Object, subItem As JsonData
    Dim result As JsonData
    
    rawObjectType = jsEngine.Run("getVariableType", rawObject)
    Select Case rawObjectType
        Case JSVARTYPE_ARRAY
            Dim maxArrayIndex As Long
            maxArrayIndex = jsEngine.Run("getArrayLength", rawObject) - 1
            If maxArrayIndex < 0 Then
                HasKeys = False
            Else
                HasKeys = True
                ReDim Keys(maxArrayIndex)
                For index = 0 To maxArrayIndex
                    Keys(index) = CStr(index)
                Next
            End If
        Case JSVARTYPE_OBJECT
            Dim serializedKeys As String
            serializedKeys = jsEngine.Run("getObjectKeys", rawObject)
            If serializedKeys = "" Then
                HasKeys = False
            Else
                HasKeys = True
                Keys = Split(serializedKeys, vbNullChar, -1, vbBinaryCompare)
            End If
        Case Else
            Err.Raise 1000, "Unexpected object type: " & rawObjectType
    End Select
    Set dictionary = CreateObject("Scripting.Dictionary")
    If HasKeys Then
        For index = LBound(Keys) To UBound(Keys)
            subType = jsEngine.Run("getVariableType", rawObject, Keys(index))
            Select Case subType
                Case JSVARTYPE_SCALAR
                    Set subItem = New JsonData
                    subItem.InitializeAsScalar jsEngine.Run("getObjectElement", rawObject, Keys(index))
                Case JSVARTYPE_OBJECT, JSVARTYPE_ARRAY
                    Set subItem = ParseObject(jsEngine, jsEngine.Run("getObjectElement", rawObject, Keys(index)))
                Case Else
                    Err.Raise 1000, "Unexpected object type: " & rawObjectType
            End Select
            dictionary.Add Keys(index), subItem
        Next
    End If
    Set result = New JsonData
    Select Case rawObjectType
        Case JSVARTYPE_ARRAY
            result.InitializeAsArray dictionary
        Case JSVARTYPE_OBJECT
            result.InitializeAsObject dictionary
        Case Else
            Err.Raise 1000, "Unexpected object type: " & rawObjectType
    End Select
    Set ParseObject = result
End Function
