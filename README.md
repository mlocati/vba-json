# Parse JSON strings in VBA (Visual Basic for Applications)

## Installation

1. Save the `Json.bas` and `JsonData.cls` files somewhere in your PC
2. Open the Visual Basic Editor
3. Choose the `File` > `Import File...` menu item
4. Import the two files saved in point 1.
5. You can safely delete the two files saved in point 1.

## Parsing JSON strings

Use the `ParseJSON` function to parse a string that contains data in JSON format.

This function always return a JsonData object.

```vb
Dim data As JsonData

Set data = ParseJSON("some json")
```

You can check its type with:

- `data.IsValid` returns `true` if the data is valid, `false` otherwise
- `data.IsScalar` returns `true` if the data is scalar (that is, `null`, a boolean, a number, a string), `false` otherwise
- `data.IsArray` returns `true` if the data is an array, `false` otherwise
- `data.IsObject` returns `true` if the data is an object, `false` otherwise
- `data.DataType` returns:
  - `JSONDATATYPE_SCALAR` if the data is scalar
  - `JSONDATATYPE_ARRAY` if the data is an array
  - `JSONDATATYPE_OBJECT` if the data is an object
  - `JSONDATATYPE_INVALID` if the data is not valid

## Invalid JSON

You can check if a JSON is valid with `IsValid`:

```vb
Dim data As JsonData

Set data = ParseJSON("");
' data.isValid is false

Set data = ParseJSON("{");
' data.isValid is false

Set data = ParseJSON("invalid json");
' data.isValid is false
```

## Scalar types

If `data.isScalar` returns `True` (or if `data.DataType` returns `JSONDATATYPE_SCALAR`), you can get the scalar value with `data.ScalarValue`.

Example:

```vb
Dim data As JsonData, value as Variant

Set data = ParseJSON("null")
If data.IsScalar Then
    value = data.ScalarValue ' It's Variant/Null - check it with IsNull()
End If

Set data = ParseJSON("true")
If data.IsScalar Then
    value = data.ScalarValue ' It's True
End If

Set data = ParseJSON("123")
If data.IsScalar Then
    value = data.ScalarValue ' It's 123
End If

Set data = ParseJSON("""Ciao!""")
If data.IsScalar Then
    value = data.ScalarValue ' It's "Ciao!"
End If
```

## Arrays

If `data.IsArray` returns `True` (or if `data.DataType` returns `JSONDATATYPE_ARRAY`), you can use the `ArrayLength` property and the `GetArrayItem()` method.

Example:

```vb
Dim data As JsonData, index As Long, item As JsonData

Set data = ParseJSON("[1, null, ""Hi"", true]")
If data.IsArray Then
    For index = 0 To data.ArrayLength - 1
        Set item = data.GetArrayItem(index)
        ' Work with item
    Next
End If
```

If you try to access an unexisting array index, `GetArrayItem()` will return a `JsonData` object whose `IsValid` property is `False`:

```vb
Dim data As JsonData, item As JsonData

Set data = ParseJSON("[0, 1]")
Set item = data.GetArrayItem(100)
' Here item.IsValid is False
```

## Objects

If `data.IsObject` returns `True` (or if `data.DataType` returns `JSONDATATYPE_OBJECT`), you can use the `ObjectHasKeys`/`ObjectKeys` properties and the `GetObjectItem()` method.

```vb
Dim data As JsonData, key As Variant, item As JsonData

Set data = ParseJSON("{""key1"": 1, ""key2"": 9}")
If data.IsObject And data.ObjectHasKeys Then
    For Each key In data.ObjectKeys
        Set item = data.GetObjectItem(key)
        ' Work with item
    Next
End If
```

If you try to access an unexisting object key, `GetObjectItem()` will return a `JsonData` object whose `IsValid` property is `False`:

```vb
Dim data As JsonData, item As JsonData

Set data = ParseJSON("{""good"": true}")
Set item = data.GetObjectItem("bad")
' Here item.IsValid is False
```

## Getting a value by path

You can use the `GetChildByPath()` method to quickly access object keys and array items.

For example, if you have this JSON data string stored in a variable named `json`:

```json
{
    "id": "AAA",
    "purchase": [
        {
            "reference": "MyOrder",
            "amount": {
                "currency": "EUR",
                "value": 123.45
            }
        }
    ]
}
```

You can retrieve the value of `value` with this code:

```vb
Dim data As JsonData, item As JsonData

Set data = ParseJSON(json)
Set item = data.GetChildByPath("purchase.0.amount.value")
' Here item.IsValid is False if the path didn't resolve to an element,
' and item.IsScalar is True if the element could be found and it's a scalar
If item.IsScalar Then
    ' ...
End If
```
