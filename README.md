# VbsJson

Code taken from: http://demon.tw/my-work/vbs-json.html

Given a **sample.json**:

```javascript
{
   "Image": {
       "Width":  800,
       "Height": 600,
       "Title":  "View from 15th Floor",
       "Thumbnail": {
           "Url":    "http://www.example.com/image/481989943",
           "Height": 125,
           "Width":  "100"
       },
       "IDs": [116, 943, 234, 38793]
     }
}
```

You can use it like:

```vbscript
Dim fso, json, str, o, i
Set json = New VbsJson
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
str = fso.OpenTextFile("sample.json").ReadAll
Set o = json.Decode(str)
WScript.Echo o("Image")("Width")
WScript.Echo o("Image")("Height")
WScript.Echo o("Image")("Title")
WScript.Echo o("Image")("Thumbnail")("Url")
For Each i In o("Image")("IDs")
    WScript.Echo i
Next
```
