Late Binding:  VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where version inconsistencies might occur.

Example:
```vbscript
Set obj = CreateObject("Some.Object.That.Might.NotExist")
```
This will only throw an error at runtime if `Some.Object.That.Might.NotExist` is not available.