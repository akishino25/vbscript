Dim nJS
Dim oJS

Set nJS = CreateObject("ScriptControl")
    nJS.Language = "JScript"
    nJS.AddCode "function hang01(a,b){" &_
                "    var c = a *  b;" & _
                "  return '""a * b""= ""'  + c + '""';" & _ 
                _
                "}"
Set oJS = nJS.CodeObject
                
Msgbox oJS.hang01(123,3)