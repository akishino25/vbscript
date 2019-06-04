' json -> csv 変換
' jsonが階層構造、もしくはリストだった時はどうする？
'  -> リストだったらそもそもCSVにできない？
'  -> 構造体だったらどうしようかな・・・
'
' とってきたJSONからキー一覧をどうにか取得するか？あらかじめキー一覧を定義するか？
'  -> キー一覧がある前提で実装する
'    -> キー一覧はまずはハードコーディングするが、もしかしたら設定ファイルにするかも
'  -> キー一覧をハードコーディングしたくなければ、JSONからkey一覧を抽出するようにするとか？

Option Explicit

Dim objFso, objFile
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.OpenTextFile("simple.json", 1, False)

Dim jsonData
If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    jsonData = objFile.readall
End If
objFile.Close

WScript.Echo jsonData

' JSONパース処理
Dim doc
Set doc = CreateObject("htmlfile")
doc.write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"
doc.write "<script>document.JsonParse=function (s) {return eval('(' + s + ')');}</script>"

Dim jsonParseData
Set jsonParseData = doc.JsonParse(jsonData)
WScript.Echo jsonParseData.key2
WScript.Echo Eval("jsonParseData." &"key3")

' JSONから値取り出し
Dim keyList
keyList = Array("key1","key2","key3","key4") ' key一覧を固定
WScript.Echo Join(keyList, ",")
Dim valList()
ReDim valList(UBound(keyList))

Dim i
For i=0 To UBound(keyList)
    valList(i) = Eval("jsonParseData." &keyList(i))
Next
WScript.Echo Join(valList, ",")

' CSVファイル書き出し
Dim outputFile
Set outputFile = objFso.OpenTextFile("simple.csv", 2, True)
If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    outputFile.Write joinList(keyList) & vbCrLf
    outputFile.Write JoinList(valList) & vbCrLf
End If
outputFile.Close

Set objFile = Nothing
Set objFso = Nothing

Function joinList(inputList)
    Dim i
    For i=0 To UBound(inputList)
        ' [TODO] CSV用に"のエスケープ処理が必要
        If i=0 Then
            joinList = """" &inputList(i) &""""
        Else
            joinList = joinList & ",""" &inputList(i) &""""
        End If 
    Next

End Function