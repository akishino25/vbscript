
Option Explicit

Dim filename
filename = "salesforceSample.json"

' jsonファイルを読み込む
On Error Resume Next
    Dim objFile, jsonData
    Set objFile = CreateObject("ADODB.Stream")
    objFile.Type = 2 'テキストデータ
    objFile.Charset = "UTF-8" '文字コードを指定
    objFile.Open 'Streamオブジェクトを開く
    objFile.LoadFromFile (filename) 'ファイルの内容を読み込む
    objFile.Position = 0 'ポインタを先頭へ
    jsonData = objFile.ReadText() 'データ読み込み
    objFile.Close 'Streamを閉じる
    Set objFile = Nothing 'オブジェクトの解放

    If Err.Number > 0 Then
        WScript.Echo "File Read Error"
        WScript.Quit
    End If
On Error Goto 0
MsgBox(jsonData)


' jsonファイルをパース
Dim objJSONTool, jsn
Set objJSONtool = CreateObject("HtmlFile")
objJSONtool.write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"
objJSONtool.write "<script>document.JsonParse=function (s) {return eval('(' + s + ')');}</script>"

On Error Resume Next
    Set jsn = objJSONtool.JsonParse(jsondata)
    If Err.Number > 0 Then
        WScript.Echo "JSON Parse Error"
        WScript.Quit
    End If
On Error Goto 0
MsgBox(jsn.totalSize)
MsgBox(jsn.records.[0].Key1)

' json -> csv変換
Dim keyList
keyList = Array("Name","key1","key2") ' key一覧を固定
Dim totalSize
totalSize = jsn.totalSize
Dim records()
ReDim records(totalSize, UBound(keyList))

Dim i, j
For i=0 To totalSize-1
    For j=0 To UBound(keyList)
        records(i,j) = Eval("jsn.records.[" &i &"]." &keyList(j))
    Next
Next

' CSVファイル書き出し
On Error Resume Next
    Dim outputFile
    outputfile = "salesforceSample.csv"

    Set objFile = CreateObject("ADODB.Stream")
    objFile.Type = 2 'テキストデータ
    objFile.Charset = "UTF-8" '文字コードを指定
    objFile.Open 'Streamオブジェクトを開く

    Dim header
    For i=0 To UBound(keyList)
        If i=0 Then
            header = """" &keyList(i) &""""
        Else
            header = header & ",""" &keyList(i) &""""
        End If 
    Next
    objFile.WriteText header & vbCrLf, 0

    Dim rec
    For i=0 To totalSize-1
        rec = ""
        For j=0 To UBound(keyList)
            ' [TODO] CSV用に"のエスケープ処理が必要
            If j=0 Then
                rec = """" &records(i,j) &""""
            Else
                rec = rec & ",""" &records(i,j) &""""
            End If
        Next
        objFile.WriteText rec & vbCrLf, 0
    Next
    objFile.SaveToFile outputfile, 2
    objFile.Close 'Streamを閉じる
    Set objFile = Nothing 'オブジェクトの解放

    If Err.Number > 0 Then
        WScript.Echo "File Write Error"
        WScript.Quit
    End If
On Error Goto 0


'[Q] PF と NON PFの見分け方は？
'[Q] 取得結果は1つのCSVにする？複数のCSVにする？
'[TODO] CSV用に”のエスケープ処理
'[TODO] WAで動作させてみる
