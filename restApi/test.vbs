' vbscriptでREST API呼び出し
' ServerXMLHTTPを利用してAPI呼び出し
' リクエストのJSONを作成
' レスポンスのJSONをパース
' 結果を利用してさらにWebAPI呼び出し
' 応答が返らないときなどの例外処理
' OAuth認証

' (公式)ServerXMLHTTPについて
' https://technet.microsoft.com/ja-jp/security/ms762278(v=vs.80)

Option Explicit

MsgBox("hello")

Dim oHTTP
' オブジェクト変数にオブジェクトや値を代入するにはSetを利用する
Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
' HTTPリクエストオブジェクトの初期化
'  - bstrMethod
'  - bstrUrl
'  - bAsync(optional)
'  - btrUser(optional)
'  - bstrPassword(optional)
oHTTP.Open "GET", "https://qiita.com/api/v2/tags/vbscript", False
' HTTPヘッダの指定
'  - bstrHeader
'  - bstrValue
oHTTP.SetRequestHeader "Content-Type", "application/json"
' HTTPリクエスト実行
oHTTP.Send()
MsgBox("HTTP Requested!")

' readyState：リクエストの状態
'  - 0 : UNINITIALIZED : openメソッドが呼ばれていないため未初期化の状態
'  - 1 : LOADING : openは呼ばれたがsendが呼ばれていない状態
'  - 2 : LOADED : sendが呼ばれたが、レスポンスが届いていない状態
'  - 3 : INTERACTIVE : いくつかのデータは受信したが、responseBody/responseTextが届いていない状態
'  - 4 : COMPLETED : データ受信完了した状態
MsgBox("readyState: " & oHTTP.readyState)
MsgBox("status: " & oHTTP.status)
MsgBox("response text: " & oHTTP.ResponseText)


' VBSからJSONをパースする方法
'  - CreateObject("ScriptControl")する方法
'    - VBSからJScriptを呼び出す方法らしいが、64bit環境だと使えない？
'  - CreateObject("HtmlFile")する方法
' http://bougyuusonnin.seesaa.net/article/446183415.html

' WSHではスクリプト終了時にすべてのオブジェクト変数に格納されたオブジェクトは解放されるため、Nothingは必須ではない。
' Set oHTTP = Nothing