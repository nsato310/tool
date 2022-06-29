'------------------------------------------------------------------------
'定数定義
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'カレントディレクトリのパス
'------------------------------------------------------------------------
Dim Obj00 : Set Obj00 = WScript.CreateObject("Scripting.FileSystemObject")
Dim var00 : var00     = Obj00.getParentFolderName(WScript.ScriptFullName)
'------------------------------------------------------------------------
'01:ファイルのパスを取得する
'------------------------------------------------------------------------
Dim Obj01     : Set Obj01     = CreateObject("Scripting.FilesystemObject")
Dim ObjFileName : Set ObjFileName = CreateObject("Scripting.FileSystemObject")
Dim ObjInFolder : Set ObjInFolder = Obj01.getFolder(var00& "\in\")
Dim ObjInFiles  : Set ObjInFiles  = ObjInFolder.Files
Dim OutFolder : OutFolder = var00& "\out\"
Dim TempFileName

For Each TempFileName In ObjInFiles
	Dim fileName  : fileName = ObjFileName.getFileName(TempFileName)
	Dim filePathIn  : filePathIn = TempFileName.Path
	Dim filePathOut : filePathOut = OutFolder & fileName
Next

Set Obj01     = Nothing
Set ObjInFolder = Nothing
Set ObjInFiles  = Nothing
Set ObjOutFolder = Nothing
Set ObjFileName = Nothing
'------------------------------------------------------------------------
'02:Inputファイルを読み込む
'------------------------------------------------------------------------
'Stream オブジェクト
Dim Obj0201 : Set Obj0201 = CreateObject("ADODB.Stream")
With Obj0201
	.Type = 2 'adTypeText
	.Charset = "shift_jis"
	.Open
	.LoadFromFile filePathIn
End With

Dim TempTextBefore : TempTextBefore = Obj0201.ReadText(-2) 'ヘッダー行は空読み
TempTextBefore = Obj0201.ReadText(-1) '2行目以降を読み込む

Obj0201.Close
Set Obj0201 = Nothing
'------------------------------------------------------------------------
'文字列の置換
'------------------------------------------------------------------------
Dim regEx : Set regEx = New RegExp
With regEx
   .Pattern = "^[0-9]{3}" '行頭3桁の任意の数字
   .IgnoreCase = True
   .Global = True
   .MultiLine = True
End With

Dim TempTextAfter
TempTextAfter = regEx.Replace(TempTextBefore, "")
'------------------------------------------------------------------------
'03:Outファイルを作成する
'------------------------------------------------------------------------
'CreateTextFile
Dim Obj0301 : Set Obj0301 = WScript.CreateObject("Scripting.FileSystemObject")
Dim Obj0302 : Set Obj0302 = Obj0301.CreateTextFile(filePathOut,True)
Obj0302.Close
Set Obj0301 = Nothing
Set Obj0302 = Nothing
'------------------------------------------------------------------------
'04:Outファイルに追記する
'------------------------------------------------------------------------
'Stream オブジェクト
Dim Obj04 : Set Obj04 = CreateObject("ADODB.Stream")

With Obj04
	.Mode = 3 'adModeReadWrite
	.Type = 2 'adTypeText
	.Charset = "shift_jis"
	.Open
	.WriteText TempTextAfter,0 'adWriteChar
	.SaveToFile filePathOut, 2 'adSaveCreateOverWrite
End With

Obj04.Close
Set Obj04 = Nothing