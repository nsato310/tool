'------------------------------------------------------------------------
'�萔��`
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'�J�����g�f�B���N�g���̃p�X
'------------------------------------------------------------------------
Dim Obj00 : Set Obj00 = WScript.CreateObject("Scripting.FileSystemObject")
Dim var00 : var00     = Obj00.getParentFolderName(WScript.ScriptFullName)

'------------------------------------------------------------------------
'01:�t�@�C���̃p�X���擾����
'------------------------------------------------------------------------
Dim Obj01     : Set Obj01     = CreateObject("Scripting.FilesystemObject")
Dim ObjFileName : Set ObjFileName = CreateObject("Scripting.FileSystemObject")
Dim ObjInFolder : Set ObjInFolder = Obj01.getFolder(var00& "\in\")
Dim ObjInFiles  : Set ObjInFiles  = ObjInFolder.Files
Dim OutFolder : OutFolder = var00& "\out\"
Dim TempFileName

For Each TempFileName In ObjInFiles
	On Error Resume Next
	Dim fileName  : fileName = ObjFileName.getFileName(TempFileName)
	Dim filePathIn  : filePathIn = TempFileName.Path
	Dim filePathOut : filePathOut = OutFolder & fileName

	If Err.Number <> 0 Then
		MsgBox "�t�@�C���p�X�擾�G���[���������܂���" & vbCrLf & _
			Err.Description, vbExclamation
	End If
	On Error Goto 0
	'------------------------------------------------------------------------
	'02:Input�t�@�C����ǂݍ���
	'------------------------------------------------------------------------
	On Error Resume Next
	'Stream �I�u�W�F�N�g
	Dim Obj0201 : Set Obj0201 = CreateObject("ADODB.Stream")

	With Obj0201
		.Type = 2 'adTypeText
		.Charset = "shift_jis"
		.Open
		.LoadFromFile filePathIn
	End With

	Dim TempTextBefore : TempTextBefore = Obj0201.ReadText(-2) '�w�b�_�[�s�͋�ǂ�
	TempTextBefore = Obj0201.ReadText(-1) '2�s�ڈȍ~��ǂݍ���

	Obj0201.Close
	Set Obj0201 = Nothing

	If Err.Number <> 0 Then
		MsgBox "Input�t�@�C���ǂݍ��݂ŃG���[���������܂���" & vbCrLf & _
			Err.Description, vbExclamation
	End If
	On Error Goto 0
	'------------------------------------------------------------------------
	'02:���X�̒u��
	'------------------------------------------------------------------------
	On Error Resume Next
	Dim regEx : Set regEx = New RegExp

	With regEx
	   .Pattern = "^[0-9]{3}" '�s��3���̔C�ӂ̐���
	   .Global = True
	   .MultiLine = True
	End With

	Dim TempTextAfter
	TempTextAfter = regEx.Replace(TempTextBefore, "")

	If Err.Number <> 0 Then
		MsgBox "���X�̒u���ŃG���[���������܂���" & vbCrLf & _
			Err.Description, vbExclamation
	End If
	On Error Goto 0
	'------------------------------------------------------------------------
	'03:Out�t�@�C�����쐬����
	'------------------------------------------------------------------------
	On Error Resume Next
	Dim Obj0301 : Set Obj0301 = WScript.CreateObject("Scripting.FileSystemObject")
	Dim Obj0302 : Set Obj0302 = Obj0301.CreateTextFile(filePathOut,True)
	Obj0302.Close
	Set Obj0301 = Nothing
	Set Obj0302 = Nothing

	If Err.Number <> 0 Then
		MsgBox "Output�t�@�C���̍쐬�ŃG���[���������܂���" & vbCrLf & _
			Err.Description, vbExclamation
	End If
	On Error Goto 0
	'------------------------------------------------------------------------
	'04:Out�t�@�C���ɒǋL����
	'------------------------------------------------------------------------
	On Error Resume Next
	'Stream �I�u�W�F�N�g
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

	If Err.Number <> 0 Then
		MsgBox "Output�t�@�C���̏������݂ŃG���[���������܂���" & vbCrLf & _
			Err.Description, vbExclamation
	End If
	On Error Goto 0

Next

	Set Obj01     = Nothing
	Set ObjInFolder = Nothing
	Set ObjInFiles  = Nothing
	Set ObjOutFolder = Nothing
	Set ObjFileName = Nothing

MsgBox("�������܂����Bout�t�H���_���m�F���Ă�������")
