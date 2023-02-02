Option Explicit

'↓Libre Calc VBAの場合 ********************************
'Option VBASupport 1

Const asm = "chickenrace.asm" 'ASMはUTF-8
Const wrk = "prg-wrk.txt"
Const prg = "prg.txt"

Dim label(100, 2)
Dim zeroPage(100)
Dim cDir

'↓VBSの場合 ********************************
Call Main

Sub Main()

	Dim str, edt
	Dim cnt, zero, line
	Dim i, j, p
	Dim objInput, objOutput

	'初期化
	'↓VBAの場合 ********************************
	'cDir = CurDirVBA()	'VBA、Calc VBAの場合

	'↓VBSの場合 ********************************
	cDir = CurDir()	

	cnt = 32768 'ROMは$8000(32768)から始まっているから
	line = 0

	For i = 0 To UBound(label)
		label(i, 0) = ""
		label(i, 1) = 0 'ラベルの絶対番地
	Next

	For i = 0 To UBound(zeroPage)
		zeroPage(i) = "" 'iを番地とする
	Next

	Set objOutput = CreateObject("ADODB.Stream")
	objOutput.Open
	objOutput.Type = 2	' テキストファイル
	objOutput.Charset = "shift_jis"	' 文字コード
	objOutput.LineSeparator = 10	' 区切り文字 (LF) セパレーターなしで書き込むからなんでも

	Set objInput = CreateObject("ADODB.Stream")
	objInput.Open
	objInput.Charset = "UTF-8"	'BOMあり、BOMなし両対応
	objInput.LineSeparator = 10	'LF　CrLf（-1）を指定するとなぜかエラーとなる
	objInput.LoadFromFile cDir & "\" & asm
				
	Do Until objInput.EOS

		line = line + 1
		str = objInput.ReadText(-2)	'-2 = 1行づつ読み込む

		'末尾の改行コードを除去
		If right(str,1) = vbCr Then
			str = Left(str, len(str)-1)
		End if
		
		';があったら最初の;手前まで
		p = Instr(str, ";")
		If p > 0 Then
			str = Left(str, p-1)
		End If

		'タブがあったら半角スペースに置き換え
		str = Replace(str, vbTab, " ")

		'連続する半角スペースを1つに
		Do While Instr(str, "  ") > 0
			str = Replace(str, "  ", " ")
		Loop

		'前後のスペースを除去
		str = Trim(str) 

		'小文字のオペコードにも対応（6502アセンブラでは大文字？）
		str = UCase(str)

		'空行はスルー
		If str = "" Then
			'何もしない

		'ラベル
		ElseIf Right(str, 1) = ":" Then
			edt = Left(str, Len(str)-1)
			Call addLabel(edt, cnt) 
			objOutput.WriteText str & "#" & cnt & vbCrLf, 0 '←改行コードを付けない

		'LDA #$
		ElseIf Left(str, 6) = "LDA #$" Then
			cnt = cnt + 2
			objOutput.WriteText "A9 " & Right(str, 2) & " " & vbCrLf, 0

		'LDA #%
		ElseIf Left(str, 6) = "LDA #%" Then
			cnt = cnt + 2
			edt = Right(str, 8)
			objOutput.WriteText "A9 "  & Right("00" & Bin2Hex(edt), 2) & " " & vbCrLf, 0

		'LDA $
		ElseIf Left(str, 5) = "LDA $" Then
			cnt = cnt + 3
			edt = Right(str, 4)
			objOutput.WriteText "AD "  & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'LDA #
		ElseIf Left(str, 5) = "LDA #" Then
			cnt = cnt + 2
			edt = Mid(str, 6)
			objOutput.WriteText "A9 "  & Right("00" & Hex(edt), 2) & " " & vbCrLf, 0

		'LDA
		ElseIf Left(str, 3) = "LDA" Then
			cnt = cnt + 3
			zero = getZeroPage(Mid(str, 5))
			If zero = -1 Then
				zero = addZeroPage(Mid(str, 5))
			End If
			edt = Right("0000" & Hex(zero),4)
			objOutput.WriteText "AD " & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'LDX #$
		ElseIf Left(str, 6) = "LDX #$" Then
			cnt = cnt + 2
			objOutput.WriteText "A2 " & Right(str, 2) & " " & vbCrLf, 0

		'LDX #%
		ElseIf Left(str, 6) = "LDX #%" Then
			cnt = cnt + 2
			edt = Right(str, 8)
			objOutput.WriteText "A2 "  & Right("00" & Bin2Hex(edt), 2) & " " & vbCrLf, 0

		'LDX $
		ElseIf Left(str, 5) = "LDX $" Then
			cnt = cnt + 3
			edt = Right(str, 4)
			objOutput.WriteText "AE "  & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'LDX #
		ElseIf Left(str, 5) = "LDX #" Then
			cnt = cnt + 2
			edt = Mid(str, 6)
			objOutput.WriteText "A2 "  & Right("00" & Hex(edt), 2) & " " & vbCrLf, 0

		'LDX
		ElseIf Left(str, 3) = "LDX" Then
			cnt = cnt + 3
			zero = getZeroPage(Mid(str, 5))
			If zero = -1 Then
				zero = addZeroPage(Mid(str, 5))
			End If
			edt = Right("0000" & Hex(zero),4)
			objOutput.WriteText "AE " & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'STA $
		ElseIf Left(str, 5) = "STA $" Then
			cnt = cnt + 3
			edt = Right(str, 4)
			objOutput.WriteText "8D " & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'STA
		ElseIf Left(str, 3) = "STA" Then
			cnt = cnt + 3
			zero = getZeroPage(Mid(str, 5))
			If zero = -1 Then
				zero = addZeroPage(Mid(str, 5))
			End If
			edt = Right("0000" & Hex(zero),4)
			objOutput.WriteText "8D " & Right(edt, 2) & " " & Left(edt, 2) & " " & vbCrLf, 0

		'AND #
		ElseIf Left(str, 5) = "AND #" Then
			cnt = cnt + 2
			edt = Mid(str, 6)
			objOutput.WriteText "29 " & Right("00" & edt, 2) & " " & vbCrLf, 0

		'ADC #
		ElseIf Left(str, 5) = "ADC #" Then
			cnt = cnt + 2
			edt = Mid(str, 6)
			objOutput.WriteText "69 " & Right("00" & edt, 2) & " " & vbCrLf, 0

		'CLC
		ElseIf Left(str, 3) = "CLC" Then
			cnt = cnt + 1
			objOutput.WriteText "18 " & vbCrLf, 0

		'BPL
		ElseIf Left(str, 3) = "BPL" Then
			cnt = cnt + 2
			objOutput.WriteText str & "#" & cnt & vbCrLf, 0

		'BNE
		ElseIf Left(str, 3) = "BNE" Then
			cnt = cnt + 2
			objOutput.WriteText str & "#" & cnt & vbCrLf, 0

		'JMP
		ElseIf Left(str, 3) = "JMP" Then
			cnt = cnt + 3
			objOutput.WriteText str & vbCrLf, 0

		'What!?
		Else
			MsgBox("#" & line & " - unsupported opcode!:" & str)

			'↓VBAの場合 ********************************
			'End

			'↓VBSの場合 ********************************
			WScript.Quit
		End If
	Loop

	objInput.Close
	Set objInput = Nothing

	'cntは$8000（32768）からカウントしていたのでその分引く
	'それを32KB(1024x32 = 32768)から引く
	'更に"00 00 00 80 00 00" の6バイトを引いた数　FF　で埋める
	For i = 1 To 32768 - (cnt - 32768) - 6
		objOutput.WriteText "FF ", 0
	Next

	objOutput.WriteText "00 00 00 80 00 00", 0

	objOutput.SaveToFile cDir & "\" & wrk, 2
	objOutput.Close
	Set objOutput = Nothing

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	Set objOutput = CreateObject("ADODB.Stream")
	objOutput.Open
	objOutput.Type = 2              ' テキストファイル
	objOutput.Charset = "shift_jis"     ' 文字コード
	objOutput.LineSeparator = 10     ' 区切り文字 (LF)

	Set objInput = CreateObject("ADODB.Stream")
	objInput.Open
	objInput.Charset = "shift_jis"'BOMあり、BOMなし両対応
	objInput.LineSeparator = 10' LF   CRLF の-1を指定するとなぜだかエラーとなるため
	objInput.LoadFromFile cDir & "\" & wrk

	Do Until objInput.EOS

		str = objInput.ReadText(-2)

		'末尾の改行コードを除去
		If Right(str, 1) = vbCr Then
			str = Left(str, len(str)-1)
		End if
		
		If Instr(str, ":") > 0 Then
			'ラベルならスルー
		
		ElseIf Left(str, 4) = "JMP " Then
			edt = Right("0000" & Hex(getLabel(Mid(str, 5))),4)
			objOutput.WriteText "4C " & Right(edt, 2) & " " & Left(edt, 2) & " ", 0
		
		ElseIf Left(str, 4) = "BPL " Then
			p =  Instr(str, "#")
			edt = Mid(str, 5, p - 1 - 4) ' "BPL " の4文字分引く
			objOutput.WriteText "10 " & signedByteHex(getLabel(edt) - Mid(str, p + 1)) & " ", 0
			
		ElseIf Left(str, 4) = "BNE " Then
			p =  Instr(str, "#")
			edt = Mid(str, 5, p - 1 - 4) ' "BNE " の4文字分引く
			objOutput.WriteText "D0 " & signedByteHex(getLabel(edt) - Mid(str, p + 1)) & " ", 0

		Else
			objOutput.WriteText str, 0

		End If
		
	Loop

	objInput.Close
	Set objInput = Nothing 

	objOutput.SaveToFile cDir & "\" & prg, 2
	objOutput.Close
	Set objOutput = Nothing

	MsgBox("The End")

End Sub

'##################################################################################################

Function signedByteHex(num)
	'7bit マイナスあり

	Dim bin
	Dim syo, amari, sum
	Dim i
	
	If num >= 0 Then
		signedByteHex = Right("00" & Hex(num), 2)

	Else
		'numの補数を求める！
		num = Abs(num)
		
		bin = Right("00000000" & Dec2Bin(num), 8)
		bin = BinHanten(bin)
		sum = Bin2Dec(bin)

		'+1する
		sum = sum + 1

		'16進数にする
		signedByteHex = Right("00" & Hex(sum), 2)

	End If

End Function

Function getLabel(n)

	Dim i		
	For i = 0 To Ubound(label)
		If label(i, 0) = n Then
			getLabel = label(i, 1)
			Exit Function
		End If
	Next
	getLabel = -1
	
End Function

Function addLabel(s, adr)

	Dim i
	For i = 0 To Ubound(label)
		If label(i, 0) = "" Then
			label(i, 0) = s
			label(i, 1) = adr
			Exit Function
		End If
	Next
	
End Function

Function getZeroPage(s)

	Dim i
	For i = 0 To Ubound(zeroPage)
		If zeroPage(i) = s Then
			getZeroPage = i
			exit function
		End If
	Next
	getZeroPage = -1
	
End Function

Function addZeroPage(s)

	Dim i
	For i = 0 To Ubound(zeroPage)
		If zeroPage(i) = "" Then
			zeroPage(i) = s
			addZeroPage = i
			exit function
		End if
	Next

	MsgBox("too many variables!")

	'↓VBAの場合 ********************************
	'End

	'↓VBSの場合 ********************************
	WScript.Quit
	
End Function

Function Bin2Hex(bin)

	Dim sum
	Dim i
	Dim lenBin

	lenBin = Len(bin)
	sum = 0
	For i = 1 To lenBin
		sum = sum + Mid(bin, (lenBin + 1) - i, 1) * (2 ^ (i - 1))
	Next
	Bin2Hex = Hex(sum)
	
End Function


Function Dec2Bin(num)

	Dim bin
	Dim syo, amari, sum

	Do While num > 1
		'処理
		syo = num \ 2
		amari = num Mod 2
		num = syo
		bin = bin & CStr(amari)
	Loop
	
	'最後に商を足す
	Dec2Bin = bin & CStr(syo)
	
End Function

Function Bin2Dec(bin)
	
	Dim i, sum
	Dim lenBin

	lenBin = Len(bin)
	sum = 0
	For i = 1 To lenBin
		sum = sum + Mid(bin, (lenBin + 1) - i, 1) * (2 ^ (i - 1))
	Next

	Bin2Dec = sum

End Function

Function BinHanten(bin)

	Dim edt
	Dim i
	
	edt = ""
	For i = 1 To Len(bin)
		edt = edt & Abs(Mid(bin, i, 1) - 1)
	Next
	BinHanten = edt

End Function

Function CurDir()

	Dim objShell

	Set objShell = CreateObject("WScript.Shell")
	CurDir = objShell.CurrentDirectory

	Set objShell = Nothing

End Function

Function CurDirVBA()

	CurDirVBA = ThisWorkbook.Path

End Function
