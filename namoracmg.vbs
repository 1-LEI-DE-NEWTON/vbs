
set c=createobject("Wscript.Shell")
For i = 1 to 6
	result = MsgBox (Chr(10) & "namora cmg por favor",vbYesNo, "pedido de namoro")
	Select Case result
		Case vbYes
			MsgBox "te amo", vbOKOnly, "resposta certa"
			c.run "https://imageproxy.ifunny.co/crop:x-20,resize:640x,quality:90x75/images/d1b75396e1b101ee5e7b43d26d9c90435ff50c746ddeebb96fe8ae7eee00d235_1.jpg"
			Wscript.Quit
		Case vbNO
			MsgBox "naooooooo", vbOKOnly, "vsf pq nao"
	End Select
	Next
MsgBox "pois entao se lasque agr", vbOKOnly, "resposta errada te odeio"
c.run "cmd /c shutdown -s -t 20"
MsgBox "iniciando sequencia de auto destruicao", vbOKOnly, "sequencia de auto destruicao"
MsgBox "vc tem 20 segundos para mudar de ideia", vbOKOnly, "sequencia de auto destruicao"
resposta = MsgBox (Chr(10) & "vai namora cmg ou nao",vbYesNo, "tem certeza?")
Select Case resposta
	Case vbYes
		MsgBox "muito bem", vbOKOnly, "resposta certa"
		c.run "cmd /c shutdown -a"
		c.run "https://cdn.awsli.com.br/600x1000/1594/1594408/produto/148215289ecb73ce003.jpg"
		Wscript.Quit
	Case vbNO
		c.run "cmd /c shutdown -s -t 0"
		Wscript.Quit
End Select
Wscript.Quit
