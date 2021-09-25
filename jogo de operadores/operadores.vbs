dim op, n1, n2, r, y, x, acertos, a1,a2,a3
op = Array("+", "-", "*")

call calculo

sub calculo()
	acertos = 0
	
	do while 1=1
		randomize(second(time))
		n1 = int(rnd * 9) + 1
		n2 = int(rnd * 9) + 1
		x = int(rnd * 3)
		
		set a3 = CreateObject("WMPlayer.OCX")
		a3.url = "inicio.wav"
		a3.controls.play
		y = cint(inputbox("=================================" + vbNewLine &_
						  "  ACERTE O CÁLCULO MATEMÁTICO" + vbNewLine &_
						  "=================================" + vbNewLine &_
						  "       RESOLVA: "& n1 & " " & op(x) & " " & n2 & " = ???"))
		select case x
			case 0:
				r = cint(n1 + n2)
			case 1:
				r = cint(n1 - n2)
			case 2:
				r = cint(n1 * n2)
		end select
		
		if y=r then
			acertos = acertos + 1
			set a1 = CreateObject("WMPlayer.OCX")
			a1.url = "acertou.wav"
			a1.controls.play
			msgbox("VOCÊ ACERTOU!" + vbNewLine &_
				   "Qtde de acertos: " & acertos & ""), vbExclamation +  vbOKOnly, "UAU"
		else
			set a2 = CreateObject("WMPlayer.OCX")
			a2.url = "perdeu.wav"
			a2.controls.play
			msgbox("ERROOOOOU!" + vbNewLine &_
				   "Qtde de acertos: " & acertos & ""), vbExclamation +  vbOKOnly, "A"
			a3.controls.play
			r = msgbox("Deseja jogar novamente?", vbQuestion + vbYesNoCancel, "É o fim?")
			if r = vbyes then
				call calculo
			else
				wscript.quit
			end if
		end if
	loop
end sub