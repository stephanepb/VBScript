dim nome, audio, a1, a2, a3, a4, nivel1, nivel2, nivel3, nivel4, acertos, ouvir, pular, palavra, jasorteadas(20), n, resp, dois, x

nivel1 = Array("rato", "pato", "gato", "cachorro", "corno")
nivel2 = Array("mickey", "donald", "pateta", "luto", "margarida")
nivel3 = Array("mendigo", "mortadela", "labirinto", "jacinto", "cérebro")
nivel4 = Array("balaústre", "cornucópia", "hebdomadário", "iconoclasta", "numismática")

call carregar_audio	
call pnome
randomize(second(time))
call jogo 

sub pnome()
	a1.controls.play
	nome = inputbox("Digite seu nome: ", "Olá!")
	if nome=false then
		wscript.quit
	elseif nome=vbNullString then
		MsgBox ("Campo Vazio. É necessário digitar um valor!"), vbInformation + vbOKOnly, "Atenção"
		call pnome
	end if
end sub 


sub carregar_audio()
	set audio=createobject("SAPI.SPVOICE")
	audio.volume=100
	audio.rate=1 
	
	set a1 = CreateObject("WMPlayer.OCX")
	a1.url = "inicio.wav"
	
end sub

sub jogo()
	acertos = 0
	pular = 1
	x = 0
	
	a1.controls.play
	msgbox("Ouça a palavra e escreva-a corretamente." + vbNewLine &_
		   "Você vence ao completar 12 acertos e perde ao errar." + vbNewLine &_
		   "Boa sorte!!!"), vbInformation + vbOKOnly, "SOLETRANDO"
	call sorteio
end sub

sub sorteio()
	if acertos<3 then
		palavra =  nivel1(int(rnd * 5))
	elseif acertos<6 then
		palavra =  nivel2(int(rnd * 5))
	elseif acertos<9 then
		palavra =  nivel3(int(rnd * 5))
	elseif acertos<12 then
		palavra =  nivel4(int(rnd * 5))
	end if
	
	for n=0 to 19 step 1
		if palavra=jasorteadas(n) then
			call sorteio
		end if
	next
	
	jasorteadas(x) = palavra
	x = x + 1
	
	audio.speak(palavra)	
	ouvir = 1	
	call rodadas	
end sub

sub rodadas()
	resp = inputbox("" & nome & ", digite a palavra ouvida: " + vbNewLine &_
					"" + vbNewLine &_
					"---------------------------------------------------------" + vbNewLine &_
					"[O]uvir a palavra novamente" + vbNewLine &_
					"[P]ular a palavra (apenas uma vez)" + vbNewLine &_
					"---------------------------------------------------------", "Sua vez")
	if resp=false then
		wscript.quit
	end if
	
	resp = LCase(resp)	
	if resp=vbNullString then
        MsgBox ("Campo Vazio. É necessário digitar um valor!"), vbInformation + vbOKOnly, "Atenção"
		call rodadas
	elseif resp="o" then
		if ouvir=1 then 
			audio.speak(palavra)
			ouvir = 0
			call rodadas
		else
			msgbox("NÃO PODE MAIS OUVIR >:("), vbInformation + vbOKOnly, "SOLETRANDO"
			call rodadas
		end if
	elseif resp="p" then
		if pular=1 then 
			pular = 0	
			call sorteio
		else 
			msgbox("NâO PODE MAIS PULAR >:("), vbInformation + vbOKOnly, "SOLETRANDO"
			call rodadas
		end if
	elseif resp=palavra then
		acertos = acertos + 1
		if acertos=12 then
			set a3 = CreateObject("WMPlayer.OCX")
			a3.url = "venceu.wav"
			a3.controls.play
						
			resp = msgbox("VOCÊ VENCEU!!!!!" + vbNewLine &_
						  "Deseja jogar novamente?", vbQuestion + vbYesNoCancel, "UHUUUUL")
			if resp=vbyes then
				call jogo
			else
				wscript.quit
			end if 
		else
			set a2 = CreateObject("WMPlayer.OCX")
			a2.url = "acertou.wav"
			a2.controls.play
			msgbox("Acertou!!!" + vbNewLine &_
				   "acertos: " & acertos & "" ), vbInformation + vbOKOnly, "UAU"
			call sorteio
		end if
	else
		set a4 = CreateObject("WMPlayer.OCX")
		a4.url = "perdeu.wav"
		a4.controls.play
		resp = msgbox("Perdeu!!!" + vbNewLine &_
					  "A palavra era: " & palavra & "" + vbNewLine &_
					  "Deseja jogar novamente?", vbQuestion + vbYesNoCancel, ":(")
		if resp=vbyes then
			call jogo
		else
			wscript.quit
		end if 
	end if
end sub