dim op, executar
call criar_objeto

sub criar_objeto()
	set executar = wscript.createobject("wscript.shell")
	call menu
end sub

sub menu()
	op = cint(inputbox("[1] Voltar a Tela 01" , "ESCOLHA UMA OPCAO"))
	select case op
		case 1:
			executar.Run("menu.vbs")
		case else
			msgbox("Opção Inválida!!!"), vbExclamation + vbOKOnly, "Atenção"
		end select

end sub
