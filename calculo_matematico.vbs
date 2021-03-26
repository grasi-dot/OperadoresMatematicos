
dim n,n1,n2,operador(3),resp,resultado,cont,calculo, resp1,sorteio,sorteio1,sorteio2 
call inicio

sub inicio()
cont=0
call jogo
end sub

sub jogo()

operador(1) = "-"
operador(2) = "+"
operador(3) = "x"

for n=1 to 3 step 1
	randomize(second(time))
	sorteio=int(rnd * 3) + 1
next

for n1=1 to 10 step 1
	randomize(second(time))
	sorteio1=int(rnd * 10) + 1
next

for n2=1 to 10 step 1
	randomize(second(time))
	sorteio2=int(rnd * 10) + 1
next

if sorteio = 1 then
	resultado = sorteio1 - sorteio2
elseif sorteio = 2 then
	resultado = sorteio1 + sorteio2
else 
    resultado = sorteio1 * sorteio2
end if

resp=cint(inputbox("ACERTE O CALCULO MATEMATICO" + vbnewline & _
				   "=============================" + vbnewline &_
				   "    "& sorteio1 &" "& operador(sorteio) &" "& sorteio2 &" = ???", "RESOLVA"))
	if resp = resultado then
		cont = cont + 1
		calculo=(msgbox("PARABENS VOCE ACERTOU!" + vbnewline &_
						"QUANTIDADE DE ACERTOS: "& cont & "", vbinformation + vbOKOnly,"AVISO"))
		call jogo
	else
		msgbox("VOCE ERROU"), vbInformation + vbOK, "ATENCAO"
		resp1=msgbox("DESEJA JOGAR NOVAMENTE",vbquestion + vbyesno,"ATENCAO")
		if resp1=vbyes then
			call inicio
		else
			wscript.quit
		end if	
	end if
end sub

sub  sair()
resp1=msgbox("DESEJA JOGAR NOVAMENTE",vbquestion + vbyesno,"ATENCAO")
if resp=vbyes then
	call inicio
else
	wscript.quit
end if	
end sub
