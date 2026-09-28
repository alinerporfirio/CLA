#Include 'Protheus.ch'
#INCLUDE 'topconn.ch'
#Include 'FWMVCDEF.ch'
#Include 'RestFul.CH'
#INCLUDE "TOTVS.CH"
#INCLUDE "FWMBROWSE.CH"
#INCLUDE "PARMTYPE.CH"
#INCLUDE "TBICONN.CH"
#Include "json.ch"
#INCLUDE "APWEBSRV.CH"
#INCLUDE "XMLXFUN.CH"
#INCLUDE "HEADERGD.CH"
#INCLUDE "FILEIO.CH"
#Include "MSOle.ch"

/*
	IMPORTANTE - LEIA ANTES DE SUBIR PARA PRODUCAO
	=================================================
	A TOTVS passou a bloquear "Prepare Environment"/"RpcSetEnv" quando chamados
	de dentro de um WSMETHOD (erro "Chamada indevida de abertura de ambiente").

	Solucao adotada: o WSMETHOD (rodando no contexto REST) NAO abre mais
	ambiente nenhum. Ele so identifica a empresa/filial pelo CNPJ (lendo a
	SM0, que independe de ambiente aberto) e dispara um JOB novo via
	StartJob(). Esse job roda numa thread independente do contexto REST,
	e por isso PODE chamar Prepare Environment normalmente - e e ele quem
	efetivamente grava o funcionario/dependente (InsertEmployee/InsertDependent),
	dentro da mesma transacao de antes.

	Pontos de atencao:
	- StartJob() nao compartilha memoria com a thread do WSMETHOD, entao
	  variaveis Public/Private setadas la (ex: CMODULO/NMODULO) precisam
	  ser setadas de novo dentro da funcao do job.
	- Parametros por referencia nao "voltam sozinhos" atraves do StartJob;
	  por isso a funcao do job devolve tudo num unico array de retorno
	  ({lRet, cErroMSE, cMatricula}), que o WSMETHOD desempacota.
	- StartJob nao aceita objeto (JsonObject) como parametro - por isso
	  aFuncionario/aDependente sao extraidos para arrays simples (via
	  ClassDataArr) ANTES de chamar o job, exatamente como o codigo
	  original ja fazia.
	- A funcao do job precisa ser "User Function" (visibilidade global no
	  RPO) para o StartJob conseguir localiza-la pelo nome.
*/

WSRESTFUL RHREST DESCRIPTION "Retorna os gets necessarios para o cadastro de funcionarios"

WSDATA funcionarios AS STRING

WSMETHOD POST funcionarios;
DESCRIPTION "Efetua o cadastro dos funcionarios no Protheus";
WSSYNTAX "funcionarios";
PATH "funcionarios"   PRODUCES APPLICATION_JSON

END WSRESTFUL

WSMETHOD POST funcionarios WSSERVICE RHREST


Local lRet       	:= .T.
Local aArea      	:= GetArea()
Local cJSON      	:= Self:GetContent()
Local cEmp       	:= ''
Local cFil       	:= ''
Local jResponse  	:= JsonObject():New()
Local cErroMSE	 	:= ''
Local aFuncionario 	:= {}
Local aDependente  	:= {}
Local cMatricula	:= ''
Local nX			:= 0
Local aFieldSM0 	:= {"M0_CODIGO"}    //Posicao [1]
Local aSM0Data  	:= {}
Local aRetJob		:= {}
Local cCnpj			:= ''

// ---- DIAGNOSTICO TEMPORARIO ----
// Isso NAO deve ir pra produção. Serve so pra descobrir por que
// o StartJob nao esta achando a funcao do job: testa primeiro
// uma funcao trivial (RHREST_PING) e loga o valor exato de
// GetEnvServer(). Depois de rodar uma chamada de teste, veja
// o log e me mande o que apareceu nessas duas linhas.
Local cEnvServ  := GetEnvServer()
Local xPing     := NIL

Private nPosRA_MAT  := 0
Private oParseJSON 	:= NIL //Chamado 674

DEFAULT hTenantId   := NIL

FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, 'INICIO REST FUNCIONARIOS: ' + TIME()/*nMensure*/, /*nElapseTime*/, /*aMessage*/)
FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, 'tenantid: ' + CValToChar(hTenantId := Self:GetHeader("tenantid"))/*nMensure*/, /*nElapseTime*/, /*aMessage*/)

::SetContentType("application/json")

IF EMPTY(cJSON)

	CONOUT("Requisicao invalida/sem conteudo json (body)")

	RETURN(.F.)

ENDIF

If FWJsonDeserialize(cJSON, @oParseJSON)

	cCnpj	:= oParseJSON:cnpj

	// So identifica a empresa/filial pelo CNPJ (le a SM0). NAO abre
	// ambiente aqui - isso agora e proibido dentro do WSMETHOD.
	IdentEmpFilCnpj(cCNPJ,@cEmp,@cFil)

	aSM0Data  := FWSM0Util():GetSM0Data(cEmp, cFil, aFieldSM0)

	If Len(aSM0Data) > 0

		FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, "cEmp: "  + cEmp /*nMensure*/, /*nElapseTime*/, /*aMessage*/)
		FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, "cFil: "  + cFil /*nMensure*/, /*nElapseTime*/, /*aMessage*/)

		VARinfo("5-) oParseJSON: ", oParseJSON, NIL, .F.) // CONOUT(VARINFO("5-) oParseJSON: ",oParseJSON))

		If Empty(Alltrim(cEmp))

			cErroMSE := "Empresa nao informada"
			lRet     := .F.

		EndIf

		If Empty(Alltrim(cFil))

			cErroMSE := "Filial nao informada"
			lRet     := .F.

		EndIf

		If Empty(Alltrim(cJSON))

			cErroMSE := "Os Dados nao foram informados"
			lRet     := .F.

		EndIf

		If lRet

			aFuncionario := ClassDataArr(oParseJSON:funcionarios[1])

			If type("oParseJSON:dependent") <> 'U'

				For nX := 1 to Len(oParseJSON:dependent)

					aadd(aDependente,ClassDataArr(oParseJSON:dependent[nX]))

				Next nX

			EndIf

			FWLogMsg("INFO", , "INFO", FunName(), , , "[DIAG] GetEnvServer(): [" + cEnvServ + "] ValType: " + ValType(cEnvServ), , )

			xPing := StartJob("U_RHREST_PING", cEnvServ, .T.)

			FWLogMsg("INFO", , "INFO", FunName(), , , "[DIAG] Retorno StartJob RHREST_PING -> ValType: " + ValType(xPing) + " Valor: " + CValToChar(xPing), , )
			// ---- FIM DIAGNOSTICO TEMPORARIO ----

			// Dispara um JOB novo (thread independente do contexto REST).
			// Dentro dele o Prepare Environment volta a ser permitido, e e
			// la que InsertEmployee/InsertDependent realmente gravam os
			// dados, dentro da mesma transacao de antes.
			// lWait = .T. -> aguarda o job terminar e devolve o retorno dele.
			aRetJob := StartJob("U_RHREST_JOBADMISSAO", cEnvServ, .T., cEmp, cFil, aFuncionario, aDependente)

			If ValType(aRetJob) == 'A' .And. Len(aRetJob) >= 3

				lRet       := aRetJob[1]
				cErroMSE   := aRetJob[2]
				cMatricula := aRetJob[3]

			Else

				lRet     := .F.
				cErroMSE := "Falha ao executar o job de admissao (StartJob nao retornou o resultado esperado)"

				FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, cErroMSE /*nMensure*/, /*nElapseTime*/, /*aMessage*/)

			EndIf

		EndIf

	else

		cErroMSE := "Empresa/Filial informadas nao sao validas"
		lRet     := .F.

	EndIf

	If lRet
		jResponse["codigo"]    := 200
		jResponse["mensagem"]  := FWhttpEncode("Operacao realizada com sucesso")
		jResponse["matricula"] := cMatricula
	Else
		jResponse["codigo"]    := 400
		jResponse["mensagem"]  := FWhttpEncode("Falha na inclusao de colaborador e/ou do seu dependente")
		jResponse["detailMessage"] := FWhttpEncode(cErroMSE)
	EndIf
	::SetResponse(jResponse)

else

	SetRestFault(500,'Parser Json Error')
    lRet    := .F.

Endif

RestArea(aArea)

FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, 'FINAL REST FUNCIONARIOS: ' + TIME() /*nMensure*/, /*nElapseTime*/, /*aMessage*/)

Return(lRet)

/*/{Protheus.doc} RHREST_PING
Funcao de diagnostico TEMPORARIA - nao subir pra producao. So serve pra
confirmar se o StartJob consegue disparar QUALQUER funcao a partir do
contexto REST. Se ela tambem der "Invalid function call", o problema nao
e especifico da RHREST_JOBADMISSAO - e o StartJob (ou o destino pro qual
ele esta indo) que nao esta funcionando a partir do WSMETHOD.
@since 18/09/2026
/*/

User Function RHREST_PING()

	ConOut("[RHREST_PING] Executado com sucesso via StartJob. " + DToC(Date()) + " - " + Time())

Return .T.

/*/{Protheus.doc} RHREST_JOBADMISSAO
Job disparado via StartJob() a partir do WSMETHOD POST funcionarios.
Roda numa thread independente do contexto REST, por isso pode abrir
ambiente (Prepare Environment) normalmente. Efetua a gravacao do
funcionario e dos dependentes, dentro de uma unica transacao, e devolve
o resultado empacotado num array: {lRet, cErroMSE, cMatricula}.
@since 18/09/2026
/*/

User Function RHREST_JOBADMISSAO(cEmp, cFil, aFuncionario, aDependente)

Local lRet       	:= .T.
Local cErroMSE    	:= ''
Local cMatricula  	:= ''
Local nX          	:= 0

Private nPosRA_MAT  := 0

Prepare Environment EMPRESA cEmp FILIAL cFil MODULO "GPE"

Begin Transaction

	lRet := InsertEmployee(aFuncionario,cEmp,cFil,@cErroMSE,@cMatricula)

	If lRet

		For nX := 1 to Len(aDependente)

			If lRet

				lRet := InsertDependent(aDependente[nX],cEmp,cFil,@cErroMSE,cMatricula,nX)

			EndIf

		Next nX

		If !lRet

			DisarmTransaction()

		EndIf

	Else

		DisarmTransaction()

	EndIf

End Transaction

Return {lRet, cErroMSE, cMatricula}

Static Function InsertEmployee(aFuncionario,cEmp,cFil,cErroMSE, cMatricula)

Local lRet 			:= .T.
Local aCabec 		:= {}
Local aCabecAj 		:= {}
Local nX			:= 0
Local cContrMat  	:= ''
Local cAliqry		:= GetNextAlias()
Local cCpf			:= ''
Local cLog			:= ''
Local cConteud 		:= ''
Local cArquivo      := "\spool\Admissao_"+DTos(Date())+"_"+StrTran( Time(), ':', '' )+".log"

PRIVATE lMsErroAuto := .F.

cContrMat  	:= SuperGetMv("MV_MATRICU",NIL,"0")

If cContrMat <> "0"

	If cContrMat == "1"
		cMatricula := GetSx8Num("SRA", "RA_MAT")
	ElseIf cContrMat == "2"
		cMatricula := GetSx8Num("SRA", "RA_MAT", FwCodEmp("SRA") + "\SRA\RA_MAT")
	ElseIf  cContrMat == "3"
		cMatricula := GetSx8Num("SRA", "RA_MAT", FWGrpCompany() + "\GRPEMP\SRA\RA_MAT")
	EndIf

	aadd(aCabec,{'RA_MAT', cMatricula, NIL})

EndIf

aadd(aCabec,{'RA_SEQTURN', '01', NIL})
aadd(aCabec,{'RA_CODRET', '0561', NIL})

For nX := 1 to Len(aFuncionario)

	If aFuncionario[nX,1] <> 'EMPRESA'

		If !Empty(FWSX3Util():GetFieldType(aFuncionario[nX,1]))

			If aFuncionario[nX,1] == 'RA_ADMISSA' .and. valtype(aFuncionario[nX,2]) == 'C'

				aadd(aCabec,{aFuncionario[nX,1], ctod(aFuncionario[nX,2]), NIL})
				aadd(aCabec,{'RA_OPCAO', ctod(aFuncionario[nX,2]), NIL})
				aadd(aCabec,{'RA_VCTOEXP', ctod(aFuncionario[nX,2])+44, NIL})
				aadd(aCabec,{'RA_VCTEXP2', ctod(aFuncionario[nX,2])+89, NIL})

			ElseIf aFuncionario[nX,1] == 'RA_HRSMES' .and. valtype(aFuncionario[nX,2]) == 'N'

				aadd(aCabec,{aFuncionario[nX,1], aFuncionario[nX,2], NIL})
				aadd(aCabec,{'RA_HRSDIA', aFuncionario[nX,2]/30, NIL})
				aadd(aCabec,{'RA_HRSEMAN', aFuncionario[nX,2]/5, NIL})

			ElseIf GetSx3Cache(aFuncionario[nX,1], "X3_TIPO") == 'D' .and. valtype(aFuncionario[nX,2]) == 'C'

				aadd(aCabec,{aFuncionario[nX,1], ctod(aFuncionario[nX,2]), NIL})

			ElseIf (GetSx3Cache(aFuncionario[nX,1], "X3_TIPO") == 'C' .AND. Valtype(aFuncionario[nX,2]) == 'C') .OR. (GetSx3Cache(aFuncionario[nX,1], "X3_TIPO") == 'N' .AND. Valtype(aFuncionario[nX,2]) == 'N')

				aadd(aCabec,{aFuncionario[nX,1], aFuncionario[nX,2], NIL})

			Else

				cErroMSE += "O campo "+aFuncionario[nX,1]+" esta com conteudo invalido "+char(13)
				lRet :=	.F.

			EndIf

			If aFuncionario[nX,1] == 'RA_CIC'

				cCpf := aFuncionario[nX,2]

			EndIf

		else

			cErroMSE += "O campo "+aFuncionario[nX,1]+" nao existe no sistema "+char(13)
			lRet :=	.F.

		EndIf

	EndIf

Next nX

If lRet

	If !Empty(cCpf)

		BeginSql alias cAliQry

			SELECT RA_CIC FROM %table:SRA%
			WHERE RA_CIC = %exp:cCpf% AND
			RA_DEMISSA = '' AND
			%NotDel%
			ORDER BY 1

		EndSql

		While (cAliQry)->(!Eof())

			cErroMSE += "O CPF ja esta cadastrado!"+char(13)
			lRet :=	.F.

			(cAliQry)->(dbSkip())

		EndDo


	Else

		cErroMSE += "O CPF do colaborador nao foi informado!"+char(13)
		lRet :=	.F.

	EndIf

EndIf

IF lRet

	For nX := 1 to len(aCabec)

		If Valtype(aCabec[nX,2]) == 'D'

			cConteud := dtoc(aCabec[nX,2])
		
		Elseif Valtype(aCabec[nX,2]) == 'N'

			cConteud := alltrim(STR(aCabec[nX,2]))
		
		Else

			cConteud := aCabec[nX,2]

		EndIf

		cLog += aCabec[nX,1]+"["+Valtype(aCabec[nX,2])+"]: "+cConteud+ CRLF

	Next nX

    MemoWrite(cArquivo, cLog) 

	aCabecAj := FWVetByDic(aCabec, "SRA")

	MSExecAuto({|x,y,k,w| GPEA010(x,y,k,w)},NIL,NIL,aCabecAj,3)

	If lMsErroAuto

		CancelSX8()

		cErroMSE := MostraErro('\log\','log.txt')

		FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, "Entrei no if lMsErroAuto"/*nMensure*/, /*nElapseTime*/, /*aMessage*/)

		FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, cErroMSE/*nMensure*/, /*nElapseTime*/, /*aMessage*/)

		lRet :=	.F.

	ELSE

		FWLogMsg("INFO", /*cTransactionId*/, "INFO",  FunName(), /*cStep*/, /*cMsgId*/, "Entrei no else lMsErroAuto"/*nMensure*/, /*nElapseTime*/, /*aMessage*/)
		ConfirmSX8()

	EndIf

EndIf

Return lRet

Static Function InsertDependent(aDependente,cEmp,cFil,cErroMSE, cMatricula, nCodDep)

Local lRet 			:= .T.
Local aCabec 		:= {}
Local aItens 		:= {}
Local nX			:= 0

PRIVATE lMsErroAuto := .F.

Prepare Environment Empresa cEmp Filial cFil

aadd(aCabec,{'RA_FILIAL', cFil, NIL})
aadd(aCabec,{'RA_MAT', cMatricula, NIL})

aadd(aItens,{'RB_FILIAL', cFil, NIL})
aadd(aItens,{'RB_MAT', cMatricula, NIL})
aadd(aItens,{'RB_COD', strzero(nCodDep,2), NIL})
aadd(aItens,{'RB_PLSAUDE', '2', NIL})
aadd(aItens,{'RB_INCT', '1', NIL})

//DbSelectArea("SX3")
//SX3->(DbSetOrder(2))

For nX := 1 to Len(aDependente)

	If !Empty(FWSX3Util():GetFieldType(aDependente[nX,1]))

		If GetSx3Cache(aDependente[nX,1],"X3_TIPO") == 'D' .and. Valtype(aDependente[nX,2]) == 'C'

			aadd(aItens,{aDependente[nX,1], ctod(aDependente[nX,2]), NIL})

		Elseif (GetSx3Cache(aDependente[nX,1],"X3_TIPO") == 'C' .AND. Valtype(aDependente[nX,2]) == 'C') .OR. (GetSx3Cache(aDependente[nX,1],"X3_TIPO") == 'N' .AND. Valtype(aDependente[nX,2]) == 'N')

			aadd(aItens,{aDependente[nX,1], aDependente[nX,2], NIL})

		Else

			cErroMSE += "O campo "+aDependente[nX,1]+" esta com conteudo invalido "+char(13)
			lRet :=	.F.

		ENDIF

	Else

		cErroMSE += "O campo "+aDependente[nX,1]+" nao existe no sistema "+char(13)
		lRet :=	.F.

	EndIf

Next nX

If lRet

	Reclock('SRB',.T.)

		For nX := 1 to Len (aItens)

			&(aItens[nX,1]) := aItens[nX,2]

		Next nX

	SRB->(MsUnlock())

/*
	If SRA->(DbSeek(cFil + cMatricula))

		MSExecAuto({|x,y,w,z| GPEA020(x,y,w,z)},3,aCabec,aItens,3)

		If lMsErroAuto

			cErroMSE := MostraErro('\log\','log.txt')

			CONOUT("Entrei no if lMsErroAuto")

			Conout(cErroMSE)

			lRet :=	.F.

		ELSE

			CONOUT("Entrei no else lMsErroAuto")

		EndIf

	EndIf
*/
EndIf

Return lRet

/*/{Protheus.doc} CancelSX8
Rollback SX8
@since 02/07/2019
@version 12.1.17
/*/

Static Function CancelSX8()

Local nLenSX8 := GetSX8Len()
Local nI := 0

FOR nI:=1 TO nLenSX8
	RollBackSX8()
NEXT i

Return .T.

/*/{Protheus.doc} IdentEmpFilCnpj
So identifica a empresa/filial correspondente ao CNPJ (lendo a SM0).
NAO abre ambiente - isso e proibido dentro do WSMETHOD (contexto REST).
A abertura de fato acontece dentro do job (RHREST_JOBADMISSAO).
@since 02/07/2019
@version 12.1.17
/*/

Static Function IdentEmpFilCnpj(cCNPJ,cEmp,cFil)

    Local aSM0      := FWLoadSM0()
    Local nPos      := 0

	nPos := aScan(aSM0, {|x| AllTrim(x[SM0_CGC]) == cCnpj })

	If nPos > 0

		cEmp  := aSM0[nPos][1]
		cFil  := aSM0[nPos][2]

		ConOut("[IdentEmpFilCnpj] Empresa: " + cEmp + " Filial: " + cFil + ". " + DToC(Date()) + " - " + Time())

	Else

		ConOut("[IdentEmpFilCnpj] CNPJ: " + cCnpj + " - Nao localizado. " + DToC(Date()) + " - " + Time())

	EndIf

Return
