#Include "Totvs.ch"
#Include 'Protheus.ch'
#Include 'TbiConn.ch'
#INCLUDE "REPORT.CH"
#INCLUDE "TOPCONN.CH"
#Include 'tryexception.ch'

/*
 * Funcao............: AfterLogin
 * Responsável.......: Aline Porfirio
 * Data..............: 16/10/2025
 * Objetivo..........: PE para setar control+U para execução de query
 *
*/
User Function AfterLogin()

	Local aALIAS	:= GetArea()
	Local _aGrp		:= UsrRetGrp(__cUserID)
	Local _nNext	:= 0

	For _nNext := 1 to Len(_aGRP)

		If "000000" $ _aGRP[_nNext]

			SetKey(K_CTRL_U, {|| U_AftrFnc() })

		EndIf

	Next

	RestArea(aALIAS)

Return .T.

/*
 * Funcao............: AftrFnc
 * Responsável.......: Aline Porfirio
 * Data..............: 16/10/2025
 * Objetivo..........: Função para execução de Query e funçlões Genéricas.
 *
*/

User Function AftrFnc()

	Local _oBtn01, _oBtn02, _oBtn03, _oBtn04, _oBtn05, _oBtn06, _oDlg

	DEFINE MSDIALOG _oDlg TITLE "Utilitarios" FROM 000, 000  TO 150, 525 COLORS 0, 16777215 PIXEL

	@ 005, 001 BUTTON _oBtn01  PROMPT "Extrator Query" 	SIZE 052, 030 OF _oDlg ACTION Processa( {|| U_ERPQuery() }, "Processamento ERPQuery"	) PIXEL
	@ 005, 053 BUTTON _oBtn02  PROMPT "Outras Funcões" 	SIZE 052, 030 OF _oDlg ACTION Processa( {|| U_AftrFnca()}, 	"Execução de Funções"		) PIXEL
	@ 005, 104 BUTTON _oBtn03  PROMPT "Gerenc. Arquivo" SIZE 052, 030 OF _oDlg ACTION Processa( {|| U_ApExplorer()},"Explorador de Arquivos" 	) PIXEL
	@ 005, 156 BUTTON _oBtn04  PROMPT "#N/A"			SIZE 052, 030 OF _oDlg ACTION Processa( {||  }, "" ) PIXEL
	@ 005, 207 BUTTON _oBtn05  PROMPT "#N/A"			SIZE 052, 030 OF _oDlg ACTION Processa( {||  }, "" ) PIXEL

	@ 040, 001 BUTTON _oBtn06 PROMPT "SAIR" SIZE 258, 030 OF _oDlg ACTION _oDlg:End() PIXEL

	ACTIVATE MSDIALOG _oDlg CENTERED

Return

User Function ERPQuery()

	Local oERPQry
	Local cERPQry   := Space(1000)
	Local lMaster   := .F.
	Local cGruposNm := ""
	Local aGruposNm := UsrRetGrp()

	Private oERPResu
	Private cERPResu  := "Exemplo de Referência em Tabelas Internas = Select * From SM4" + Alltrim(cEmpAnt) + "0"
	Private oDlg
	Private oGetDB
	PRivate cQry   	:= getNextAlias()

	cGruposNm := aGruposNm[1]

	// Bloquear Acesso - Libera apenas para administradores.
	If "000000"$cGruposNm
		lMaster := .T.
	EndIf
	lMaster := .T.

	DEFINE MSDIALOG oDlg TITLE "Extrator de Query's" FROM 000,000 TO 560,900 PIXEL

	oERPQry := tMultiget():New(002,002,{|u|if(Pcount()>0,cERPQry:=u,cERPQry)},oDlg,448,090,,,,,,.T.,,,,,,!lMaster,,,,,.T.,"Expressão SQL",1,,CLR_GREEN)

	@ 104, 005 BUTTON "Novo"		SIZE  50,16 PIXEL OF oDlg ACTION NwQry(@cERPQry)
	@ 104, 060 BUTTON "Abrir"		SIZE  50,16 PIXEL OF oDlg ACTION OpnQry(@cERPQry)
	@ 104, 120 BUTTON "Salvar"		SIZE  50,16 PIXEL OF oDlg ACTION SvQry( cERPQry)
	@ 104, 180 BUTTON "Executar"	SIZE  50,16 PIXEL OF oDlg ACTION ExecQry(@cERPQry, @cERPResu, @lMaster )
	@ 104, 240 BUTTON "Excel"		SIZE  50,16 PIXEL OF oDlg ACTION MsAguarde({|| ExecXcel(@cERPQry,cQry)},"Aguarde","Executando query...",.F.)
	@ 104, 400 BUTTON "Sair"		SIZE  50,16 PIXEL OF oDlg ACTION oDlg:End()

	oERPResu := tMultiget():New(122,002,{|u|if(Pcount()>0,cERPResu:=u,cERPResu)},oDlg,448,27,,,,,,.T.,,,{||.F.},,,,,,,,.T.,"Log. (Empresa Conectada = " + Alltrim(cEmpAnt) + "0)",1,,CLR_RED)

	SetKey(VK_F5, {|| ExecQry(@cERPQry, @cERPResu, @lMaster )})
	SetKey(VK_F8, {|| MsAguarde({||ExecXcel(@cERPQry,cQry)},"Aguarde","Executando Query e Excel da consulta...",.F.)})
	SetKey(VK_ESCAPE, {|| oDlg:End() })

	ACTIVATE MSDIALOG oDlg CENTER

Return

Static Function ExecXcel(_cQuery, _aAlias)

	Local nCnt 	  := 0
	Local aStruQry  := {}
	Local cArquivo  := "query"+ DToS( MsDate() ) + "_" + StrTran( Time(), ":", "" )+".xlsx"
	Local oFWMsExcel
	Local oExcelApp
	Local cPath     := "C:\TEMP\"
	Local aLinExcel := {}
	Local cWorkSheet := "Query"
	Local cTitulo    := "Resultado da Query"

	If ("DELETE " $ Upper(_cQuery) .OR. "UPDATE " $ Upper(_cQuery) .OR. "INSERT INTO " $ Upper(_cQuery))
		Alert("Regras de DELETE/UPDATE/INSERT não podem ser exportadas para o excel.")
		Return
	Endif

	If !Empty(_cQuery)
		
		TryException
		MPSysOpenQuery( _cQuery , cQry)

		(cQry)->(DbGoTop())
		(cQry)->(DbEval({|| nCnt++}))
		(cQry)->(DbGoTop())

		aStruQry  := (cQry)->(dbStruct())
		CatchException using oException
		Alert("Houve um erro na execução da Query, por favor verifique!")
		Return
		EndException

	Endif

	If nCnt <= 0
		Alert("Não ha Dados na Tabela para Gerar o EXCEL.")
		Return
	Endif

	aColunas := {}
	aLocais := {}
	oBrush1 := TBrush():New(, RGB(193,205,205))

	// Verifica se o Excel está instalado na máquina

	If !ApOleClient("MSExcel")
		MsgAlert("Microsoft Excel não instalado!")
		Return
	EndIf

	// Configurando Excel
	oFWMsExcel := FwMsExcelXlsx():New()
	oFWMsExcel:AddworkSheet(cWorkSheet)
	oFWMsExcel:AddTable(cWorkSheet,cTitulo)

	// Criação de colunas
	For nCnt := 1 To Len(aStruQry)
		oFWMsExcel:AddColumn(cWorkSheet,cTitulo,aStruQry[nCnt,1],IIF(aStruQry[nCnt,2]=="N",3,1),IIF(aStruQry[nCnt,2]=="N",1,1),.F.)
	Next nCnt

	While !(cQry)->(Eof())
		// Criação de Linhas
		aLinExcel := {}
		For nCnt := 1 To Len(aStruQry)
			aAdd(aLinExcel,(cQry)->&(aStruQry[nCnt,1]))
		Next nCnt

		oFWMsExcel:AddRow(cWorkSheet,cTitulo, aLinExcel)

		(cQry)->(DBSkip(1))
	EndDo

	oFWMsExcel:Activate()
	oFWMsExcel:GetXMLFile(cArquivo)

	CpyS2T("\SYSTEM\"+cArquivo, cPath)

	oExcelApp := MsExcel():New()
	oExcelApp:WorkBooks:Open(cPath+cArquivo) // Abre a planilha
	oExcelApp:SetVisible(.T.)
	oExcelApp:Destroy()

Return

Static Function ExecQry(cERPQry, cERPResu, lMaster)

	Local aLinha    := {}
	Local nRet		:= 0
	Local aQuery	:= {}
	Local nX		:= 0
	Local lRet		:= .T.

	aLinha := StrToKarr(cERPQry,CHR(13)+CHR(10))

	If !Empty(cERPQry)

		If Select(cQry) > 0
			
			(cQry)->(DbCloseArea())

		EndIf

		If ("DELETE " $ Upper(cERPQry) .OR. "UPDATE " $ Upper(cERPQry) .OR. "INSERT INTO " $ Upper(cERPQry))

			cERPQry := StrTran(cERPQry, Chr(13), "")
			cERPQry := StrTran(cERPQry, Chr(10), "")

			If Substr(cERPQry,Len(alltrim(cERPQry)),1) == ";"

				cERPQry := Substr(cERPQry,1,Len(alltrim(cERPQry))-1)

			EndIf

			If ";" $ cERPQry 
			
				aQuery := StrTokArr2(cERPQry,";",.T.)

			Else

				aadd(aQuery,cERPQry)

			EndIf

			Begin Transaction

				For nX := 1 to len(aQuery)

					nRet := TCSqlExec(aQuery[nX])

					If nRet < 0

						Alert("Erro na linha "+alltrim(str(nX))+": "+ TCSqlError())
						lRet	:= .F.
			
					EndIf
					
				Next nX

				If lRet

					Alert("Executado com sucesso! ")

				EndIf

			END Transaction


		Else

			TryException

			MPSysOpenQuery( cERPQry , cQry)
			(cQry)->(DbGoTop())
			MsAguarde({||GeraArq(cQry)},"Aguarde","Executando Query...",.F.)
			(cQry)->(DbCloseArea())

			CatchException using oException
			Alert("Houve um erro na execução da Query, por favor verifique!")
			EndException

		EndIf

	Else

		MsgInfo("Comando Query não informado")
	
	EndIf

Return cQry

Static Function GeraArq(cQry)

	Local nRet    := 0
	Local nI      := 0
	Local cDirDocs  := "C:\TEMP\"
	Local aStruQry	:= {}
	Local aCmpQry   := {}
	Local aResQry   := {}
	Local aResult   := {}
	Local choraIni  := Time()
	Local choraFim  := Time()
	Local cRetSql   := ""
	Local nQtdSql   := 0

	If !lIsDir(cDirDocs)
		nRet := MakeDir( cDirDocs, Nil, .F. )
		if nRet != 0
			Alert( "Não foi possível criar o diretório "+cDirDocs+", crie manualmente. Erro: " + cValToChar( FError() ) )
		endif
	Endif

	DbSelectArea(cQry)

	aStruQry := (cQry)->(dbStruct())

	For nI := 1 To Len(aStruQry)
		If aStruQry[ni,2] <> "C"
			aAdd(aCmpQry,{aStruQry[ni,1], aStruQry[ni,1], "@E", aStruQry[ni,3], aStruQry[ni,4]})
		Else
			aAdd(aCmpQry,{aStruQry[ni,1], aStruQry[ni,1], "@!", aStruQry[ni,3], aStruQry[ni,4]})
		Endif
	Next

	(cQry)->(DbGoTop())
	While (cQry)->(!Eof())
		nQtdSql++
		aResQry := {}
		For nI := 1 To Len(aCmpQry)
			aAdd(aResQry,(cQry)->&(aCmpQry[nI,1]))
		Next nI
		aAdd(aResQry,.F.)
		aAdd(aResult,aResQry)
		(cQry)->(DbSkip())
	End

	oGetDB := MsNewGetDados():New(161,002,278,451,,,,,{''},,,,,,oDlg,aCmpQry,aResult)

	choraFim  := Time()
	cRetSql += "Exemplo de Referência em Tabelas Internas = Select * From SM4" + Alltrim(cEmpAnt) + "0" + Chr(13)+Chr(10)
	cRetSql += "Tempo Execução = " + ElapTime(choraIni,choraFim) + Chr(13)+Chr(10)
	cRetSql += "Qtd. Registros = " + Alltrim(Str(nQtdSql))+ Chr(13)+Chr(10)
	cERPResu := cRetSql
	oERPResu:Refresh()
	oDlg:Refresh()

Return

Static Function SvQry(cERPQry)
	Local cFile
	Local nHdl

	cFile := cGetFile('Advanced Protheus Query Analyzer |*.APQ |SQL Query Analyzer |*.SQL |All Files|*.*','Save File',,,.F.,GETF_ONLYSERVER)
	If Empty(cFile)
		Return
	EndIf
	If At(".",cFile) == 0
		cFile += ".APQ"
	EndIf

	If File(cFile)
		If ApMsgYesNo("Sobrescrever o arquivo existente?","Atenção")
			FErase(cFile)
		Else
			Return
		EndIf
	EndIf

	nHdl := FCreate(cFile,0)
	If nHdl < 0
		ApMsgAlert("Falha na gravação do arquivo "+cFile )
		Return
	EndIf
	FWrite(nHdl,cERPQry)
	FClose(nHdl)
Return

Static Function OpnQry(cERPQry)
	Local cFile
	Local cBuffer
	Local nLength
	Local nHdl

	cFile := cGetFile('Advanced Protheus Query Analyzer |*.APQ |SQL Query Analyzer |*.SQL |All Files|*.*','Open File',,,.T.,GETF_ONLYSERVER)
	If Empty(cFile)
		Return
	EndIf

	If At(".",cFile) == 0
		cFile += ".APQ"
	EndIf

	nHdl := FOpen(cFile,0)
	If nHdl < 0
		ApMsgAlert("Falha na leitura do arquivo "+cFile )
		Return
	EndIf

	nLength := FSeek(nHdl,0,2)
	cBuffer := Space(nLength)
	FSeek(nHdl,0)
	FRead(nHdl,@cBuffer,nLength)
	FClose(nHdl)
	cERPQry := cBuffer
Return

Static Function NwQry(cERPQry)

	If !Empty(cERPQry) .And. ApMsgYesNo("Salvar script corrente?","Atenção")
		ApQrySvQry()
	EndIf

	cERPQry := ""

Return

User Function AftrFnca()
	Local _aPergunta	:= {}
	MV_PAR01 := Space(100)
	aAdd(_aPergunta, {1,"Nome da Função:", MV_PAR01, "", "", "", ".T.", 100, .F.})
	If ParamBox(_aPergunta, "Executar Função ?")
		CheckExecForm( MV_PAR01,.T.)
	EndIf
Return
