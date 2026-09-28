#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#Include "Report.ch"

/*/{Protheus.doc} RGPER01()
Função para gerar relatório em excel conforme layout pré-definido
@type function
@version V 1.00
@author Maicon Brito
@since 29/09/2025
@param 
/*/
User Function RGPER01()

Local nReg			:= 0

Private cPerg 		:= "RGPER01"
Private oReport
Private oBreak
Private oSecEMP 	:= NIL
Private oSecFUN 	:= NIL

Private cSituacao   := ""          //  Situacoes
Private cCategoria  := ""          //  Categorias

Private cSitQuery	:= ""
Private cCatQuery	:= ""

Private nQtdCol	:=  10
	

AjustaSX1(cPerg)
Pergunte(cPerg, .T.)

cSituacao  	:= mv_par08
cCategoria 	:= mv_par09

//Modifica variaveis para a Query
	cSitQuery := ""

	For nReg:=1 to Len(cSituacao)
		cSitQuery += Subs(cSituacao,nReg,1)
		If ( nReg+1 ) <= Len(cSituacao)
			cSitQuery += "/"
		EndIf
	Next nReg

	cCatQuery := ""
	For nReg:=1 to Len(cCategoria)
		cCatQuery += Subs(cCategoria,nReg,1)
		If ( nReg+1 ) <= Len(cCategoria)
			cCatQuery += "/"
		EndIf
	Next nReg

	

oReport := ReportDef()
oReport:PrintDialog()


Return


/*/{Protheus.doc} ReportDef()
Definicao do Relatorio
@type function
@version V 1.00
@author Maicon Brito
@since 29/09/2025
@param 
/*/
Static Function ReportDef()

Local cTitulo	:= "Relatório BENEO"
Local nLoop		:= 0

oReport:= TReport():New("RGPER01",cTitulo,cPerg, {|oReport| ReportPrint(oReport,cTitulo)},"Este relatório irá imprimir a relacao de funcionario")
oReport:SetLandscape() //Define a orientacao default de pagina do relatorio como retrato. SetPortrait()/SetLandscape()
//oReport:HideParamPage()   // Desabilita a impressao da pagina de parametros.
oReport:PrintHeader(.T.,.T.)
oReport:nFontBody := 10 // Define o tamanho da fonte.

	
DEFINE SECTION oSecEMP OF oReport TABLES "QRY" TITLE "Empresa"
	
	oSecEMP:SetLineStyle() //Define a impressao da secao em linha

	DEFINE CELL NAME "CNPJ"			OF oSecEMP ALIAS "QRY" SIZE 20 TITLE "CNPJ"					PICTURE "@R 99.999.999/9999-99"
	DEFINE CELL NAME "NOME"			OF oSecEMP ALIAS "QRY" SIZE 60 TITLE "EMPRESA"				//PICTURE "@!"
	DEFINE CELL NAME "CODIGO"		OF oSecEMP ALIAS "QRY" SIZE 60 TITLE "CODIGO"				//PICTURE "@!"
	DEFINE CELL NAME "ENDERECO"		OF oSecEMP ALIAS "QRY" SIZE 60 TITLE "LOGRADOURO"			//PICTURE "@!"
	DEFINE CELL NAME "NUMERO"		OF oSecEMP ALIAS "QRY" SIZE 15 TITLE "NUMERO"				//PICTURE "@!"
	DEFINE CELL NAME "COMPLEM"		OF oSecEMP ALIAS "QRY" SIZE 40 TITLE "COMPLEMENTO"			//PICTURE "@!"
	DEFINE CELL NAME "CEP"			OF oSecEMP ALIAS "QRY" SIZE 15 TITLE "CEP"					//PICTURE "@!"
	DEFINE CELL NAME "BAIRRO"		OF oSecEMP ALIAS "QRY" SIZE 30 TITLE "BAIRRO"				//PICTURE "@!"
	DEFINE CELL NAME "CIDADE"		OF oSecEMP ALIAS "QRY" SIZE 30 TITLE "CIDADE"				//PICTURE "@!"
	DEFINE CELL NAME "UF"			OF oSecEMP ALIAS "QRY" SIZE 10 TITLE "ESTADO"				//PICTURE "@!"
	

DEFINE SECTION oSecFUN OF oReport TABLES "QRY" TITLE "Usuarios"


	DEFINE CELL NAME "CNPJ"			OF oSecFUN ALIAS "QRY" SIZE 30 TITLE "CNPJ"		 			PICTURE "@R 99.999.999/9999-99"
	DEFINE CELL NAME "MATRICULA"	OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "MATRICULA"		 	//PICTURE "@!"
	DEFINE CELL NAME "NOME"			OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "NOME COMPLETO"		//PICTURE "@!"
	DEFINE CELL NAME "EMAIL"		OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "EMAIL"		 		//PICTURE "@!"
	DEFINE CELL NAME "CELULAR"		OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "CELULAR"				// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "ATIVO"		OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "ATIVO"				// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "CODIGO"		OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "CODIGO"			 	//PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "CARGO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CARGO"		 		//PICTURE "@!"
	DEFINE CELL NAME "DEPTO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DEPARTAMENTO"		 	//PICTURE "@!"
	DEFINE CELL NAME "DIASTRAB"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DIAS_TRABALHADOS"		//PICTURE "@!"
	DEFINE CELL NAME "CPF"			OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CPF"		 			//PICTURE "@!"
	DEFINE CELL NAME "RG"			OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "RG"		 			//PICTURE "@!"
	DEFINE CELL NAME "DG"			OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DG"		 			//PICTURE "@!"
	DEFINE CELL NAME "EST_RG"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "EST.RG"		 		//PICTURE "@!"
	DEFINE CELL NAME "DTNASC"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DATA NASCIMENTO"	 	//PICTURE "@!"
	DEFINE CELL NAME "NOMEMAE"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "NOME DA MAE"		 	//PICTURE "@!"
	DEFINE CELL NAME "LOGRADOURO"	OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "LOGRADOURO"		 	//PICTURE "@!"
	DEFINE CELL NAME "NUMERO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "NUMERO"		 		//PICTURE "@!"
	DEFINE CELL NAME "COMPLEMENTO"	OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "COMPLMENTO"		 	//PICTURE "@!"
	DEFINE CELL NAME "BAIRRO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "BAIRRO"		 		//PICTURE "@!"
	DEFINE CELL NAME "CEP"			OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CEP"		 			//PICTURE "@!"
	DEFINE CELL NAME "CIDADE"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CIDADE"		 		//PICTURE "@!"
	DEFINE CELL NAME "ESTADO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "ESTADO"		 		//PICTURE "@!"
	
	For nLoop := 1 to nQtdCol

		DEFINE CELL NAME "COD"+AllTrim(Str(nLoop))		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "COD."+AllTrim(Str(nLoop)) //PICTURE "@!"
		DEFINE CELL NAME "QTD"+AllTrim(Str(nLoop))		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "QTD."+AllTrim(Str(nLoop)) //PICTURE "@!"
		DEFINE CELL NAME "DIAS"+AllTrim(Str(nLoop))		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DIAS"+AllTrim(Str(nLoop)) //PICTURE "@!"
		DEFINE CELL NAME "VALOR"+AllTrim(Str(nLoop))	OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "VALOR"+AllTrim(Str(nLoop)) PICTURE "@E 999,999,999.99"

	Next nLoop 

	DEFINE CELL NAME "BANCO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "COD. BANCO"		 	//PICTURE "@!"
	DEFINE CELL NAME "AGENCIA"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "COD. AGENCIA"		 	//PICTURE "@!"
	DEFINE CELL NAME "DIGAGE"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DIG. AGENCIA"		 	//PICTURE "@!"
	DEFINE CELL NAME "CTACORR"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CONTA CORRENTE"		//PICTURE "@!"
	DEFINE CELL NAME "DIGCTA"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DIG. CONTA"		 	//PICTURE "@!"
	DEFINE CELL NAME "FAVORECIDO"	OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "FAVORECIDO"		 	//PICTURE "@!"
	DEFINE CELL NAME "CPFFAVOREC"	OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CPF FAVORECIDO"		//PICTURE "@!"

Return(oReport)

/*/{Protheus.doc} ReportPrint()
Definicao do Relatorio
@type function
@version V 1.00
@author Maicon Brito
@since 29/09/2025
@param 
/*/
Static Function ReportPrint(oReport,cTitulo)

Local cQuery		:= ""
Local nTotReg		:= 100
Local nLoop			:= 0
Local nX			:= 0
Local cSit 			:= ""

Local cChvCgc		:= ""
Local cChvMat		:= ""


Private aDadosEMP	:= {}
Private aDadosFUN	:= {}
Private cAliasTrb := GetNextAlias()


// query para fazer a busca de empresa
/*cQuery	:= " SELECT M0_CGC CNPJ, M0_FULNAME NOME, M0_ENDCOB ENDERECO, M0_COMPCOB COMPLEM, M0_CEPCOB CEP, "
cQuery	+= " M0_BAIRCOB BAIRRO, M0_CIDCOB CIDADE, M0_ESTCOB UF "
cQuery	+= " FROM SYS_COMPANY WHERE D_E_L_E_T_ <> '*' "*/

cQuery	:= " SELECT M0_CGC CNPJ, M0_FULNAME NOME, M0_ENDCOB ENDERECO, M0_COMPCOB COMPLEM, M0_CEPCOB CEP, "
cQuery	+= " M0_BAIRCOB BAIRRO, M0_CIDCOB CIDADE, M0_ESTCOB UF , LTRIM(RTRIM(SUBSTR(RCC.RCC_CONTEU,9,LENGTH(RCC.RCC_CONTEU)))) RCC_CONTEUDO "
cQuery	+= " FROM SYS_COMPANY SYSCOMP "
cQuery	+= " INNER JOIN " + RetSqlName("RCC") + " RCC ON RCC.RCC_CODIGO = 'U001' "
cQuery	+= " AND LTRIM(RTRIM(RCC.RCC_FIL)) = LTRIM(RTRIM(M0_CODFIL)) AND RCC.D_E_L_E_T_ <> '*' "
cQuery	+= " WHERE SYSCOMP.D_E_L_E_T_ <> '*' "

//Executa a query e alimenta o array com o resultado 
aDadosEMP := QryArr(cQuery)


//query funcionários
cQuery	:= "  SELECT DISTINCT SYSCOMP.M0_CGC AS CNPJ, SRA.RA_MAT AS MATRICULA, SRA.RA_NOME AS NOME_COMPLETO, SRA.RA_EMAIL AS EMAIL, "
cQuery	+= " 		 SRA.RA_TELEFON AS CELULAR, SRA.RA_SITFOLH AS ATIVO, SRA.RA_ENDEREC AS ENDERECO, SQ3.Q3_DESCSUM AS CARGO, "
cQuery	+= " 		 SQB.QB_DESCRIC AS DEPARTAMENTO, SR0.R0_DUTILM AS DIAS_TRABALHADOS, SRA.RA_CIC AS CPF, SRA.RA_RG AS RG, '' AS DG, "
cQuery	+= " 		 SRA.RA_RGUF AS EST_RG, SRA.RA_NASC AS DT_NASCIMENTO, "
cQuery	+= "         SRA.RA_MAE AS NOME_MAE, SRA.RA_ENDEREC AS LOGRADOURO, SRA.RA_NUMENDE AS NUMERO, SRA.RA_COMPLEM AS COMPLEMENTO, "
cQuery	+= "         SRA.RA_BAIRRO AS BAIRRO, SRA.RA_CEP AS CEP, SRA.RA_MUNICIP AS CIDADE, SRA.RA_ESTADO AS ESTADO, SRN.RN_BNFFOR AS COD, "
cQuery	+= "         SR0.R0_CODIGO AS COD, SM7.M7_QDIAINF AS QTD, 0 AS DIAS, SR0.R0_VLRVALE AS VALOR, 
cQuery	+= " 		 SUBSTR(SRA.RA_BCDEPSA,1,3) AS COD_BANCO, SUBSTR(SRA.RA_BCDEPSA,5,4) AS COD_AGENCIA, '' AS DIG_AGENCIA, "
cQuery	+= " 		 SRA.RA_CTDEPSA AS CONTA_CORRENTE, '' AS DIG_CONTA, "
cQuery	+= "         SRA.RA_NOME AS FAVORECIDO, SRA.RA_CIC AS CPF_FAVORECIDO, LTRIM(RTRIM(SUBSTR(RCC.RCC_CONTEU,9,LENGTH(RCC.RCC_CONTEU)))) RCC_CONTEUDO, RTRIM(LTRIM(CTT.CTT_DESC01)) AS CCUSTO "
cQuery	+= "     FROM " + RetSqlName("SR0") + "  SR0 "
cQuery	+= "     INNER JOIN SYS_COMPANY SYSCOMP "
cQuery	+= "         ON SR0.R0_FILIAL = SYSCOMP.M0_CODFIL "
cQuery	+= "        AND SYSCOMP.D_E_L_E_T_ <> '*' "
cQuery	+= "        AND SYSCOMP.M0_CODIGO = '01' "
cQuery	+= "     INNER JOIN " + RetSqlName("SRA") + "  SRA "
cQuery	+= "         ON SR0.R0_FILIAL = SRA.RA_FILIAL "
cQuery	+= "        AND SR0.R0_MAT = SRA.RA_MAT "
cQuery	+= "        AND SRA.D_E_L_E_T_ <> '*' "
cQuery	+= "     INNER JOIN " + RetSqlName("SRN") + "  SRN "
cQuery	+= "         ON LTRIM(RTRIM(SRN.RN_FILIAL)) = SUBSTR(LTRIM(RTRIM(SR0.R0_FILIAL)),1,4) "
cQuery	+= "        AND SRN.RN_COD = SR0.R0_CODIGO "
cQuery	+= "        AND SRN.D_E_L_E_T_ <> '*' "
cQuery	+= "     LEFT OUTER JOIN " + RetSqlName("SQ3") + "  SQ3 "
cQuery	+= "         ON LTRIM(RTRIM(SQ3.Q3_FILIAL)) = SUBSTR(LTRIM(RTRIM(SRA.RA_FILIAL)),1,4) "
cQuery	+= "        AND SRA.RA_CARGO = SQ3.Q3_CARGO "
cQuery	+= "        AND SQ3.D_E_L_E_T_ <> '*' "
cQuery	+= "     INNER JOIN " + RetSqlName("SQB") + "  SQB "
cQuery	+= "         ON LTRIM(RTRIM(SRA.RA_FILIAL)) = LTRIM(RTRIM(SQB.QB_FILIAL)) "
cQuery	+= "        AND SRA.RA_DEPTO = SQB.QB_DEPTO "
cQuery	+= "        AND SQB.D_E_L_E_T_ <> '*' "
cQuery	+= "	 INNER JOIN " + RetSqlName("RCC") + " RCC  "
cQuery	+= "         ON RCC.RCC_CODIGO = 'U001' "
cQuery	+= "         AND LTRIM(RTRIM(RCC.RCC_FIL)) = LTRIM(RTRIM(M0_CODFIL)) "
cQuery	+= "         AND RCC.D_E_L_E_T_ <> '*' "
cQuery	+= "	 INNER JOIN " + RetSqlName("SM7") + " SM7  "
cQuery	+= "         ON SM7.M7_FILIAL = SR0.R0_FILIAL "
cQuery	+= "        AND SM7.M7_MAT = SR0.R0_MAT "
cQuery	+= "        AND SM7.M7_TPVALE = SR0.R0_TPVALE "
cQuery	+= "        AND SM7.M7_CODIGO = SR0.R0_CODIGO "
cQuery	+= "        AND SM7.D_E_L_E_T_ <> '*' "
cQuery	+= "     INNER JOIN " + RetSqlName("CTT") + "  CTT "
cQuery	+= "         ON CTT.CTT_FILIAL = SRA.RA_FILIAL "
cQuery	+= "        AND CTT.CTT_CUSTO = SRA.RA_CC "
cQuery	+= "        AND CTT.D_E_L_E_T_ <> '*' "
cQuery	+= " WHERE SR0.D_E_L_E_T_ <> '*' AND SRA.RA_XVCOMBU <> '1' AND SR0.R0_TPVALE = '0' "

cQuery	+= " AND SRA.RA_FILIAL BETWEEN '" + AllTrim(mv_par01) + "' AND '" + AllTrim(mv_par02) + "'

//IF !Empty(mv_par03)
	cQuery	+= " AND SRA.RA_CC BETWEEN '" + AllTrim(mv_par03) + "' AND '" + AllTrim(mv_par04) + "'
//Endif 

//IF !Empty(mv_par04)
	cQuery	+= " AND SRA.RA_MAT BETWEEN '" + AllTrim(mv_par05) + "' AND '" + AllTrim(mv_par06) + "'
//Endif 

IF !Empty(mv_par07)
	cQuery	+= " AND SRA.RA_NOME LIKE '%" + AllTrim(mv_par07) + "%'	"
Endif 

cQuery	+= " AND SRA.RA_SITFOLH	IN	" + FormatIn(cSitQuery,"/") + "	"
cQuery	+= " AND SRA.RA_CATFUNC	IN	" + FormatIn(cCatQuery,"/") + "	"

IF !Empty(mv_par11)
	cQuery	+= " AND SR0.R0_ROTEIR = '" + AllTrim(mv_par11) + "'	"
Endif 

IF !Empty(mv_par12)
	cQuery	+= " AND SR0.R0_PERIOD = '" + AllTrim(mv_par12) + "'	"
Endif 

cQuery	+= " ORDER BY CNPJ, MATRICULA "

cQuery := ChangeQuery(cQuery)

MemoWrite("\DATA\RGPER01.SQL",cQuery)

//Executa a query e alimenta o array com o resultado 

If Select(cAliasTrb) > 0
	dbSelectArea(cAliasTrb)
	dbCloseArea()
EndIf

TcQuery cQuery New Alias (cAliasTrb)

dbSelectArea(cAliasTrb)
dbGoTop()

oReport:SetMeter(nTotReg)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Impressao do relatorio                                                         	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Empresa                                                                        	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
For nLoop := 1 to Len(aDadosEMP)

    oSecEMP:Init()
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Verifica se nao foi cancelado           ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	If oReport:Cancel()
		Exit
	EndIf

   oSecEMP:Cell("CNPJ"):SetValue(aDadosEMP[nLoop][1])
   oSecEMP:Cell("NOME"):SetValue(aDadosEMP[nLoop][2])
   oSecEMP:Cell("ENDERECO"):SetValue(aDadosEMP[nLoop][3])
   oSecEMP:Cell("NUMERO"):SetValue("")
   oSecEMP:Cell("COMPLEM"):SetValue(aDadosEMP[nLoop][4])
   oSecEMP:Cell("CEP"):SetValue(aDadosEMP[nLoop][5])
   oSecEMP:Cell("BAIRRO"):SetValue(aDadosEMP[nLoop][6])
   oSecEMP:Cell("CIDADE"):SetValue(aDadosEMP[nLoop][7])
   oSecEMP:Cell("UF"):SetValue(aDadosEMP[nLoop][8])
   oSecEMP:Cell("CODIGO"):SetValue(aDadosEMP[nLoop][9])
   
   oSecEMP:PrintLine()
	
   oReport:IncMeter()

Next

oSecEMP:Finish()


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Funcionarios                                                               	  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
dbSelectArea(cAliasTrb)
dbGoTop()

COUNT TO nTotReg
oReport:SetMeter(nTotReg)

dbGoTop()
oSecFUN:Init()
While (cAliasTrb)->(!Eof())

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Verifica se nao foi cancelado           ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	If oReport:Cancel()
		Exit
	EndIf

	oSecFUN:Cell("CNPJ"):SetValue((cAliasTrb)->CNPJ)
	oSecFUN:Cell("MATRICULA"):SetValue((cAliasTrb)->MATRICULA)
	oSecFUN:Cell("NOME"):SetValue((cAliasTrb)->NOME_COMPLETO)
	oSecFUN:Cell("EMAIL"):SetValue((cAliasTrb)->EMAIL)
	oSecFUN:Cell("CELULAR"):SetValue((cAliasTrb)->CELULAR)

	If Empty(AllTrim((cAliasTrb)->ATIVO))
		cSit := "ATIVO"
	ElseIf (cAliasTrb)->ATIVO == "A"
		cSit := "INATIVO/AFASTADO"
	ElseIf (cAliasTrb)->ATIVO == "F"
		cSit := "FERIAS"
	ElseIf (cAliasTrb)->ATIVO == "D"
		cSit := "DEMITIDO"
	Endif 

	oSecFUN:Cell("ATIVO"):SetValue(cSit)
	oSecFUN:Cell("CODIGO"):SetValue((cAliasTrb)->RCC_CONTEUDO)
	oSecFUN:Cell("CARGO"):SetValue((cAliasTrb)->CARGO)
	oSecFUN:Cell("DEPTO"):SetValue((cAliasTrb)->CCUSTO)
	oSecFUN:Cell("DIASTRAB"):SetValue((cAliasTrb)->DIAS_TRABALHADOS)
	oSecFUN:Cell("CPF"):SetValue((cAliasTrb)->CPF)
	oSecFUN:Cell("RG"):SetValue((cAliasTrb)->RG)
	oSecFUN:Cell("DG"):SetValue((cAliasTrb)->DG)
	oSecFUN:Cell("EST_RG"):SetValue((cAliasTrb)->EST_RG)
	oSecFUN:Cell("DTNASC"):SetValue(DTOC(STOD((cAliasTrb)->DT_NASCIMENTO)))
	oSecFUN:Cell("NOMEMAE"):SetValue((cAliasTrb)->NOME_MAE)
	oSecFUN:Cell("LOGRADOURO"):SetValue((cAliasTrb)->LOGRADOURO)
	oSecFUN:Cell("NUMERO"):SetValue((cAliasTrb)->NUMERO)
	oSecFUN:Cell("COMPLEMENTO"):SetValue((cAliasTrb)->COMPLEMENTO)
	oSecFUN:Cell("BAIRRO"):SetValue((cAliasTrb)->BAIRRO)
	oSecFUN:Cell("CEP"):SetValue((cAliasTrb)->CEP)
	oSecFUN:Cell("CIDADE"):SetValue((cAliasTrb)->CIDADE)
	oSecFUN:Cell("ESTADO"):SetValue((cAliasTrb)->ESTADO)

	oSecFUN:Cell("BANCO"):SetValue((cAliasTrb)->COD_BANCO)
	oSecFUN:Cell("AGENCIA"):SetValue((cAliasTrb)->COD_AGENCIA)
	oSecFUN:Cell("DIGAGE"):SetValue((cAliasTrb)->DIG_AGENCIA)
	oSecFUN:Cell("CTACORR"):SetValue((cAliasTrb)->CONTA_CORRENTE)
	oSecFUN:Cell("DIGCTA"):SetValue((cAliasTrb)->DIG_CONTA)
	oSecFUN:Cell("FAVORECIDO"):SetValue((cAliasTrb)->FAVORECIDO)
	oSecFUN:Cell("CPFFAVOREC"):SetValue((cAliasTrb)->CPF_FAVORECIDO)

	cChvCgc := (cAliasTrb)->CNPJ
	cChvMat := (cAliasTrb)->MATRICULA
	nLoop := 0

	While (cAliasTrb)->(!Eof()) .And. (cAliasTrb)->CNPJ == cChvCgc .And. (cAliasTrb)->MATRICULA == cChvMat

		nLoop++

		If nLoop <= nQtdCol // Limita a quantidade de 10 colunas

			oSecFUN:Cell("COD"+AllTrim(Str(nLoop))):SetValue((cAliasTrb)->COD)
			oSecFUN:Cell("QTD"+AllTrim(Str(nLoop))):SetValue((cAliasTrb)->QTD)
			oSecFUN:Cell("DIAS"+AllTrim(Str(nLoop))):SetValue((cAliasTrb)->DIAS)
			oSecFUN:Cell("VALOR"+AllTrim(Str(nLoop))):SetValue((cAliasTrb)->VALOR)

		Endif 

		(cAliasTrb)->(dbSkip())

	EndDo

	//Limpar colunas restantes
	nLoop++
	For nX := nLoop to nQtdCol 

		oSecFUN:Cell("COD"+AllTrim(Str(nX))):SetValue("")
		oSecFUN:Cell("QTD"+AllTrim(Str(nX))):SetValue(0)
		oSecFUN:Cell("DIAS"+AllTrim(Str(nX))):SetValue(0)
		oSecFUN:Cell("VALOR"+AllTrim(Str(nX))):SetValue(0)

	Next nX 

	cChvCgc := ""
	cChvMat := ""
	nX := 0

   	oSecFUN:PrintLine()
	
   	oReport:IncMeter()
   
EndDo

oSecFUN:Finish()         

oReport:EndPage()

Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ AjustaSX1³ Autor ³  Maicon Brito         ³ Data ³ 08/03/13 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Cria o pergunte caso nao encontre no SX1                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³   	                                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function AjustaSX1(cPerg)

Local aArea1 := Getarea()
Local aRegs  := {}

Aadd(aRegs,{cPerg, "01","Filial De  ? "			,"","","mv_ch1","C",08,0,0,"G",""			 ,"mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SM0","","","","",""})
Aadd(aRegs,{cPerg, "02","Filial Ate ? "			,"","","mv_ch2","C",08,0,0,"G",""			 ,"mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SM0","","","","",""})
Aadd(aRegs,{cPerg, "03","Centro de Custo De ? "	,"","","mv_ch3","C",20,0,0,"G",""			 ,"mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","","","",""})
Aadd(aRegs,{cPerg, "04","Centro de Custo Ate ?"	,"","","mv_ch4","C",20,0,0,"G",""			 ,"mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","","","",""})
Aadd(aRegs,{cPerg, "05","Matrícula De ? " 		,"","","mv_ch5","C",06,0,0,"G",""			 ,"mv_par05","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","","","",""})
Aadd(aRegs,{cPerg, "06","Matrícula Ate ? "		,"","","mv_ch6","C",06,0,0,"G",""			 ,"mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","","","",""})
Aadd(aRegs,{cPerg, "07","Nome Contem ? "		,"","","mv_ch7","C",30,0,0,"G",""			 ,"mv_par07","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "08","Situações ? "			,"","","mv_ch8","C",05,0,0,"G","fSituacao()" ,"mv_par08","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "09","Categorias ? "			,"","","mv_ch9","C",15,0,0,"G","fCategoria()","mv_par09","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "10","Processo   ? "			,"","","mv_chA","C",05,0,0,"G",""	 		 ,"mv_par10","","","","","","","","","","","","","","","","","","","","","","","","","RCHBEN","","","","",""})
Aadd(aRegs,{cPerg, "11","Roteiro    ? "			,"","","mv_chB","C",03,0,0,"G","" 			 ,"mv_par11","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg, "12","Periodo	? "			,"","","mv_chC","C",06,0,0,"S",""			 ,"mv_par12","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
CriaSx1(cPerg,aRegs)                                                                                               

RestArea(aArea1)

Return Nil



/*/
//=================================================================================
// Funcao    | QryArr   |Autor  | 				     | Data |             
//=================================================================================
// Descricao | Funcao para rodar uma Query e retornar como Array          
//=================================================================================
// Parametros| cQuery - Query SQL a ser executado                         
// Retorno   | aTrb   - Array com o conteudo da Query                     
//=================================================================================
// Uso       | Exclusivo 
//=================================================================================
/*/
Static Function QryArr(cQuery)

//=================================================================================
// Gravacao do Ambiente Atual e Variaveis para Utilizacao                   
//=================================================================================
Local aRet    := {}
Local aRet1   := {}
Local nRegAtu := 0
Local x       := 0

//=================================================================================
// Ajustes e Execucao da Query                                              
//=================================================================================
cQuery := ChangeQuery(cQuery)
TCQUERY cQuery NEW ALIAS "_TRB"

//=================================================================================
// Montagem do Array para Retorno                                           
//=================================================================================
dbSelectArea("_TRB")
aRet1   := Array(Fcount())
nRegAtu := 1

While !Eof()
	
	For x:=1 To Fcount()
		aRet1[x] := FieldGet(x)
	Next
	Aadd(aRet,aclone(aRet1))

	//aRet[Len(aRet)][1] := .F.
	
	dbSkip()
	nRegAtu += 1
Enddo

//=================================================================================
// Encerra Query e Retorna Ambiente                                         
//=================================================================================
dbSelectArea("_TRB")
_TRB->(DbCloseArea())

Return(aRet)


/*/
//=================================================================================
// Funcao    | CriaSX1   |Autor  | 				     | Data |             
//=================================================================================
// Descricao | Criação de perguntas dentro do SX1        
//=================================================================================
// Uso       | Exclusivo 
//=================================================================================
/*/
Static Function CriaSX1(cPerg,aRegs)                                        

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Variaveis da Rotina                                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Local i 	:= 0
Local j 	:= 0    
Local aArtu := GetArea()

dbSelectArea("SX1")
SX1->(dbSetOrder(1))

For i:= 1 to Len(aRegs)
	If !dbSeek(PadR(cPerg,Len(X1_GRUPO))+aRegs[i,2])
		RecLock("SX1", .T.)
		For j := 1 to Len(aRegs[1])
			FieldPut(j,aRegs[i,j])
		Next j
		MsUnlock()
	EndIf
Next i     

RestArea(aArtu)

Return Nil
