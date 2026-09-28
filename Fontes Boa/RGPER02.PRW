#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#Include "Report.ch"

/*/{Protheus.doc} RGPER02()
Função para gerar relatório em excel conforme layout pré-definido - IFood
@type function
@version V 1.00
@author Maicon Brito
@since 20/10/2025
@param 
/*/
User Function RGPER02()

Local nReg			:= 0

Private cPerg 		:= "RGPER02"
Private oReport
Private oBreak
Private oSecEMP 	:= NIL
Private oSecFUN 	:= NIL

Private cSituacao   := ""          //  Situacoes
Private cCategoria  := ""          //  Categorias

Private cSitQuery	:= ""
Private cCatQuery	:= ""

AjustaSX1(cPerg)
If !Pergunte(cPerg, .T.)
	Return()
Endif 

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
@since 20/10/2025
@param 
/*/
Static Function ReportDef()

Local cTitulo	:= "Relatório IFOOD"

oReport:= TReport():New("RGPER02",cTitulo,cPerg, {|oReport| ReportPrint(oReport,cTitulo)},"Este relatório irá imprimir a relacao de funcionario")
oReport:SetLandscape() //Define a orientacao default de pagina do relatorio como retrato. SetPortrait()/SetLandscape()
//oReport:HideParamPage()   // Desabilita a impressao da pagina de parametros.
oReport:PrintHeader(.T.,.T.)
oReport:nFontBody := 10 // Define o tamanho da fonte.


DEFINE SECTION oSecFUN OF oReport TABLES "QRY" TITLE "Usuarios"


	DEFINE CELL NAME "CNPJ"			OF oSecFUN ALIAS "QRY" SIZE 30 TITLE "CNPJ"		 			PICTURE "@R 99.999.999/9999-99"
	DEFINE CELL NAME "NOME"			OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "NOME COMPLETO"		//PICTURE "@!"
	DEFINE CELL NAME "CPF"			OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "CPF"		 			//PICTURE "@!"
	DEFINE CELL NAME "DTNASC"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "DATA NASCIMENTO"	 	//PICTURE "@!"
	DEFINE CELL NAME "EMAIL"		OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "EMAIL"		 		//PICTURE "@!"
	DEFINE CELL NAME "CELULAR"		OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "CELULAR"				// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "CCUSTO"		OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "CENTRO_CUSTO"			// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "CONVENCAO"	OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "CONVENCAO_COLETIVA"	// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "GRPENTREG"	OF oSecFUN ALIAS "QRY" SIZE 30 TITLE "GRUPO_ENTREGA"		// PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "MATRICULA"	OF oSecFUN ALIAS "QRY" SIZE 20 TITLE "MATRICULA"		 	//PICTURE "@!"
	DEFINE CELL NAME "FILTROREL"	OF oSecFUN ALIAS "QRY" SIZE 15 TITLE "FILTRO_REL_REGARGA"	// PICTURE "@E 999,999,999.99"
	
	DEFINE CELL NAME "REFEICAO"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "REFEICAO"				 PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "ALIMENTA"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "ALIMENTACAO"			 PICTURE "@E 999,999,999.99"
	DEFINE CELL NAME "MOBILIDA"		OF oSecFUN ALIAS "QRY" SIZE 40 TITLE "MODILIDADE"			 PICTURE "@E 999,999,999.99"


Return(oReport)

/*/{Protheus.doc} ReportPrint()
Definicao do Relatorio
@type function
@version V 1.00
@author Maicon Brito
@since 20/10/2025
@param 
/*/
Static Function ReportPrint(oReport,cTitulo)

Local cQuery		:= ""
Local nTotReg		:= 100

Private aDadosFUN	:= {}
Private cAliasTrb := GetNextAlias()


//query funcionários
cQuery	:= " SELECT SUBQRY.CNPJ, SUBQRY.RCC_CONTEUDO, SUBQRY.MATRICULA,SUBQRY.NOME_COMPLETO,SUBQRY.EMAIL,SUBQRY.CELULAR,SUBQRY.ATIVO,SUBQRY.ENDERECO,SUBQRY.CARGO,SUBQRY.DEPARTAMENTO, "
cQuery	+= "     SUBQRY.CPF,SUBQRY.RG,SUBQRY.DG,SUBQRY.EST_RG,SUBQRY.DT_NASCIMENTO,SUBQRY.NOME_MAE,SUBQRY.LOGRADOURO,SUBQRY.NUMERO,SUBQRY.COMPLEMENTO,SUBQRY.BAIRRO,SUBQRY.CEP, "
cQuery	+= "     SUBQRY.CIDADE,SUBQRY.ESTADO, "

cQuery	+= "     SUM(CASE WHEN SUBQRY.TPVALE = '1' 	 THEN SUBQRY.VALOR ELSE 0 END) AS REFEICAO, "
cQuery	+= "     SUM(CASE WHEN SUBQRY.TPVALE = '2' 	 THEN SUBQRY.VALOR ELSE 0 END) AS ALIMENTACAO, "
cQuery	+= "     SUM(CASE WHEN SUBQRY.TPVALE = '0' AND SUBQRY.XVCOMBU = '1'	 THEN SUBQRY.VALOR ELSE 0 END) AS MOBILIDADE, "

cQuery	+= "     SUBQRY.COD_BANCO,SUBQRY.COD_AGENCIA,SUBQRY.DIG_AGENCIA,SUBQRY.CONTA_CORRENTE,SUBQRY.DIG_CONTA,SUBQRY.FAVORECIDO,SUBQRY.CPF_FAVORECIDO, SUBQRY.CCUSTO, SUBQRY.GRPENTREGA "
cQuery	+= "     FROM "
cQuery	+= "     	(SELECT DISTINCT SYSCOMP.M0_CGC AS CNPJ,LTRIM(RTRIM(SUBSTR(RCC.RCC_CONTEU,9,LENGTH(RCC.RCC_CONTEU)))) RCC_CONTEUDO, SRA.RA_MAT AS MATRICULA,SRA.RA_NOME AS NOME_COMPLETO,SRA.RA_EMAIL AS EMAIL,SRA.RA_TELEFON AS CELULAR,SRA.RA_SITFOLH AS ATIVO, "
cQuery	+= "     		SRA.RA_ENDEREC AS ENDERECO,SRA.RA_CARGO AS CARGO,SRA.RA_DEPTO AS DEPARTAMENTO,SR0.R0_DUTILM AS DIAS_TRABALHADOS,SRA.RA_CIC AS CPF,SRA.RA_RG AS RG,' ' AS DG, "
cQuery	+= "     		SRA.RA_RGUF AS EST_RG,SRA.RA_NASC AS DT_NASCIMENTO,SRA.RA_MAE AS NOME_MAE,SRA.RA_ENDEREC AS LOGRADOURO,SRA.RA_NUMENDE AS NUMERO,SRA.RA_COMPLEM AS COMPLEMENTO, "
cQuery	+= "     		SRA.RA_BAIRRO AS BAIRRO,SRA.RA_CEP AS CEP,SRA.RA_MUNICIP AS CIDADE,SRA.RA_ESTADO AS ESTADO, "
cQuery	+= "     		SR0.R0_CODIGO AS R0CODIGO,SR0.R0_QDIACAL AS QTD,SR0.R0_DUTILM AS DIAS,SR0.R0_VALCAL AS VALOR, "
cQuery	+= "     		SUBSTR(SRA.RA_BCDEPSA,1,3) AS COD_BANCO,SUBSTR(SRA.RA_BCDEPSA,5,4) AS COD_AGENCIA,' ' AS DIG_AGENCIA, "
cQuery	+= "     		SRA.RA_CTDEPSA AS CONTA_CORRENTE,' ' AS DIG_CONTA,SRA.RA_NOME AS FAVORECIDO,SRA.RA_CIC AS CPF_FAVORECIDO, SRA.RA_XVCOMBU XVCOMBU, SR0.R0_TPVALE TPVALE, SRA.RA_CC CCUSTO, SYSCOMP.M0_FILIAL AS GRPENTREGA "
cQuery	+= "     		FROM " + RetSqlName("SR0") + " SR0 INNER JOIN SYS_COMPANY SYSCOMP ON SR0.R0_FILIAL = SYSCOMP.M0_CODFIL AND SYSCOMP.D_E_L_E_T_ <> '*' "
cQuery	+= "     		INNER JOIN " + RetSqlName("RCC") + " RCC ON RCC.RCC_CODIGO = 'U002' AND LTRIM(RTRIM(RCC.RCC_FIL)) = LTRIM(RTRIM(SYSCOMP.M0_CODFIL)) AND RCC.D_E_L_E_T_ <> '*' "
cQuery	+= "     		INNER JOIN " + RetSqlName("SRA") + " SRA ON SR0.R0_FILIAL = SRA.RA_FILIAL AND SR0.R0_MAT = SRA.RA_MAT AND SRA.D_E_L_E_T_ <> '*' " 
cQuery	+= "    WHERE  SR0.D_E_L_E_T_ <> '*'  "

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

IF !Empty(AllTrim(mv_par11))
	cQuery	+= " AND SR0.R0_ROTEIR IN " + FormatIn(AllTrim(mv_par11),"/")  + "	"
Endif 

IF !Empty(mv_par12)
	cQuery	+= " AND SR0.R0_PERIOD = '" + AllTrim(mv_par12) + "'	"
Endif 


cQuery	+= " ) SUBQRY "

cQuery	+= "  GROUP BY SUBQRY.CNPJ, SUBQRY.RCC_CONTEUDO, SUBQRY.MATRICULA, SUBQRY.NOME_COMPLETO, SUBQRY.EMAIL, SUBQRY.CELULAR, SUBQRY.ATIVO, SUBQRY.ENDERECO, SUBQRY.CARGO, SUBQRY.DEPARTAMENTO, "
cQuery	+= "  SUBQRY.CPF, SUBQRY.RG, SUBQRY.DG, SUBQRY.EST_RG, SUBQRY.DT_NASCIMENTO, SUBQRY.NOME_MAE, SUBQRY.LOGRADOURO, SUBQRY.NUMERO, "
cQuery	+= "  SUBQRY.COMPLEMENTO, SUBQRY.BAIRRO, SUBQRY.CEP, SUBQRY.CIDADE, SUBQRY.ESTADO, SUBQRY.COD_BANCO, SUBQRY.COD_AGENCIA, SUBQRY.DIG_AGENCIA, "
cQuery	+= "  SUBQRY.CONTA_CORRENTE, SUBQRY.DIG_CONTA, SUBQRY.FAVORECIDO, SUBQRY.CPF_FAVORECIDO, SUBQRY.CCUSTO, SUBQRY.GRPENTREGA  "
cQuery	+= "  ORDER BY  SUBQRY.CNPJ, SUBQRY.MATRICULA 

cQuery := ChangeQuery(cQuery)

MemoWrite("\DATA\RGPER02.SQL",cQuery)


//aDadosFUN := QryArr(cQuery)

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

	oSecFUN:Cell("CNPJ"):SetValue((cAliasTrb)->RCC_CONTEUDO)
	oSecFUN:Cell("NOME"):SetValue((cAliasTrb)->NOME_COMPLETO)
	oSecFUN:Cell("CPF"):SetValue((cAliasTrb)->CPF)
	oSecFUN:Cell("DTNASC"):SetValue(DTOC(STOD((cAliasTrb)->DT_NASCIMENTO)))
	oSecFUN:Cell("EMAIL"):SetValue((cAliasTrb)->EMAIL)
	oSecFUN:Cell("CELULAR"):SetValue((cAliasTrb)->CELULAR)
	oSecFUN:Cell("CCUSTO"):SetValue((cAliasTrb)->CCUSTO)
	oSecFUN:Cell("GRPENTREG"):SetValue((cAliasTrb)->GRPENTREGA)
	oSecFUN:Cell("MATRICULA"):SetValue((cAliasTrb)->MATRICULA)
	oSecFUN:Cell("REFEICAO"):SetValue((cAliasTrb)->REFEICAO)
	oSecFUN:Cell("ALIMENTA"):SetValue((cAliasTrb)->ALIMENTACAO)
	oSecFUN:Cell("MOBILIDA"):SetValue((cAliasTrb)->MOBILIDADE)
  

   	oSecFUN:PrintLine()
	
   	oReport:IncMeter()
   
	(cAliasTrb)->(dbSkip())

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

Aadd(aRegs,{cPerg, "01","Filial De ? "			,"","","mv_ch1","C",08,0,0,"G",""			 ,"mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SM0","","","","",""})
Aadd(aRegs,{cPerg, "02","Filial Ate ? "			,"","","mv_ch2","C",08,0,0,"G",""			 ,"mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SM0","","","","",""})
Aadd(aRegs,{cPerg, "03","Centro de Custo De ? "	,"","","mv_ch3","C",20,0,0,"G",""			 ,"mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","","","",""})
Aadd(aRegs,{cPerg, "04","Centro de Custo Ate ?"	,"","","mv_ch4","C",20,0,0,"G",""			 ,"mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","","","",""})
Aadd(aRegs,{cPerg, "05","Matrícula De ? " 		,"","","mv_ch5","C",06,0,0,"G",""			 ,"mv_par05","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","","","",""})
Aadd(aRegs,{cPerg, "06","Matrícula Ate ? "		,"","","mv_ch6","C",06,0,0,"G",""			 ,"mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","","","",""})
Aadd(aRegs,{cPerg, "07","Nome Contem ? "		,"","","mv_ch7","C",30,0,0,"G",""			 ,"mv_par07","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "08","Situações ? "			,"","","mv_ch8","C",05,0,0,"G","fSituacao()" ,"mv_par08","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "09","Categorias ? "			,"","","mv_ch9","C",15,0,0,"G","fCategoria()","mv_par09","","","","","","","","","","","","","","","","","","","","","","","","",""   ,"","","","",""})
Aadd(aRegs,{cPerg, "10","Processo   ? "			,"","","mv_chA","C",05,0,0,"G",""	 		 ,"mv_par10","","","","","","","","","","","","","","","","","","","","","","","","","RCHBEN","","","","",""})
Aadd(aRegs,{cPerg, "11","Roteiro    ? "			,"","","mv_chB","C",20,0,0,"G","" 			 ,"mv_par11","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
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
