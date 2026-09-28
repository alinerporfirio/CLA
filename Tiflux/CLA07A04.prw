#INCLUDE "TOTVS.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "APWIZARD.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "FWBROWSE.CH"
#INCLUDE "FWMVCDEF.CH"

/*/{Protheus.doc} CLA07A04

Rotina que faz a integracao do Protheus com o Tiflux
Autor: Aline Rocha Porfirio
Data: 25/09/2026

/*/

User Function CLA07A04(cCodFilS,cSoliCod)

    Local cUrl          := "https://api.tiflux.com"
    Local _cRecURL      := "/api/v2/tickets"
    Local aHeader       := {}
    Local cError        := ''
    Local nStatus       := 0
    Local aRet          := {.T.,''}
    Local cToken        := SUPERGETMV('MV_XTTIFLU', .F.)
    Local cArquivo      := "\log\CLA07A04_Tiflux_Erro_"+DTos(Date())+"_"+StrTran( Time(), ':', '' )+".log"
    Local cLog          := ''
    Local cBody         := ''
    Local cTexto        := ''
    Local cTitulo       := ''
    Local cBoundary     := "----AdvPLBoundary" + AllTrim(Str(Seconds()*1000,15,0))
    Local aFields       := {}
    Local nI            := 0
    Local cAliasQry     := GetNextAlias()
    Local nRecRH3       := 0
    Local cTicTiflux    := ''
    Local cConteudo     := ''
    Local oRest         as object
    Local oResult       as object

    BeginSql alias cAliasQry

        SELECT
            RH4.RH4_FILIAL,
            RH4.RH4_CODIGO,
            RH3.RH3_TIPO,
            RH4.RH4_CAMPO,
            Isnull(SX3.X3_TITULO,'') AS DESC_CAMPO,
            RH4.RH4_VALNOV,
            SX5.X5_DESCRI AS TIPO,
            RH3.R_E_C_N_O_ RECRH3
        FROM %Table:RH4% RH4

        INNER JOIN %Table:RH3% RH3 ON RH3.%NotDel%
            AND RH3.RH3_FILIAL = RH4.RH4_FILIAL
            AND RH3.RH3_CODIGO = RH4.RH4_CODIGO

        LEFT JOIN SX3010 SX3
            ON SX3.X3_CAMPO = RH4.RH4_CAMPO
            AND SX3.%NotDel%

        LEFT JOIN SX5010 SX5 
            ON RH3.RH3_TIPO = SX5.X5_CHAVE 
            AND X5_TABELA = 'JQ' 
            AND X5_FILIAL = ''
            AND SX5.%NotDel%

        WHERE RH4.RH4_CODIGO = %exp:cSoliCod%
            AND RH4.RH4_FILIAL = %exp:cCodFilS%
            AND RH4.%NotDel%

    EndSql

    nRecRH3 := (cAliasQry)->RECRH3

    cTitulo := '[APP MEU RH] Filial: '+cCodFilS+' - Codigo: '+cSoliCod+ ' - Tipo: '+(cAliasQry)->TIPO

    cTexto := "<b>Solicitacao inserida via APP Meu RH</b> <br><br>"  
    cTexto += "<b>Tipo: </b>"+(cAliasQry)->TIPO+"<br><br>"
    cTexto += "<b>Dados da solicitação: </b><br>"

    While (cAliasQry)->(!Eof())

        If alltrim((cAliasQry)->RH4_VALNOV) == '.F.'

            cConteudo := 'NÃO'

        ElseIf alltrim((cAliasQry)->RH4_VALNOV) == '.T.'

            cConteudo := 'SIM'
        
        Else

            cConteudo := (cAliasQry)->RH4_VALNOV

        EndIf

        If Empty((cAliasQry)->DESC_CAMPO)

            cTexto += "<b>"+RH4Desc((cAliasQry)->RH4_CAMPO) +":</b> "+ cConteudo + "<br>"

        Else

            cTexto += "<b>"+(cAliasQry)->DESC_CAMPO +":</b> "+ cConteudo + "<br>"

        EndIF

        (cAliasQry)->(dbSkip())

    EndDo

    (cAliasQry)->(dbCloseArea())
   
    cTitulo := EncodeUTF8(cTitulo, "cp1252")
    cTexto  := EncodeUTF8(cTexto , "cp1252")

    aAdd(aFields, {"client_id"                 , SUPERGETMV('MV_XCLITF',.F.,'',cCodFilS)    })
    aAdd(aFields, {"desk_id"                   , "67433"                                    })
    aAdd(aFields, {"title"                     , cTitulo                                    })
    aAdd(aFields, {"description"               , cTexto                                     })
    aAdd(aFields, {"services_catalogs_item_id" , "1492150"                                  })
    aAdd(aFields, {"requestor_name"            , "API Protheus - APP Meu RH"                })
    aAdd(aFields, {"requestor_email"           , "aline.porfirio_ext@claglobal.com.br"      })

    For nI := 1 To Len(aFields)

        cBody += "--" + cBoundary + CRLF
        cBody += 'Content-Disposition: form-data; name="' + aFields[nI][1] + '"' + CRLF + CRLF
        cBody += aFields[nI][2] + CRLF

    Next nI

    cBody += "--" + cBoundary + "--" + CRLF

    aAdd(aHeader, "User-Agent: Protheus")
    aAdd(aHeader, "Accept: application/json")
    aAdd(aHeader, "Content-Type: multipart/form-data; boundary=" + cBoundary)
    aAdd(aHeader, "Authorization: Bearer " + AllTrim(cToken))

    oRest := FWRest():New(cUrl)
    oRest:SetPath(_cRecURL)

    oRest:SetPostParams(cBody)

    If oRest:Post(aHeader)

        nStatus := HTTPGetStatus(@cError)

        if nStatus >= 200 .And. nStatus <= 299

            aRet[1] := .T.
            aRet[2] := 'Sucesso: ' + oRest:getResult()

            FWJsonDeserialize(oRest:getResult(), @oResult)

            cTicTiflux	:= alltrim(STR(oResult:ticket:ticket_number))

            RH3->(DbGoTo(nRecRH3))

            RecLock("RH3",.F.)
            
                RH3->RH3_XTIFLU := cTicTiflux	
            
            RH3->(MsUnlock())

        else

            aRet[1] := .F.
            aRet[2] := "HTTP Status " + cValToChar(nStatus) + " - " + oRest:getResult()

            cLog += "Integração Tiflux " + CRLF+ CRLF

            cLog += "[Integrou] " + IIf(aRet[1], "TRUE", "FALSE") + CRLF+ CRLF
            cLog += "[Retorno] " + aRet[2] + CRLF + CRLF
            cLog += "[Corpo] "+ CRLF+ CRLF + cBody + CRLF + CRLF

            MemoWrite(cArquivo, cLog)

        endif

    else

        aRet[1] := .F.
        
        If Valtype(oRest:getLastError()) == 'C' .And. !Empty(oRest:getLastError())

            aRet[2] := oRest:getLastError() + CRLF + oRest:getResult()
        
        Else
        
            aRet[2] := 'Erro no envio da requisicao HTTP. Resultado: ' + oRest:getResult()
        
        EndIf

        cLog += "Integração Tiflux " + CRLF+ CRLF

        cLog += "[Integrou] " + IIf(aRet[1], "TRUE", "FALSE") + CRLF+ CRLF
        cLog += "[Retorno] " + aRet[2] + CRLF + CRLF
        cLog += "[Corpo] "+ CRLF+ CRLF + cBody + CRLF + CRLF

        MemoWrite(cArquivo, cLog)

    endif

    FreeObj(oRest)

Return

Static Function RH4Desc(cCampo)

	Local cDesc     := ''
	Local aDescCampos := {{"TMP_1P13SL"  , "1a Parcela 13o Salario"},; //"1a Parcela 13o Salario"
							{"TMP_ABOND"   , "Abono Descricao"},; //"Abono Descricao"
							{"TMP_ABONO"   , "Abono Pecuniario"},; //"Abono Pecuniario"
							{"TMP_COD"     , "Codigo"},; //"Codigo"
							{"TMP_CONTAT"  , "Nome do Contato"},; //"Nome do Contato"
							{"TMP_DABONO"  , "Dias Abono"},; //"Dias Abono"
							{"TMP_DCARGO"  , "Desc. Sumaria"},; //"Desc. Sumaria"
							{"TMP_DCC"     , "Desc. Centro Custo"},; //"Desc. Centro Custo"
							{"TMP_DCCP"    , "Desc. Centro Custo"},; //"Desc. Centro Custo"
							{"TMP_DDEPTO"  , "Desc. Departamento"},; //"Desc. Departamento"
							{"TMP_DDEPTOP" , "Desc. Departamento"},; //"Desc. Departamento"
							{"TMP_DESC"    , "Descricao"},; //"Descricao"
							{"TMP_DFUNCA"  , "Desc. Funcao"},; //"Desc. Funcao"
							{"TMP_DIAREQ"  , "Dias Licenca"},; //"Dias Licenca"
							{"TMP_DOPORT"  , "Dias Oport."},; //"Dias Oport."
							{"TMP_DPER1"   , "Dias Prim. Per."},; //"Dias Prim. Per."
							{"TMP_DPER2"   , "Dias Segun. Per."},; //"Dias Segun. Per."
							{"TMP_DPROCP"  , "Desc. Processo"},; //"Desc. Processo"
							{"TMP_DSCFIL"  , "Desc. Filial"},; //"Desc. Filial"
							{"TMP_DTBFIM"  , "Data Base Final"},; //"Data Base Final"
							{"TMP_DTBINI"  , "Data Base Inicial"},; //"Data Base Inicial"
							{"TMP_DTFIM"   , "Data Final"},; //"Data Final"
							{"TMP_DTFIM1"  , "Primeira Data Final"},; //"Primeira Data Final"
							{"TMP_DTFIM2"  , "Segunda Data Final"},; //"Segunda Data Final"
							{"TMP_DTINI"   , "Data Inicial"},; //"Data Inicial"
							{"TMP_DTINI1"  , "Primeira Data Inicial"},; //"Primeira Data Inicial"
							{"TMP_DTINI2"  , "Segunda Data Inicial"},; //"Segunda Data Inicial"
							{"TMP_FILIAL"  , "Filial"},; //"Filial"
							{"TMP_FSUB1"   , "Primeira Filial Sub."},; //"Primeira Filial Sub."
							{"TMP_FSUB2"   , "Segunda Filial Sub."},; //"Segunda Filial Sub."
							{"TMP_FSUBST"  , "Filial Sub."},; //"Filial Sub."
							{"TMP_MAT"     , "Matricula"},; //"Matricula"
							{"TMP_MSUB1"   , "Primeira Mat. Sub."},; //"Primeira Mat. Sub."
							{"TMP_MSUB2"   , "Segunda Mat. Sub."},; //"Segunda Mat. Sub."
							{"TMP_MSUBST"  , "Matricula Sub."},; //"Matricula Sub."
							{"TMP_NMCURS"  , "Nome do Curso"},; //"Nome do Curso"
							{"TMP_NMINST"  , "Nome da Instituicao"},; //"Nome da Instituicao"
							{"TMP_NOME"    , "Nome"},; //"Nome"
							{"TMP_NOTA"    , "Nota"},; //"Nota"
							{"TMP_NOVAC"   , "Gera Nova Contratacao?"},; //"Gera Nova Contratacao?"
							{"TMP_NOVACO"  , "Novo Contrato"},; //"Novo Contrato"
							{"TMP_NSUB1"   , "Primeiro Nome Sub."},; //"Primeiro Nome Sub."
							{"TMP_NSUB2"   , "Segundo Nome Sub."},; //"Segundo Nome Sub."
							{"TMP_NSUBST"  , "Nome Sub."},; //"Nome Sub."
							{"TMP_PD"      , "Verba"},; //"Verba"
							{"TMP_POSTO"   , "Posto"},; //"Posto"
							{"TMP_QTDEPA"  , "Qtde. Parcelas"},; //"Qtde. Parcelas"
							{"TMP_RAZAO"   , "Motivo Desligamento"},; //"Motivo Desligamento"
							{"TMP_REGID"   , "Id Registro"},; //"Id Registro"
							{"TMP_SEQ"     , "Sequencia"},; //"Sequencia"
							{"TMP_SITUAC"  , "Situacao"},; //"Situacao"
							{"TMP_TABELA"  , "Tabela"},; //"Tabela"
							{"TMP_TELEFO"  , "Telefone"},; //"Telefone"
							{"TMP_TEST"    , "Teste"},; //"Teste"
							{"TMP_TIPO"    , "Tipo"},; //"Tipo"
							{"TMP_TPDESC"  , "Tipo Dia Desc."},; //"Tipo Dia Desc."
							{"TMP_VAGA"    , "Vaga"},; //"Vaga"
							{"TMP_VLRMEN"  , "Valor Mensal"},; //"Valor Mensal"
							{"TMP_TEXT"    , "Justificativa"},; //"Justificativa"
							{"TMP_MOTIVO"  , "Motivo do afastamento"},; //"Motivo do afastamento"
							{"TMP_OBS"     , "Justificativa"},; //"Justificativa"
							{"TMP_DIRECT"  , "Direcao"} } //"Direcao"


    If (nPos := aScan(aDescCampos, {|x| Alltrim(x[1]) == AllTrim(cCampo)})) > 0
        cDesc := aDescCampos[nPos][2]
    EndIf

Return cDesc
