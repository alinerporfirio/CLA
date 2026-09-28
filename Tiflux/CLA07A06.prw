#Include "Protheus.ch"
#Include "TopConn.ch"

/*/{Protheus.doc} CLA07A06

Job que chama a rotina do Tiflux CLA07A04
Autor: Aline Rocha Porfirio
Data: 25/09/2026

/*/

User Function CLA07A06(aParam)

    Local cAlias   := ""
    Local cEmp     := IIf(ValType(aParam) == "A", aParam[1], "01")
    Local cFil     := IIf(ValType(aParam) == "A", aParam[2], "01")
    Local aTipoInt := {}
    Local cFilAnt2 := ""

    RpcSetType(3)
    RpcSetEnv(cEmp, cFil)

    // Evita duas execuções simultâneas do job
    If !LockByName("CLA07A06", .T., .F.)

        ConOut("[CLA07A06] Job já em execução.")
        RpcClearEnv()
    
        Return
    
    EndIf

    cAlias := GetNextAlias()

    BeginSql Alias cAlias
        SELECT RH3.RH3_FILIAL, RH3.RH3_CODIGO, RH3.RH3_TIPO
        FROM %Table:RH3% RH3
        WHERE RH3.%NotDel%
          AND RH3.RH3_STATUS = '4'
          AND RH3.RH3_XTIFLU = ' '
        ORDER BY RH3.RH3_FILIAL, RH3.RH3_CODIGO
    EndSql

    While !(cAlias)->(Eof())

        // Relê o parâmetro só quando muda a filial (mesma lógica do PontoRH)
        If (cAlias)->RH3_FILIAL <> cFilAnt2

            cFilAnt2 := (cAlias)->RH3_FILIAL
            aTipoInt := StrTokArr(SuperGetMV("MV_XSOLTF", .F., "", cFilAnt2), ",")

        EndIf

        If AScan(aTipoInt, {|x| AllTrim(x) == AllTrim((cAlias)->RH3_TIPO)}) > 0

            U_CLA07A04((cAlias)->RH3_FILIAL, (cAlias)->RH3_CODIGO)

        EndIf

        (cAlias)->(DbSkip())

    EndDo

    (cAlias)->(DbCloseArea())
    UnLockByName("CLA07A06", .T., .F.)
    RpcClearEnv()

Return
