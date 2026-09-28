#Include 'TopConn.ch'
#Include 'FwMVCDef.ch'
#Include 'Protheus.ch'
#Include 'Parmtype.ch'

#Define cTitApp "Cadastro de Querys"
#Define Enter Chr(13) + Chr(10) 

/*/{Protheus.doc} CadGenQry
Tela para realizar o Cadastro de Querys para API de Consulta Genérica em MVC
@type function
@version V 1.00
@author Edson Hornberger
@since 03/01/2024
/*/
User Function CadGenQry

Local aArea      := GetArea() as Array
Local cFunNameBk := FunName() as Character
Local cAlias     := "EZX"     as Character
Local oBrowse    := Nil       as Object

If !ChkFile(cAlias)
		
    Aviso(cTitApp, "Erro ao tentar abrir a Tabela " + cAlias + "!", {"Ok"}, 1, "Erro")
    Return()
    
EndIf
	
SetFunName("CadGenQry")

oBrowse := FwMBrowse():New()
oBrowse:SetAlias(cAlias)
oBrowse:SetDescription("Cadastro de Querys")
oBrowse:Activate()

SetFunName(cFunNameBk)

RestArea(aArea)

Return Nil


/*/{Protheus.doc} MenuDef
Função para criar Menu Padrão para a Rotina
@type function
@version V 1.00
@author Edson Hornberger
@since 03/01/2024
@return array, Array com as funções do Menu
/*/
Static Function MenuDef()

Local aRotina := {} as Array

ADD OPTION aRotina TITLE "Visualizar"   ACTION "VIEWDEF.CadGenQry" OPERATION MODEL_OPERATION_VIEW   ACCESS 0
ADD OPTION aRotina TITLE "Incluir" 	    ACTION "VIEWDEF.CadGenQry" OPERATION MODEL_OPERATION_INSERT ACCESS 0
ADD OPTION aRotina TITLE "Alterar" 	    ACTION "VIEWDEF.CadGenQry" OPERATION MODEL_OPERATION_UPDATE	ACCESS 0
ADD OPTION aRotina TITLE "Excluir" 	    ACTION "VIEWDEF.CadGenQry" OPERATION MODEL_OPERATION_DELETE	ACCESS 0

Return(aRotina)


/*/{Protheus.doc} ModelDef
Função para gerar Modelo de Dados 
@type function
@version V 1.00
@author Edson Hornberger
@since 03/01/2024
@return object, Objeto com Modelo de Dados
/*/
Static Function ModelDef()

Local oModel := Nil                     as Object
Local oStCab := FWFormStruct(1, 'EZX' ) as Object

//Criando o FormModel, adicionando o Cabeçalho e Grid
oModel := MpFormModel():New("MODEZX",,)
oModel:AddFields("EZXMASTER", NIL, oStCab) //, oPreExc)
//Setando outras informações do Modelo de Dados
oModel:SetDescription("Manutenção " + cTitApp)
oModel:SetPrimaryKey({'EZX_FILIAL', 'EZX_GRUPO'})

Return(oModel)


/*/{Protheus.doc} ViewDef
Função para gerar View de Modelo de Dados
@type function
@version V 1.00
@author Edson Hornberger
@since 03/01/2024
@return object, Objeto com View de Modelo de Dados
/*/
Static Function ViewDef()

// INSTANCIA A VIEW
Local oView    := FwFormView():New()       as Object
// INSTANCIA AS SUBVIEWS
Local oStruEZX := FwFormStruct(2, "EZX")   as Object
// RECEBE O MODELO DE DADOS
Local oModel   := FwLoadModel("CadGenQry") as Object

// INDICA O MODELO DA VIEW
oView:SetModel(oModel)
// CRIA ESTRUTURA VISUAL DE CAMPOS
oView:AddField("VIEW_EZX", oStruEZX, "EZXMASTER")
// CRIA BOXES HORIZONTAIS
oView:CreateHorizontalBox("TELA", 100)
// RELACIONA OS BOXES COM AS ESTRUTURAS VISUAIS
oView:SetOwnerView("VIEW_EZX", "TELA")
// DEFINE OS TÍTULOS DAS SUBVIEWS
oView:EnableTitleView("VIEW_EZX", "Querys")
oView:SetCloseOnOk({|| .T.})

Return(oView)
