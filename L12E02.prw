#INCLUDE "Totvs.ch"
#INCLUDE "Topconn.ch"
#INCLUDE "Tbiconn.ch"

//! Alinhamento
#DEFINE LEFT   1
#DEFINE CENTER 2
#DEFINE RIGHT  3

//! Formatação
#DEFINE GERAL     1
#DEFINE NUMERO    2
#DEFINE MONETARIO 3
#DEFINE DATETIME  4

/*/{Protheus.doc} User Function L12E02
    Planilha com todos os cadastros de Produtos
    @type  Function
    @author Vinicius Silva
    @since 19/04/2023
/*/
User Function L12E02()
    Local cPath       := "C:\Users\TOTVS\Desktop\listas\L12\" 
    Local cArq        := "L12E02.xls"
    Local cDados       := ConsSql()

    Private oExcel      := FwMsExcelEx():New()   
    Private cWorkSheet  := "Produtos"
    Private cTable      := "Produtos Cadastrados"

    //? Coloca o nome na planilha 
    oExcel:AddWorkSheet(cWorkSheet)

    //? Adiciona uma grade na planilha
    oExcel:AddTable(cWorkSheet, cTable) 

    //? Adiciona as colunas na planilha
    oExcel:AddColumn(cWorkSheet, cTable, "Codigo"    , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Desc."     , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Tipo"      , CENTER, GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "UM"        , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Preco"     , LEFT  , MONETARIO)     

    //! Estilizações

    //? Linhas da Coluna
    oExcel:SetLineFont("Arial")
    oExcel:SetLineSizeFont(10)
    oExcel:SetLineBgColor("#FFD3B0")

    oExcel:Set2LineFont("Arial")
    oExcel:Set2LineSizeFont(10)
    oExcel:Set2LineBgColor("#FFF9DE")

    //? Títulos da coluna
    oExcel:SetHeaderFont("Arial")
    oExcel:SetHeaderSizeFont(14)
    oExcel:SetHeaderBold(.T.)
    oExcel:SetBgColorHeader("#FF6969")
    oExcel:SetFrColorHeader("#A6D0DD")

    Info(cDados)

    oExcel:Activate() 
    oExcel:GetXMLFile(cPath + cArq)

    //? Verifica se tem o excel
    if ApOleClien("MsExcel")
        oExec := MsExcel():New()
        oExec:WorkBooks:Open(cPath + cArq)
        oExec:SetVisible(.T.)
        oExec:Destroy()
    else
        FwAlertError("Excel não encontrado no Windows", "Excel não encontrado!")
    endif

    oExcel:DeActivate()
Return 

Static Function ConsSql()
    Local aArea  := GetArea()
    Local cAlias := GetNextAlias()
    Local cQuery := ""

    cQuery += "SELECT B1_COD, B1_DESC, B1_TIPO, B1_UM, B1_PRV1, R_E_C_D_E_L_" + CRLF
	cQuery += "FROM  SB1990" + CRLF

    PREPARE ENVIRONMENT EMPRESA "99" FILIAL "01" TABLES "SB1" MODULO "COM"
    TCQUERY cQuery ALIAS &(cAlias) NEW 

    (cAlias)->(DbGoTop())

    if (cAlias)->(EOF()) 
        cAlias := ""
    end

    RestArea(aArea)
Return cAlias

//? Função p/ preencher as linhas da tabela com as informações dos fornecedores
Static Function Info(cDados)
    Local cCod, cDesc, cTipo, cUM, cPreco, cDel

    DbSelectArea(cDados)

    (cDados)->(DbGoTop())
    while (cDados) -> (!EOF())
        cCod   := (cDados)->(B1_COD)
        cDesc  := (cDados)->(B1_DESC)
        cTipo  := (cDados)->(B1_TIPO)
        cUM    := (cDados)->(B1_UM)
        cPreco := (cDados)->(B1_PRV1)
        cDel   := (cDados)->(R_E_C_D_E_L_)
        
        if  cDel <> 0
            //? Estilização caso o registro foi deletado
            oExcel:SetCelFrColor("#FFF9DE") 
            oExcel:SetCelBgColor("#FF0000") 

            oExcel:AddRow(cWorkSheet, cTable, {AllTrim(cCod), AllTrim(cDesc), AllTrim(cTipo), AllTrim(cUM), "R$ " + AllTrim(Str(cPreco,,2))}, {1,2,3,4,5}) 
        else
            oExcel:AddRow(cWorkSheet, cTable, {AllTrim(cCod), AllTrim(cDesc), AllTrim(cTipo), AllTrim(cUM), "R$ " + AllTrim(Str(cPreco,,2))})
        endif

       (cDados)->(DbSkip())
    end

    (cDados)->(DbCloseArea())
Return 
