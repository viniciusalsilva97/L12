#INCLUDE "Totvs.ch"
#INCLUDE "Topconn.ch"
#INCLUDE "Tbiconn.ch"

//! Alinhamento
#DEFINE LEFT 1
#DEFINE CENTER 2
#DEFINE RIGHT 3

//! Formatação
#DEFINE GERAL     1
#DEFINE NUMERO    2
#DEFINE MONETARIO 3
#DEFINE DATETIME  4

/*/{Protheus.doc} User Function L12E01
    Planilha com todos os cadastros dos fornecedores
    @type  Function
    @author Vinicius Silva
    @since 19/04/2023
/*/
User Function L12E01()
    Local cPath       := "C:\Users\TOTVS\Desktop\listas\" 
    Local cArq        := "L12E01.xls"
    Local cDados       := ConsSql()

    Private oExcel      := FwMsExcelEx():New()   
    Private cWorkSheet  := "Fornecedores"
    Private cTable      := "Fornecedores Cadastrados"

    //? Coloca o nome na planilha 
    oExcel:AddWorkSheet(cWorkSheet)

    //? Adiciona uma grade na planilha
    oExcel:AddTable(cWorkSheet, cTable) 

    //? Adiciona as colunas na planilha
    oExcel:AddColumn(cWorkSheet, cTable, "Codigo"   , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Nome"     , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Loja"     , CENTER, GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "CNPJ"     , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Endereço" , LEFT  , GERAL)     
    oExcel:AddColumn(cWorkSheet, cTable, "Bairro"   , LEFT  , GERAL)
    oExcel:AddColumn(cWorkSheet, cTable, "Cidade"   , LEFT  , GERAL)
    oExcel:AddColumn(cWorkSheet, cTable, "UF"       , CENTER, GERAL)

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

    cQuery += "SELECT A2_COD, A2_NOME, A2_LOJA, A2_CGC, A2_END, A2_BAIRRO, A2_MUN, A2_EST" + CRLF
	cQuery += "FROM  SA2990" + CRLF
	cQuery += "WHERE D_E_L_E_T_= ' '"

    PREPARE ENVIRONMENT EMPRESA "99" FILIAL "01" TABLES "SA2" MODULO "COM"
    TCQUERY cQuery ALIAS &(cAlias) NEW 

    (cAlias)->(DbGoTop())

    if (cAlias)->(EOF()) 
        cAlias := ""
    end

    RestArea(aArea)
Return cAlias

//? Função p/ preencher as linhas da tabela com as informações dos fornecedores
Static Function Info(cDados)
    Local cCod, cNome, cLoja, cCNPJ, cEnd, cBair, cMun, cEst

    DbSelectArea(cDados)

    (cDados)->(DbGoTop())
    while (cDados) -> (!EOF())
        cCod  := (cDados)->(A2_COD)
        cNome := (cDados)->(A2_NOME)
        cLoja := (cDados)->(A2_LOJA)
        cCNPJ := (cDados)->(A2_CGC)
        cEnd  := (cDados)->(A2_END)
        cBair := (cDados)->(A2_BAIRRO)
        cMun  := (cDados)->(A2_MUN)
        cEst  := (cDados)->(A2_EST)

       oExcel:AddRow(cWorkSheet, cTable, {AllTrim(cCod), AllTrim(cNome), AllTrim(cLoja), AllTrim(cCNPJ), AllTrim(cEnd), AllTrim(cBair), AllTrim(cMun), AllTrim(cEst)}) 

       (cDados)->(DbSkip())
    end

    (cDados)->(DbCloseArea())
Return 
