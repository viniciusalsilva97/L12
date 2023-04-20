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

/*/{Protheus.doc} User Function L12E03
    Planilha com todos os cursos e seus respectivos alunos
    @type  Function
    @author Vinicius Silva
    @since 19/04/2023
/*/
User Function L12E03()
    Local cPath       := "C:\Users\TOTVS\Desktop\listas\L12\" 
    Local cArq        := "L12E03.xls"
    Local cDados      := ConsSql()
    Local cWorkSheet  := "" 
    Local cTable      := "Alunos Cadastrados"
    Local cCod, cNome, nIdade, cNomeCurso 
    Private oExcel      := FwMsExcelEx():New()   

    DbSelectArea(cDados)

    (cDados)->(DbGoTop())
    while (cDados)->(!EOF())
        cCod       := (cDados)->(ZZB_COD)
        cNome      := (cDados)->(ZZB_NOME)
        nIdade     := (cDados)->(ZZS_IDADE)
        cNomeCurso := (cDados)->(ZZC_NOME)

        if cWorkSheet != cNomeCurso

            cWorkSheet := cNomeCurso
            oExcel:AddWorkSheet(cWorkSheet)
            oExcel:AddTable(cWorkSheet, cTable) 

            oExcel:AddColumn(cWorkSheet, cTable, "Codigo" , LEFT  , GERAL)     
            oExcel:AddColumn(cWorkSheet, cTable, "Nome"   , LEFT  , GERAL)     
            oExcel:AddColumn(cWorkSheet, cTable, "Idade"  , CENTER, NUMERO)   
        endif

        Estiliza()

        oExcel:AddRow(cWorkSheet, cTable, {AllTrim(cCod), AllTrim(cNome), cValToChar(nIdade)})
        (cDados)->(DbSkip())    
    end
    
    (cDados)->(DbCloseArea())

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

//! Consulta SQL
Static Function ConsSql()
    Local aArea  := GetArea()
    Local cAlias := GetNextAlias()
    Local cQuery := ""

    cQuery += "SELECT ZZB_COD, ZZB_NOME, ZZC_NOME, ZZC_COD, ZZS_IDADE, ZZS_COD" + CRLF
	cQuery += "FROM  ZZB990" + CRLF
    cQuery += "INNER JOIN ZZC990 ON ZZB_CURSO = ZZC_COD AND ZZC990.D_E_L_E_T_ = ' '" + CRLF
    cQuery += "RIGHT OUTER JOIN ZZS990 ON ZZB_COD = ZZS_COD AND ZZS990.D_E_L_E_T_ = ' '" + CRLF 
    cQuery += "WHERE ZZB990.D_E_L_E_T_ = ' '"

    PREPARE ENVIRONMENT EMPRESA "99" FILIAL "01" TABLES "SB1" MODULO "COM"
    TCQUERY cQuery ALIAS &(cAlias) NEW 

    (cAlias)->(DbGoTop())

    if (cAlias)->(EOF()) 
        cAlias := ""
    end

    RestArea(aArea)
Return cAlias

//! Função para estilizar a tabela
Static Function Estiliza()
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
Return 
