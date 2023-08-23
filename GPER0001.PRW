/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒ¬ƒƒaƒƒƒƒƒƒƒ¥±±
±±≥ Funcao   ≥ FATR0004 ≥ Autor ≥ AndrÈ Salgado 	≥ Data ≥22/07/2021    ≥±±
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒ¥±±
±±≥Descricao ≥ Relatorio Resumo da Folha								  ≥±±
±±≥          ≥ 															  ≥±±
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¥±±
±±≥ Uso      ≥ IMPLACIL / EQUIPE RH                                       ≥±±
±±¿ƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ±±
*/

/*  ------------------------------------------------------------
    | Reestruturacao do Fonte - TPR Isack Lagares(17/08/2023)  |
    ------------------------------------------------------------
    | Motivo      | Impressao do Relatorio em outras extensoes |
    ------------------------------------------------------------
    | Modificacao | Relatorio Reestruturado para o TReport     |
    ------------------------------------------------------------
*/


/*----------------+
|  Biblioteca     |
+----------------*/ 
#INCLUDE "PROTHEUS.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "COLORS.CH"
#INCLUDE "RPTDEF.CH"


*===========================*
User Function GPSR0001()
*===========================*
    
    /*-----------------------------+
    |  Declaracao de variaveis     |
    +-----------------------------*/ 
    Local   cNome := "GPSR0001"
    Private	aRet  := { Space(06), Space(01), Space(01), Space(02) } // Variavel das Perguntas
    

    /*--------------------------+
    |  Declaracao de Objetos    |
    +--------------------------*/ 
    Private oReport  as Object 
    

    /*---------------------------------------+
    |  Pergunta para Filtro da Informacao    |
    +---------------------------------------*/ 
    If !ParamBox( {                                                                                        ;
	                {1, "Periodo de Folha ", aRet[1], "@!",,,                                  , 70, .F. },;
	                {2, "Categoria p/gerar", aRet[2], { "CLT"   , "Prolabore", "Autonomo" }, 50,   , .F. },;
	                {2, "Selecione a Folha", aRet[3], { "Aberta", "Encerrada"             }, 50,   , .F. },;
	                {1, "Informar a Filial", aRet[4], "@!",,,                                  , 70, .F. } ;    
                  }, "Parametro para Gerar Resumo da Folha", @aRet,,,,,,,, .T., .T. )
	    Return
    EndIf


    /*------------------------------------+
    |  Parametros para Filtro na Query    |
    +------------------------------------*/
    MV_PAR01 := StrZero( Val( aRet[1] ), 6 )  	
    MV_PAR02 := IIF( aRet[2] == "CLT", "M", IIF( aRet[2] == "Prolabore", "P", "A" ) ) 
    MV_PAR03 := IIF( aRet[3] == "Aberta", "SRC", "SRD" ) 
    MV_PAR04 := StrZero( Val( aRet[4] ), 2 )


    /*--------------------------------+
    |  Busca Folha Encerrada - SRD    |
    +--------------------------------*/
    IF MV_PAR03 == "SRD"

        //BUSCA OS DADOS DNA FOLHA DE PAGAMENTO - FOLHA FECHADA
        cQuery := " SELECT PERIODO, TP, RD_PD, RV_DESC,"
        cQuery += " SUM(REF) REF, SUM(VALOR) VALOR"
        cQuery += " FROM("
        cQuery += "     SELECT "
        cQuery += "     RD_DATARQ PERIODO, "
        cQuery += "     CASE WHEN RV_TIPOCOD = '1' THEN 'PROVENTO' ELSE 'DESCONTO' END TP, "
        cQuery += "     CASE WHEN LEFT(RV_CODMEMO,3)<>' ' THEN  LEFT(RV_CODMEMO,3) ELSE RD_PD END RD_PD,"
        cQuery += "     RD_HORAS REF, RD_VALOR VALOR"
        cQuery += " FROM "+RETSQLNAME("SRD")+" RD(NOLOCK)"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV (NOLOCK) ON RD_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " INNER JOIN "+RETSQLNAME("SRA")+" RA (NOLOCK) ON RD_FILIAL=RA_FILIAL AND RD_MAT=RA_MAT AND RA.D_E_L_E_T_=' ' "
        cQuery += "            AND RA_CATFUNC='"+MV_PAR02+"' " //PERGUNTA
        cQuery += " WHERE RD.D_E_L_E_T_=' '" 
        cQuery += "     AND RD_DATARQ = '"+MV_PAR01+"' " //PERGUNTA
        cQuery += "     AND RD_FILIAL = '"+MV_PAR04+"' " //PERGUNTA
        cQuery += "     AND RD_ROTEIR IN ('FOL','AUT')"
        cQuery += "     AND RV_TIPOCOD IN ('1','2') "
        cQuery += " )RETORNO"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV ON RD_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " GROUP BY TP,  RD_PD,  RV_DESC, PERIODO"
        cQuery += " ORDER BY 1,2 DESC, 3"


    /*-----------------------------+
    |  Busca Folha Aberta - SRC    |
    +-----------------------------*/
    Else
        cQuery := " SELECT PERIODO, TP, RD_PD, RV_DESC,"
        cQuery += " SUM(REF) REF, SUM(VALOR) VALOR"
        cQuery += " FROM("
        cQuery += "     SELECT "
        cQuery += "     RC_PERIODO PERIODO, "
        cQuery += "     CASE WHEN RV_TIPOCOD = '1' THEN 'PROVENTO' ELSE 'DESCONTO' END TP, "
        cQuery += "     CASE WHEN LEFT(RV_CODMEMO,3)<>' ' THEN  LEFT(RV_CODMEMO,3) ELSE RC_PD END RD_PD,"
        cQuery += "     RC_HORAS REF, RC_VALOR VALOR"
        cQuery += " FROM "+RETSQLNAME("SRC")+" RC(NOLOCK)"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV (NOLOCK) ON RC_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " INNER JOIN "+RETSQLNAME("SRA")+" RA (NOLOCK) ON RC_FILIAL=RA_FILIAL AND RC_MAT=RA_MAT AND RA.D_E_L_E_T_=' ' "
        cQuery += "            AND RA_CATFUNC='"+MV_PAR02+"' " //PERGUNTA
        cQuery += " WHERE RC.D_E_L_E_T_=' '" 
        cQuery += "     AND RC_FILIAL = '"+MV_PAR04+"' " //PERGUNTA
        cQuery += "     AND RC_ROTEIR IN ('FOL','AUT')"
        cQuery += "     AND RV_TIPOCOD IN ('1','2') "
        cQuery += " )RETORNO"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV ON RD_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " GROUP BY TP,  RD_PD,  RV_DESC, PERIODO"
        cQuery += " ORDER BY 1,2 DESC, 3"
    Endif

    MemoWrit( "C:\AA\BUSCAC6.TXT", cQuery )

    IF Select( "TRBSC5")  <> 0
	    DbSelectArea( "TRBSC5" )
	    DbCloseArea()
    ENDIF

    TCQUERY cQuery NEW ALIAS "TRBSC5"

    /*--------------------+
    |  Classe TReport     |
    +--------------------*/ 
    oReport := ReportDef( cNome )
    oReport:PrintDialog() 

Return( Nil )


*===================================*
Static Function ReportDef( cNome )
*===================================*

    /*--------------------------+
    |  Declaracao de Objetos    |
    +--------------------------*/ 
    Private oReport   as Object 
    Private oSection1 as Object 
    Private oSection2 as Object 
    Private oSection3 as Object 


    /*--------------------+
    |  Classe TReport     |
    +--------------------*/
    oReport := TReport():New( "GPSR0001", "RelatÛrio da Folha de Pagamento RH", /*cPerg*/, { | oReport | PrintReport( oReport ) }, "Impress„o do RelatÛrio da Folha de Pagamento" )
    oReport:SetPortrait()    
	oReport:SetTotalInLine( .F. )


    /*-------------------+
    |  Primeira Secao    |
    +-------------------*/
    oSection1 := TRSection():New( oReport, "FOLHA DE PAGAMENTO" )
    TRCell():New( oSection1, "CODIGO"    , "TRBSC5", "C”DIGO"         , "@!"             , 3  )
    TRCell():New( oSection1, "PROVENTO"  , "TRBSC5", "PROVENTOS"      , "@!"             , 60 )
    TRCell():New( oSection1, "VALOR"     , "TRBSC5", "VALOR PROVENTOS", "@E 9,999,999.99", 12 )
    TRCell():New( oSection1, "DESCONTO"  , "TRBSC5", "DESCONTOS"      , "@!"             , 60 )
    TRCell():New( oSection1, "VALOR_DESC", "TRBSC5", "VALOR DESCONTOS", "@E 9,999,999.99", 12 )


    /*-------------------+
    |  Segunda Secao     |
    +-------------------*/
    oSection2 := TRSection():New( oReport, "VALOR TOTAL " )
    TRCell():New( oSection2, "TOTAL_PROV", "TRBASE", "TOTAL PROVENTOS", "@E 9,999,999.99", 12 )
    TRCell():New( oSection2, "TOTAL_DESC", "TRBASE", "TOTAL DESCONTOS", "@E 9,999,999.99", 12 )
    TRCell():New( oSection2, "TOTAL_LIQ" , "TRBASE", "TOTAL LÕQUIDO"  , "@E 9,999,999.99", 12 )


    /*-------------------+
    |  Terceira Secao    |
    +-------------------*/
    oSection3 := TRSection():New( oReport, "BASE DE CALCULO - FOLHA" )
    TRCell():New( oSection3, "VERBA", "TRBASE", "VERBA", "@!"             , 30 )
    TRCell():New( oSection3, "VALOR", "TRBASE", "VALOR", "@E 9,999,999.99", 12 )

Return( oReport )



*=======================================*
Static Function PrintReport( oReport )
*=======================================*
    
    /*-----------------------------+
    |  Declaracao de variaveis     |
    +-----------------------------*/ 
    Private lImpCabec := .T.
    Private nTotDesc  := 0
    Private nTotProv  := 0 

    
    /*--------------------------+
    |  Declaracao de Objetos    |
    +--------------------------*/ 
    Private oSection1 := oReport:Section(1)
    Private oSection2 := oReport:Section(2)
    Private oSection3 := oReport:Section(3)


    /*-------------------------------------+
    |  Iniciando a Montagem do Relatorio   |
    +-------------------------------------*/ 
    DbSelectArea( "TRBSC5" )
    TRBSC5 -> ( DbGoTop() )


    cNovoCont := "S"

    oReport:SetMeter( TRBSC5 -> ( LastRec() ) )

    While !EoF() 

        If oReport:Cancel()
			Exit
		EndIf

        /*------------------------------+
        |  Inicializo a Primeira Secao  |
        +------------------------------*/
		oSection1:Init()


        oReport:IncMeter()
        IncProc( "Imprimindo RelatÛrio " + AllTrim( TRBSC5 -> RV_DESC ) )

        If TRBSC5 -> TP == "PROVENTO"

            /*---------------------------+
            |  Imprimo a Primeira Secao  |
            +---------------------------*/				
		    oSection1:Cell( "CODIGO"   ):SetValue( TRBSC5 -> RD_PD   )
            oSection1:Cell( "PROVENTO" ):SetValue( TRBSC5 -> RV_DESC )
		    oSection1:Cell( "VALOR"    ):SetValue( TRBSC5 -> VALOR   )
            oSection1:Printline()


            /*-----------------------+
            |  Variaveis Auxiliares  |
            +-----------------------*/
            nTotProv += TRBSC5 -> VALOR

        ElseIf TRBSC5 -> TP == "DESCONTO" .AND. cNovoCont == "S"
            /*---------------------------+
            |  Imprimo a Primeira Secao  |
            +---------------------------*/				
		    oSection1:Cell( "CODIGO"     ):SetValue( TRBSC5 -> RD_PD   )
            oSection1:Cell( "DESCONTO"   ):SetValue( TRBSC5 -> RV_DESC )
		    oSection1:Cell( "VALOR_DESC" ):SetValue( TRBSC5 -> VALOR   )
            oSection1:Printline()


            /*-----------------------+
            |  Variaveis Auxiliares  |
            +-----------------------*/
            nTotDesc += TRBSC5 -> VALOR
            /*-------------------------*/
            cNovoCont := "N"
        Else
            /*---------------------------+
            |  Imprimo a Primeira Secao  |
            +---------------------------*/				
		    oSection1:Cell( "CODIGO"   ):SetValue( TRBSC5 -> RD_PD   )
            oSection1:Cell( "PROVENTO" ):SetValue( TRBSC5 -> RV_DESC )
		    oSection1:Cell( "VALOR"    ):SetValue( TRBSC5 -> VALOR   )
            oSection1:Printline()


            /*-----------------------+
            |  Variaveis Auxiliares  |
            +-----------------------*/
            nTotDesc += TRBSC5 -> VALOR
        EndIf
        
        cNovoCont := "S"

	    TRBSC5 -> ( DbSkip() )
    EndDo


    
    /*------------------------------+
    |  Inicializo a Terceira Secao  |
    +------------------------------*/
    oSection2:Init()


    /*---------------------------------------------+
    |  Total dos Proventos / Descontos / Liquidos  |
    +---------------------------------------------*/
    oSection2:Cell( "TOTAL_PROV" ):SetValue( nTotProv )
    oSection2:Cell( "TOTAL_DESC" ):SetValue( nTotDesc )
    oSection2:Cell( "TOTAL_LIQ"  ):SetValue( nTotProv - nTotDesc )
    oSection2:Printline()


    /*--------------------------------+
    |  Busca Folha Encerrada - SRD    |
    +--------------------------------*/
    IF MV_PAR03 == "SRD"

        //Base de Calculo
        cQuery := " SELECT PERIODO, TP, RD_PD, SUM(VALOR) VALOR"
        cQuery += " FROM("
        cQuery += " SELECT RD_DATARQ PERIODO, 'BASES' TP,"
        cQuery += " CASE "
        cQuery += " WHEN RD_PD IN('B00','B01','B04') THEN '1 - BASE INSS'"
        cQuery += " WHEN RD_PD IN('C58') THEN '2 - BASE IRRF'"
        cQuery += " WHEN RD_PD IN('726') THEN '3 - VALOR IRF'"
        cQuery += " WHEN RD_PD IN('D17','D43') THEN '4 - BASE FGTS'"
        cQuery += " WHEN RD_PD IN('D18','D45') THEN '5 -VALOR FGTS'"
        cQuery += " WHEN RD_PD IN('700','702','706') THEN 'GPS - SEGURADO'"
        cQuery += " WHEN RD_PD IN('B87') THEN 'GPS - EMPRESA'"
        cQuery += " WHEN RD_PD IN('B90') THEN 'GPS - TERCEIROS'"
        cQuery += " WHEN RD_PD IN('B93') THEN 'GPS - AC.TRABALHO'"
        cQuery += " ELSE ' ' END RD_PD,"

        cQuery += " CASE "
        cQuery += " WHEN RD_PD IN('B00','B01','B04') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('C58') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('726') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('D17','D43') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('D18','D45') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('700','702','706') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('B87') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('B90') THEN RD_VALOR"
        cQuery += " WHEN RD_PD IN('B93') THEN RD_VALOR"
        cQuery += " ELSE 0 END VALOR"

        cQuery += " FROM "+RETSQLNAME("SRD")+" RD(NOLOCK)"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV (NOLOCK) ON RD_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " INNER JOIN "+RETSQLNAME("SRA")+" RA (NOLOCK) ON RD_FILIAL=RA_FILIAL AND RD_MAT=RA_MAT AND RA.D_E_L_E_T_=' '"
        cQuery += " AND RA_CATFUNC = '"+MV_PAR02+"' "

        cQuery += " WHERE RD.D_E_L_E_T_=' '" 
        cQuery += " AND RD_DATARQ = '"+MV_PAR01+"' " //PERGUNTA
        cQuery += " AND RD_FILIAL = '"+MV_PAR04+"' " //PERGUNTA
        cQuery += " AND RD_ROTEIR IN ('FOL','AUT')"
        cQuery += " AND RV_TIPOCOD IN ('2', '3','4')"

        cQuery += " )RETORNO"
        cQuery += " WHERE RD_PD<>' ' "
        cQuery += " GROUP BY TP, RD_PD, PERIODO"
        cQuery += " ORDER BY 1,2,3"


    /*-----------------------------+
    |  Busca Folha Aberta - SRC    |
    +-----------------------------*/
    Else

        //Base de Calculo
        cQuery := " SELECT PERIODO, TP, RD_PD, SUM(VALOR) VALOR"
        cQuery += " FROM("
        cQuery += " SELECT RC_PERIODO PERIODO, 'BASES' TP,"
        cQuery += " CASE "
        cQuery += " WHEN RC_PD IN('B00','B01','B04') THEN '1 - BASE INSS'"
        cQuery += " WHEN RC_PD IN('C58') THEN '2 - BASE IRRF'"
        cQuery += " WHEN RC_PD IN('726') THEN '3 - VALOR IRF'"
        cQuery += " WHEN RC_PD IN('D17','D43') THEN '4 - BASE FGTS'"
        cQuery += " WHEN RC_PD IN('D18','D45') THEN '5 -VALOR FGTS'"
        cQuery += " WHEN RC_PD IN('700','702','706') THEN 'GPS - SEGURADO'"
        cQuery += " WHEN RC_PD IN('B87') THEN 'GPS - EMPRESA'"
        cQuery += " WHEN RC_PD IN('B90') THEN 'GPS - TERCEIROS'"
        cQuery += " WHEN RC_PD IN('B93') THEN 'GPS - AC.TRABALHO'"
        cQuery += " ELSE ' ' END RD_PD,"

        cQuery += " CASE "
        cQuery += " WHEN RC_PD IN('B00','B01','B04') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('C58') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('726') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('D17','D43') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('D18','D45') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('700','702','706') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('B87') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('B90') THEN RC_VALOR"
        cQuery += " WHEN RC_PD IN('B93') THEN RC_VALOR"
        cQuery += " ELSE 0 END VALOR"

        cQuery += " FROM "+RETSQLNAME("SRC")+" RC(NOLOCK)"
        cQuery += " INNER JOIN "+RETSQLNAME("SRV")+" RV (NOLOCK) ON RC_PD=RV_COD AND RV.D_E_L_E_T_=' ' "
        cQuery += " INNER JOIN "+RETSQLNAME("SRA")+" RA (NOLOCK) ON RC_FILIAL=RA_FILIAL AND RC_MAT=RA_MAT AND RA.D_E_L_E_T_=' '"
        cQuery += " AND RA_CATFUNC = '"+MV_PAR02+"' "

        cQuery += " WHERE RC.D_E_L_E_T_=' '" 
        cQuery += " AND RC_FILIAL = '"+MV_PAR04+"' " //PERGUNTA
        cQuery += " AND RC_ROTEIR IN ('FOL','AUT')"
        cQuery += " AND RV_TIPOCOD IN ('2', '3','4')"

        cQuery += " )RETORNO"
        cQuery += " WHERE RD_PD<>' ' "
        cQuery += " GROUP BY TP, RD_PD, PERIODO"
        cQuery += " ORDER BY 1,2,3"
    EndIf 

    IF Select( "TRBASE" ) <> 0
	    DbSelectArea( "TRBASE" )
	    DbCloseArea()
    EndIf

    TCQUERY cQuery NEW Alias "TRBASE"

    DbSelectArea( "TRBASE" )
    TRBASE -> ( DbGoTop() )


    While !EoF() 

        If oReport:Cancel()
			Exit
		EndIf
        
        /*------------------------------+
        |  Inicializo a Terceira Secao  |
        +------------------------------*/
	    oSection3:Init()


        oSection3:Cell( "VERBA" ):SetValue( TRBASE -> RD_PD )
        oSection3:Cell( "VALOR" ):SetValue( TRBASE -> VALOR )
        oSection3:Printline()

	    TRBASE -> ( DbSkip() )
    EndDo


    /*----------------------+
    |  Finalizo as Secoes   |
    +----------------------*/
	oSection1:Finish()
    oSection2:Finish()
    oSection3:Finish()

Return( Nil )
