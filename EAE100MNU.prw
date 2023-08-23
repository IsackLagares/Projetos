/*
+----------------------------------------------------------------------------+
!                         FICHA TECNICA DO PROGRAMA                          !
+----------------------------------------------------------------------------+
!Modulo            ! Faturamento                                             !
+------------------+---------------------------------------------------------+
!Nome              ! EAE100MNU()                                             !
+------------------+---------------------------------------------------------+
!                  ! 1. Gerar Relatorio INVOICE                              !
!Descricao         !                                                         !
!                  ! 2. Gerar Relatorio PACKING LIST                         !
+------------------+---------------------------------------------------------+
!Autor             ! TPR System - Isack Lagares                      	     !
+------------------+---------------------------------------------------------+
!Cliente           ! Kem Parts                                               !
+------------------+---------------------------------------------------------+
!Data de Criacao   ! 30/06/2023                                              !
+------------------+---------------------------------------------------------+
*/

/*--------------------------+
| Bibliotecas               |
+--------------------------*/
#INCLUDE "RWMAKE.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TOTVS.CH"
#INCLUDE 'FONT.CH'
#INCLUDE "RPTDEF.CH"
#INCLUDE "FWPRINTSETUP.CH"


/*-------------------+
| Declarando a Cor   |
+-------------------*/
Static nCorC := RGB( 255, 247, 0 )


/*-------------------+
| Quebra de Linha    |
+-------------------*/
#DEFINE ENTER ( CHR(13) + CHR(10) )


*=========================*
User Function EAE100MNU()
*=========================*

    /*------------------------------+
    |  Declaracao de Variaveis      |
    +------------------------------*/
    Local aRotina := {}


    /*--------------------------------------+
    |  Incluindo Rotina em Outras Acoes     |
    +--------------------------------------*/
    aAdd( aRotina, { "Impressão Invoice / Packing List", "U_GERAREL()", 0, 8 } )


Return( aRotina )


*=======================*
User Function GERAREL()
*=======================*

    /*------------------------------+
    |  Declaracao de Variaveis      |
    +------------------------------*/
    Local  aArea     := GetArea()         // Dados dos Alias abertos em memoria
    Local  aParamBox := {} 
    Local  lErro     := .F. 
    Local  cProcesso := EEC -> EEC_PREEMB // Auxiliar Numero do Processo 
    Static nLineEnd  := 830               // Quantidade de Linhas por Pagina
    Static nLine     := 215               // Linha de Partida( a partir do layout construido ) 


    /*-------------------------------------+
    | Declarando Parametros para Consulta  |
    +-------------------------------------*/
    aAdd( aParamBox, { 1, "Código Porcesso Exportação", Space(20),,, "EE7",, 60 , .T. } )

    /*-------------------------------------------------------------+
    | Enquanto Documento nao Existir Retorna a Tela de Parametros  |
    +-------------------------------------------------------------*/
    If Empty( cProcesso )
        While lErro == .F.
            /*------------------------------+
            | Exibindo Tela de Parametros   |
            +------------------------------*/
            If ParamBox( aParamBox, "Parâmetros Relatórios de Exportação" )  
                /*----------------------------------+
                | Validacao se Documento Existe     |
                +----------------------------------*/
                DbSelectArea( "EE7" ) // Seleciona a Tabela de Cabecalhos da NF
                /*-----------------------------------------------------------------------------------------------------------*/
                EE7 -> ( DbSetOrder( 1 ) ) // Indicando Indice( EE7_FILIAL + EE7_PEDIDO ) 
                /*-----------------------------------------------------------------------------------------------------------*/
            EE7 -> ( DbGoTop() ) // Posiciona no Primeiro Registro
                /*-----------------------------------------------------------------------------------------------------------*/
                // Valida se o Documento Existe - Se nao Existir Exibi uma Mensagem de Atencao e Retorna a Tela de Parametros
                If EE7 -> ( DbSeek( FWxFilial( "EE7" ) + cProcesso ) )
                    lErro := .T. // Retorna Verdadeiro
                    /*-------------------------------------------*/
                    EE7 -> ( DbCloseArea() ) // Fecha a Tabela


                    /*--------------------------------------+
                    | Tela de Processamento e Finalizacao   |
                    +--------------------------------------*/
                    MsgRun( "Gerando Relatórios..."   , "Invoice & Packing List", { || GERINVO() } )
                    FWAlertSuccess( "Relatório Concluído", "Invoice & Packing List" )
                Else
                    /*---------------+
                    | Tela de Erro   |
                    +---------------*/
                    FWAlertWarning( "Documento não encontrado", "Atenção" ) 
                EndIf
            Else 
                /*-----------------+
                | Botao Cancelar   |
                +-----------------*/
                Exit
            EndIf
        EndDo
    Else    
        /*--------------------------------------+
        | Tela de Processamento e Finalizacao   |
        +--------------------------------------*/
        MsgRun( "Gerando Relatórios..."   , "Invoice & Packing List", { || GERINVO() } )
        FWAlertSuccess( "Relatório Concluído", "Invoice & Packing List" )
    EndIf


    /*----------------------------+
    | Restaura dados armazenados  |
    +----------------------------*/ 
    RestArea( aArea )
Return( Nil )


*===========================*
Static Function GERINVO() 
*===========================*

    /*-------------------------------+
    |  Declaracao de Variaveis       |
    +-------------------------------*/
    Local  aArea        := GetArea()                                                // Dados dos Alias abertos em memoria
    Static cDiretorio   := ""                                                       // Diretorio para salvar relatorio
    Local  cTitulo      := "Selecione a Pasta para Salvar o Relatório"              // Titulo da Janela de Diretorio 
    Static cArquivo     := ""                                                       // Arquivo PDF
    Local  cProcesso    := EEC -> EEC_PREEMB                                        // Auxiliar Numero do Processo 
    Local  nLine        := 290                                                      // Linha de Partida( a partir do layout construido )
    Static cQuery       := ""                                                       // Consulta SQL Cabecalho e Itens da NF
    Local  nValTotal    := 0                                                        // Auxiliar da Soma dos Valores dos Itens
    Local  nPeso        := 0                                                        // Auxiliar para Valor Total do Peso
    Local  cPacking     := ""                                                       // Auxiliar para Peso da Embalagem 
    Local  nPrcTotal    := 0                                                        // Auxiliar para Valor Final do Item
    Static aHeader      := {}                                                       // Dados Cliente/Pedido
    Static aDadosEmp    := {}                                                       // Dados da Empresa
    Static aDadosBanco  := {}                                                       // Dados dos Bancos 
    Static cLogo        := GetSrvProfString( "Startpath", "" ) + "logokemparts.png" // Diretorio da logo


    /*-------------------------------+
    |  Declaracao de Objetos         |
    +-------------------------------*/
    Local  oInvoice    // Relatorio Invoice
    Static oFontBold07 // Fonte 
    Static oFontBold14 // Fonte
    Static oFontLine   // Fonte 
    Static oFont       // Fonte


    /*-------------------------------------------+
    |  Armazenando Informacoes da Empresa        |
    +-------------------------------------------*/
    aDadosEmp := FWSM0Util():GetSM0Data()


    /*----------------+
    | SELECT          |
    +----------------*/
    cQuery := "SELECT * " + ENTER
    

    /*----------------+
    | FROM            |
    +----------------*/
    cQuery += "FROM"                                   + ENTER 
    cQuery += + RetSqlName( "EE7" ) + " EE7 (NOLOCK) " + ENTER


    /*----------------+
    | INNER JOIN      |
    +----------------*/
    cQuery += "INNER JOIN " + RetSqlName( "EE8" ) + " (NOLOCK) AS EE8 ON EE7_PEDIDO = EE8_PEDIDO " + ENTER
    

    /*----------------+
    | WHERE           |
    +----------------*/
    cQuery += "WHERE "                               + ENTER 
    cQuery += "EE7_PEDIDO = '"+ cProcesso +"' AND "  + ENTER
    cQuery += "EE7.D_E_L_E_T_ = '' "                 + ENTER


    /*------------------------------------------+ 
    | Gera um Alias Temporario                  |
    +------------------------------------------*/
    If Select( "TMP" ) > 0 // Valida se o Alias esta aberto
        TMP -> ( DbCloseArea() )
    EndIf
    /*------------------------------------------------------------------*/
    TCQuery cQuery New Alias "TMP" //  Gera um Alias a partir da Query
    /*------------------------------------------------------------------*/
    DbSelectArea( "TMP" ) // Seleciona o Alias Gerado


    /*-----------------------------+
    |  Armazenando Informacoes     |
    +-----------------------------*/
    cPacking := AllTrim( TMP -> EE8_EMBAL1 ) // Auxiliar para Packing( EE5_PESO )
    /*------------------------------------------------------------------------------------*/
    aAdd( aHeader, { TMP -> EE7_PEDIDO                                                          , ; // 01 - Codigo Processo Exportacao 
		     TMP -> EE7_DTPROC                                                          , ; // 02 - Data de Processo
                     TMP -> EE7_MOEDA                                                           , ; // 03 - Moeda 
                     AllTrim( TMP -> EE7_RESPON )                                               , ; // 04 - Contato Fornecedor
                     AllTrim( TMP -> EE7_REFIMP )                                               , ; // 05 - Referencia do Importador
		     AllTrim( TMP -> EE7_ENDIMP )                                               , ; // 06 - Endereco Importador
	             Posicione( "SY9", 2, FWxFilial( "SY9" ) + TMP -> EE7_DEST  , "Y9_CIDADE"  ), ; // 07 - Cidade Desembarque
                     Posicione( "SY9", 2, FWxFilial( "SY9" ) + TMP -> EE7_DEST  , "Y9_SIGLA"   ), ; // 08 - Sigla Embarque
                     Posicione( "SY9", 2, FWxFilial( "SY9" ) + TMP -> EE7_ORIGEM, "Y9_CIDADE"  ), ; // 09 - Cidade Embarque  
                     Posicione( "SY9", 2, FWxFilial( "SY9" ) + TMP -> EE7_ORIGEM, "Y9_ESTADO"  ), ; // 10 - Estado Desembarque
                     Posicione( "SY6", 1, FWxFilial( "SY6" ) + TMP -> EE7_CONDPA, "Y6_XDESCI"  ), ; // 11 - Condicao Pagamento
                     Posicione( "SYJ", 1, FWxFilial( "SYJ" ) + TMP -> EE7_INCOTE, "YJ_DESCR"   ), ; // 12 - Condicao Venda(INCOTERM)
                     AllTrim( TMP -> EE7_IMPODE )                                               , ; // 13 - Nome Importador
                     Posicione( "SYA", 1, FWxFilial( "SYA" ) + TMP -> EE7_PAISET, "YA_DESCR"   ), ; // 14 - Pais de Origem
                     Posicione( "EE5", 1, FWxFilial( "EE5" ) + cPacking         , "EE5_PESO"   ), ; // 15 - Peso Embalagem
                     Posicione( "SYQ", 1, FWxFilial( "SYQ" ) + TMP -> EE7_VIA   , "YQ_DESCR"   ), ; // 16 - Via de Transporte 
                     Posicione( "SA2", 1, FWxFilial( "SA2" ) + TMP -> EE7_FORN  , "A2_NOME"    ), ; // 17 - Nome do Fornecedor
                     Posicione( "SA2", 1, FWxFilial( "SA2" ) + TMP -> EE7_FORN  , "A2_DDD"     ), ; // 18 - DDD Fornecedor
                     Posicione( "SA2", 1, FWxFilial( "SA2" ) + TMP -> EE7_FORN  , "A2_TEL"     ), ; // 19 - Telefone Forncedor 
                     Posicione( "SA2", 1, FWxFilial( "SA2" ) + TMP -> EE7_FORN  , "A2_EMAIL"   ), ; // 20 - Email Forncedor
                    } )
    /*-----------------------------------------------------------------------------------------------------------------*/
    aAdd( aDadosBanco, { Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_NOME"    ), ; // 01 - Nome Banco 
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_XIBAN"   ), ; // 02 - Iban Banco 
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_XSWIFT"  ), ; // 03 - Codigo Swift 
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_NUMCON"  ), ; // 04 - Numero da Conta
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_AGENCIA" ), ; // 05 - Numero da Agencia
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_END"     ), ; // 06 - Endereco Banco
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_MUN"     ), ; // 07 - Municipio Banco
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_EST"     ), ; // 08 - Estado Banco
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBANCO, "A6_CEP"     ), ; // 09 - CEP Banco
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_NOME"    ), ; // 10 - Nome Banco Cliente
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_XSWIFT"  ), ; // 11 - Codigo Swift Cliente
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_NUMCON"  ), ; // 12 - Numero da Conta Cliente
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_XCODLIM" ), ; // 13 - Clearing Code Cliente
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_MUN"     ), ; // 14 - Estado Banco Cliente
                         Posicione( "SA6", 1, FWxFilial( "SA6" ) + TMP -> EE7_XBACLI, "A6_PAISBCO" )  ; // 15 - Pais Banco Cliente
                        } )


    /*-------------------------------------------+
    |  Definindo Dados da Impressao              |
    +-------------------------------------------*/
    oBrush := TBrush():New( , nCorC ) // Definindo a Cor da Tinta
    /*-----------------------------------------------------------------------------------------------------------------------*/
    cArquivo := "Invoice_" + AllTrim( aHeader[1][1] ) + dToS( Date() ) + ".pdf" // Nome do Arquivo
    /*-----------------------------------------------------------------------------------------------------------------------*/
    cDiretorio := cGetFile( , cTitulo,,, .T., GETF_LOCALHARD + GETF_RETDIRECTORY, .T., .T. ) // Seleciona o Diretorio
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice := FwMsPrinter():New( cArquivo, IMP_PDF, .F., cDiretorio, .T.,, @oInvoice,,,,, .T. ) // Gerar Relatorio Grafico
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice:cPathPDF := cDiretorio // Armazena Diretorio Selecionado
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice:SetResolution( 72 )
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice:SetLandscape()
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice:SetPaperSize( DMPAPER_A4 ) // Tamanho da Folha 
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oInvoice:SetMargin( 0, 0, 0, 0 )
    /*-----------------------------------------------------------------------------------------------------------------------*/


    /*-------------------------------------------+
    |  Definindo fontes para Impressao           |
    +-------------------------------------------*/
    oFontBold07 := TFont():New( "Arial" ,, 07,, .T.,,,,, .F., .F. ) // Com Negrito - Tamanho 14
    /*---------------------------------------------------------------------------------------------*/
    oFontBold14 := TFont():New( "Arial" ,, 14,, .T.,,,,, .F., .F. ) // Com Negrito - Tamanho 07
    /*---------------------------------------------------------------------------------------------*/
    oFont       := TFont():New( "Arial" ,, 07,, .F.,,,,, .F., .F. ) // Sem Negrito - Tamanho 07
    /*---------------------------------------------------------------------------------------------*/
    oFontLine   := TFont():New( "Arial" ,, 07,, .T.,,,,, .T., .F. ) // Sublinhado  - Tamanho 07    


    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Box )     |
    +-------------------------------------------*/
    oInvoice:Box( 030, 030, 085, 570 )
    oInvoice:Box( 085, 030, 280, 300 )
    oInvoice:Box( 085, 300, 280, 570 )
    oInvoice:Box( 280, 030, 410, 570 )
    oInvoice:Box( 410, 030, 445, 570 )
    oInvoice:Box( 445, 030, 525, 570 )
    oInvoice:Box( 525, 030, 540, 570 )
    oInvoice:Box( 540, 030, 735, 300 )
    oInvoice:Box( 540, 300, 735, 570 )
    oInvoice:Box( 735, 030, 757, 570 )


    /*-----------------------+
    |  Impressao da Logo     |
    +-----------------------*/
    oInvoice:SayBitMap( 028, 035, cLogo, 150, 60 )
    

    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Say )     |
    +-------------------------------------------*/
    oInvoice:Say( 060, 250, "COMMERCIAL INVOICE"                                                      , oFontBold14 )
    oInvoice:Say( 100, 065, AllTrim( aDadosEmp[5][2] )                                                , oFontBold07 )
    oInvoice:Say( 110, 035, "SELLER:"                                                                 , oFontBold07 )
    oInvoice:Say( 110, 065, AllTrim( aDadosEmp[22][2] )                                               , oFontBold07 )
    oInvoice:Say( 120, 065, AllTrim( aDadosEmp[17][2] ) + " - " + AllTrim( aDadosEmp[25][2] )  + " - ";
                            + AllTrim( aDadosEmp[19][2] )                                             ;
                            + " - Brasil - " + Transform( aDadosEmp[27][2], "@R 99999-999" )          , oFontBold07 )
    oInvoice:Say( 130, 065, "CNPJ: " + Transform( aDadosEmp[11][2], "@R 99.999.999/9999-99" )         , oFontBold07 )
    oInvoice:Say( 170, 035, "BUYER: " + aHeader[1][13]                                                , oFontBold07 )
    oInvoice:Say( 180, 062, aHeader[1][6]                                                             , oFontBold07 )
    //oInvoice:Say( 190, 067, aHeader[1][]                                                              , oFontBold07 )
    oInvoice:Say( 230, 035, "Port of loading: " + AllTrim( aHeader[1][9] ) + ", " + aHeader[1][10]    , oFont       )
    oInvoice:Say( 240, 035, "Port of dischargerer: " + AllTrim( aHeader[1][7] ) + ", " + aHeader[1][8], oFont       )
    oInvoice:Say( 250, 035, "Incoterms 2020: " + aHeader[1][12]                                       , oFont       )
    oInvoice:Say( 260, 035, "Terms of payment: " + aHeader[1][11]                                     , oFont       )
    oInvoice:Say( 270, 035, "Notify: " + aHeader[1][13]                                               , oFont       )
    /*------------------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 100, 390   , "REFERENCE INFORMATION"                                                      , oFontBold07 )
    oInvoice:Say( 120, 346   , "Invoice N°: " + aHeader[1][1]                                               , oFont       )
    oInvoice:Say( 130, 361   , "Date: " + Transform( StoD( aHeader[1][2] ), "@E 99/99/9999" )               , oFont       )
    oInvoice:Say( 160, 390   , "OUR ACCOUNT WITH YOU "                                                      , oFontBold07 )
    oInvoice:Say( 180, 331   , "Contact person: " + aHeader[1][17]                                          , oFont       )
    oInvoice:Say( 190, 344   , "Telephone: + " + aHeader[1][18] + Transform(aHeader[1][19], "@R 9999-9999" ), oFont       )
    oInvoice:Say( 200, 355.50, "E-Mail: " + aHeader[1][20]                                                  , oFont       )
    oInvoice:Say( 230, 348   , "Currency: " + aHeader[1][3]                                                 , oFont       )
    oInvoice:Say( 240, 315   , "Customer Reference: " + aHeader[1][5]                                       , oFont       )
    oInvoice:Say( 250, 339   , "Ship by Sea: " + aHeader[1][16]                                             , oFont       )
    /*-------------------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 290, 040, "Item"                       , oFontLine   )
    oInvoice:Say( 290, 070, "Description"                , oFontLine   )
    oInvoice:Say( 290, 280, "KG "                        , oFontLine   )
    oInvoice:Say( 290, 315, "Unit "                      , oFontLine   )
    oInvoice:Say( 290, 370, "Price/Unit " + aHeader[1][3], oFontLine   )
    oInvoice:Say( 290, 520, "Total " + aHeader[1][3]     , oFontLine   )


    TMP -> ( DbGoTop() ) // Posiciona no Primeiro Registro da Tabela
    /*----------------------------------------------------------------------------*/
    While !TMP -> ( EoF() )
        oInvoice:Say( nLine + 20, 040, TMP -> EE8_SEQUEN                                  , oFontBold07 )
        oInvoice:Say( nLine + 20, 070, TMP -> EE8_XDESCI                                  , oFontBold07 )
        oInvoice:Say( nLine + 20, 260, Transform( TMP -> EE8_PSLQTO, "@E 9,999,999.99" )  , oFontBold07 )
        oInvoice:Say( nLine + 20, 317, TMP -> EE8_UNIDAD                                  , oFontBold07 )
        oInvoice:Say( nLine + 20, 368, Transform( TMP -> EE8_PRECO , "@E 99,999,999.999" ), oFontBold07 )


        /*--------------------------------------------+
        |  Armezenando dados para variaveis auxiliar  |
        +--------------------------------------------*/
        nPrcTotal := TMP -> EE7_PESLIQ * TMP -> EE8_PRECO // Valor Final do Item
        /*-------------------------------------------------------------------------*/
        nValTotal += nPrcTotal // Soma dos Valores dos Itens
        /*-------------------------------------------------------------------------*/
        nPeso += TMP -> EE8_PSLQTO


        oInvoice:Say( nLine + 20, 512, Transform( nPrcTotal, "@E 99,999,999.999" ), oFontBold07 )


        /*-----------------------+
        |  Controle de Itens     |
        +-----------------------*/
        nLine := nLine + 20 


        /*-----------------------+
        |  Proximo Registro      |
        +-----------------------*/
        TMP -> ( DbSkip() )
    EndDo 


    TMP -> ( DbGoTop() ) // Posiciona no Primeiro Registro da Tabela
    /*----------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 360, 070, "Custom Tariff No. " + Transform( TMP -> EE8_POSIPI, "@R 99 99 99 99" ), oFontBold07 )
    oInvoice:Say( 370, 070, "Country of Origin: " + aHeader[1][14]                                 , oFontBold07 )
    oInvoice:Say( 380, 070, "Packing: " + Transform( aHeader[1][15], "@E 9,999,999.99" ) + " "     ;
                                        + TMP -> EE8_UNIDAD + " Each"                              , oFontBold07 )
    oInvoice:Say( 390, 070, "Net Weight: " + Transform( TMP -> EE7_PESLIQ, "@E 9,999,999.99" )     , oFontBold07 )
    oInvoice:Say( 400, 070, "Gross Weight: " + Transform( TMP -> EE7_PESBRU, "@E 9,999,999.99" )   , oFontBold07 )
    /*----------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 430, 040, "Total Amount"                                   , oFontBold07 )
    oInvoice:Say( 430, 255, Transform( nPeso, "@E 9,999,999.99" )            , oFontBold07 )
    oInvoice:Say( 430, 310, TMP -> EE8_UNIDAD                                , oFontBold07 )
    oInvoice:Say( 430, 450, aHeader[1][3]                                    , oFontBold07 )
    oInvoice:Say( 430, 500, "$" + Transform( nValTotal, "@E 99,999,999.999" ), oFontBold07 )
    /*-------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 455, 265, "GENERAL CONDITIONS"                                                , oFontBold07 )
    oInvoice:Say( 475, 040, "The Claim is not aceptable after 30 days against receiving goods." , oFont       )
    oInvoice:Say( 485, 040, "Insurance: To be covered by buyer"                                 , oFont       )
    oInvoice:Say( 495, 040, "The exporter of the products covered by this document declares " + ; 
                            "that, except where otherwise clearly indicated, these"             , oFont       )
    oInvoice:Say( 505, 040, "products are of Brazilian preferential origin according to "     + ;
                            "rules of origin of the Generalized System of Preferences."         , oFont       )
    /*-------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 535, 270, "BANK INFORMATION", oFontBold07 )
    /*------------------------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 560, 070, "Correspondent Bank:"                                                                   , oFontLine   )
    oInvoice:Say( 570, 070, "Account with: " + AllTrim( aDadosBanco[1][10] ) + " - " + AllTrim( aDadosBanco[1][14] );
                                             + " - " + aDadosBanco[1][15]                                           , oFont       )
    oInvoice:Say( 570, 040, "Field 56"                                                                              , oFontBold07 )
    oInvoice:Say( 580, 070, "Swift Code: " + aDadosBanco[1][11]                                                     , oFont       )
    oInvoice:Say( 590, 070, "Clearing code: " + aDadosBanco[1][13]                                                  , oFont       )
    oInvoice:FillRect( { 594, 065, 602, 220 }                                                                       , oBrush      )
    oInvoice:Say( 600, 070, "Account Number: " + aDadosBanco[1][12]                                                 , oFont       )
    /*------------------------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 560, 350, "Beneficiary Bank:"                                              , oFontLine   )
    oInvoice:Say( 570, 320, "Field 57"                                                       , oFontBold07 )
    oInvoice:Say( 570, 350, "In favor of " + aDadosBanco[1][1]                               , oFont       )
    oInvoice:Say( 580, 350, "Swift Code: " + aDadosBanco[1][3]                               , oFont       )
    oInvoice:Say( 590, 350, "Bank Adress: " + AllTrim( aDadosBanco[1][6] ) + " - "           ;
                                            + AllTrim( aDadosBanco[1][8] )                   , oFont       )
    oInvoice:Say( 600, 350, "CEP " + Transform( aDadosBanco[1][9] , "@R 99999-999" )         , oFont       )
    oInvoice:Say( 620, 350, "Final Beneficiary:"                                             , oFontLine   )
    oInvoice:Say( 630, 320, "Field 58"                                                       , oFontBold07 )
    oInvoice:Say( 630, 350, "For Further credit to: " + SubStr( aDadosEmp[5][2], 1, 31 )     , oFont       )
    oInvoice:Say( 640, 350, Right( aDadosEmp[5][2], 31 )                                     , oFont       )
    oInvoice:Say( 650, 350, "Account number: " + aDadosBanco[1][4]                           , oFont       )
    oInvoice:Say( 660, 350, "Branch number: " + aDadosBanco[1][5]                            , oFont       )
    oInvoice:Say( 670, 350, "IBAN: " + aDadosBanco[1][2]                                     , oFont       )
    oInvoice:Say( 680, 350, "Address: " + AllTrim( aDadosEmp[22][2] ) + " "                  ;
                                        + AllTrim( aDadosEmp[17][2] ) + " - "                ;
                                        + AllTrim( aDadosEmp[25][2] ) + " - "                ;
                                        + AllTrim( aDadosEmp[19][2] )                        , oFont       )         
    oInvoice:Say( 690, 350, "Brasil - CEP: " + Transform( aDadosEmp[27][2], "@R 99999-999" ) , oFont       )
    oInvoice:Say( 700, 350, "CNPJ: " + Transform( aDadosEmp[11][2], "@R 99.999.999/9999-99" ), oFont       )
    /*---------------------------------------------------------------------------------------------------------------------*/
    oInvoice:Say( 745, 160, AllTrim( aDadosEmp[22][2] ) + " - " + AllTrim( aDadosEmp[17][2] ) + " / CEP ";
                            + Transform( aDadosEmp[27][2], "@R 99999-999" )  + " - "                     ;
                            + AllTrim( aDadosEmp[25][2] ) + " / " + AllTrim( aDadosEmp[19][2] ) + " / "  ;
                            + " Brasil" + " - Tel: " + Transform( aDadosEmp[7][2], "@R 99 9999-9999" )   , oFontBold07 )
    oInvoice:Say( 755, 275, "www.kemparts.com.br"                                                        , oFontBold07 )
    

    /*-------------------------------+
    |  Finalizando e Exibindo PDF    |
    +-------------------------------*/ 
    oInvoice:EndPage() // Finaliza a Pagina
    /*--------------------------------------------------------------*/
    oInvoice:Preview() // Mostra um preview do PDF

    /*-------------------------------+
    |  Funcao imprimir Packing List  |
    +-------------------------------*/ 
    GERPACK()

    /*----------------------------+
    | Restaura dados armazenados  |
    +----------------------------*/ 
    RestArea( aArea )
Return( Nil )


*===========================*
Static Function GERPACK() 
*===========================*

    /*-------------------------------+
    |  Declaracao de Variaveis       |
    +-------------------------------*/
    Local nLine   := 425 // Linha de Partida( a partir do layout construido )
    Local nPBruto := 0
    Local nPLiqui := 0 

    /*-------------------------------+
    |  Declaracao de Objetos         |
    +-------------------------------*/
    Local  oPacking // Relatorio Packing

    
    /*-------------------------------------------+
    |  Definindo Dados da Impressao              |
    +-------------------------------------------*/
    cArquivo := "PackingList_" + AllTrim( aHeader[1][1] ) + "_" + dToS( Date() ) + ".pdf" // Nome do Arquivo
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking := FwMsPrinter():New( cArquivo, IMP_PDF, .F., cDiretorio, .T.,, @oPacking,,,,, .T. ) // Gerar Relatorio Grafico
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking:cPathPDF := cDiretorio // Armazena Diretorio Selecionado
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking:SetResolution( 72 )
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking:SetLandscape()
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking:SetPaperSize( DMPAPER_A4 ) // Tamanho da Folha 
    /*-----------------------------------------------------------------------------------------------------------------------*/
    oPacking:SetMargin( 0, 0, 0, 0 )


    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Box )     |
    +-------------------------------------------*/
    oPacking:Box( 150, 030, 230, 570 )
    oPacking:Box( 230, 030, 340, 350 )
    oPacking:Box( 230, 350, 340, 570 )
    oPacking:Box( 340, 030, 390, 570 )
    oPacking:Box( 390, 030, 425, 080 )
    oPacking:Box( 390, 080, 425, 270 )
    oPacking:Box( 390, 270, 425, 470 )
    oPacking:Box( 390, 470, 425, 520 )
    oPacking:Box( 390, 520, 425, 570 )
    oPacking:Box( 425, 030, 525, 570 )


    /*-----------------------+
    |  Impressao da Logo     |
    +-----------------------*/
    oPacking:SayBitMap( 160, 045, cLogo, 150, 60 )


    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Say )     |
    +-------------------------------------------*/
    oPacking:Say( 190, 265, "PACKING LIST"                                                         , oFontBold14 )
    oPacking:Say( 238, 033, "EXPORTER"                                                             , oFontBold07 )
    oPacking:Say( 248, 033, aDadosEmp[5][2]                                                        , oFont       )
    oPacking:Say( 258, 033, "CNPJ.: " + Transform( aDadosEmp[11][2], "@R 99.999.999/9999-99" )     , oFont       )
    oPacking:Say( 268, 033, aDadosEmp[22][2]                                                       , oFont       )
    oPacking:Say( 278, 033, AllTrim( aDadosEmp[17][2] ) + "," + AllTrim( aDadosEmp[25][2] )  + "," ;
                            + AllTrim( aDadosEmp[19][2] ) + " "                                    ;
                            + Transform( aDadosEmp[27][2], "@R 99999-999" )                        , oFont       )
    oPacking:Say( 288, 033, "Brasil "                                                              , oFont       )
    oPacking:Say( 238, 355, "SOLD TO"                                                              , oFontBold07 )
    oPacking:Say( 248, 355, aHeader[1][13]                                                         , oFont       )
    oPacking:Say( 258, 355, aHeader[1][6]                                                          , oFont       )
    oPacking:Say( 318, 033, "SHIP TO"                                                              , oFontBold07 )
    oPacking:Say( 328, 033, aHeader[1][13]                                                         , oFont       )
    oPacking:Say( 338, 033, aHeader[1][6]                                                          , oFont       )
    oPacking:Say( 348, 370, "INVOICE N°: " + aHeader[1][1]                                         , oFontBold07 )
    oPacking:Say( 348, 033, "PORT OF LOADING: " + AllTrim( aHeader[1][9] ) + ", " + aHeader[1][10] , oFontBold07 )
    oPacking:Say( 358, 033, "PORT OF DISCHARGE: " + AllTrim( aHeader[1][7] ) + ", " + aHeader[1][8], oFontBold07 )
    oPacking:Say( 368, 033, "VESSEL: " + EEC -> EEC_EMBARC                                         , oFontBold07 )
    oPacking:Say( 378, 033, "YOUR ORDER: " + aHeader[1][5]                                         , oFontBold07 )
    oPacking:Say( 388, 033, "OUR ORDER: " + aHeader[1][1]                                          , oFontBold07 )
    oPacking:Say( 398, 048, "ITEM"                                                                 , oFontBold07 )
    oPacking:Say( 408, 051, "NR"                                                                   , oFontBold07 )
    oPacking:Say( 410, 165, "QTY"                                                                  , oFontBold07 )
    oPacking:Say( 410, 335, "GOODS DESCRIPTION"                                                    , oFontBold07 )
    oPacking:Say( 398, 488, "NET"                                                                  , oFontBold07 )
    oPacking:Say( 408, 483, "WEIGHT"                                                               , oFontBold07 )
    oPacking:Say( 418, 490, "KG"                                                                   , oFontBold07 )
    oPacking:Say( 398, 533, "GROSS"                                                                , oFontBold07 )
    oPacking:Say( 408, 532, "WEIGHT"                                                               , oFontBold07 )
    oPacking:Say( 418, 538, "KG"                                                                   , oFontBold07 )


    TMP -> ( DbGoTop() ) // Posiciona no Primeiro Registro da Tabela
    /*----------------------------------------------------------------------------*/
    While !TMP -> ( EoF() )
        /*----------------------------------+
        |  Imprimindo inforamcoes de itens  |
        +----------------------------------*/
        oPacking:Box( nLine, 030, nLine + 10, 080 )
        oPacking:Box( nLine, 080, nLine + 10, 270 )
        oPacking:Box( nLine, 270, nLine + 10, 470 )
        oPacking:Box( nLine, 470, nLine + 10, 520 )
        oPacking:Box( nLine, 520, nLine + 10, 570 )
    /*----------------------------------------------------------------------------*/
        oPacking:Say( nLine + 8, 045, TMP -> EE8_SEQUEN                                , oFontBold07 )
        oPacking:Say( nLine + 8, 083, EEC -> EEC_XBAG                                  , oFontBold07 )
        oPacking:Say( nLine + 8, 273, TMP -> EE8_XDESCI                                , oFontBold07 )
        oPacking:Say( nLine + 8, 473, Transform( TMP -> EE7_PESLIQ, "@E 9,999,999.99" ), oFontBold07 )
        oPacking:Say( nLine + 8, 523, Transform( TMP -> EE7_PESBRU, "@E 9,999,999.99" ), oFontBold07 )


        /*-----------------------+
        |  Controle de Itens     |
        +-----------------------*/
        nLine := nLine + 10 


        /*--------------------------------------------+
        |  Armezenando dados para variaveis auxiliar  |
        +--------------------------------------------*/
        nPBruto += TMP -> EE7_PESLIQ // Soma do Peso Bruto
        /*----------------------------------------------------------------------------*/
        nPLiqui += TMP -> EE7_PESBRU // Soma do Peso Liquido


        /*-----------------------+
        |  Proximo Registro      |
        +-----------------------*/
        TMP -> ( DbSkip() )
    EndDo 


    TMP -> ( DbGoTop() ) // Posiciona no Primeiro Registro da Tabela
    /*----------------------------------------------------------------------------------------------------------------*/
    oPacking:Box( nLine, 470, nLine + 10, 520 )
    oPacking:Box( nLine, 520, nLine + 10, 570 )
    /*----------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( nLine + 8, 473, Transform( nPBruto, "@E 9,999,999.99" ), oFontBold07 )
    oPacking:Say( nLine + 8, 523, Transform( nPLiqui, "@E 9,999,999.99" ), oFontBold07 )
    oPacking:Say( 465, 033, "MARKS:"                                     , oFontBold07 )


    TMP -> ( DbGoTop() ) // Posiciona no Primeiro Registro da Tabela
    /*----------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( 465, 085, "Custom Tariff No. " + Transform( TMP -> EE8_POSIPI, "@R 99 99 99 99" ), oFontBold07 )
    oPacking:Say( 475, 085, "Country of Origin: " + aHeader[1][14]                                 , oFontBold07 )
    oPacking:Say( 485, 085, "Packing: " + Transform( aHeader[1][15], "@E 9,999,999.99" ) + " "     ;
                                        + TMP -> EE8_UNIDAD + " Each"                              , oFontBold07 )
    oPacking:Say( 495, 085, "Net Weight: " + Transform( TMP -> EE7_PESLIQ, "@E 9,999,999.99" )     , oFontBold07 )
    oPacking:Say( 505, 085, "Gross Weight: " + Transform( TMP -> EE7_PESBRU, "@E 9,999,999.99" )   , oFontBold07 )


    /*-------------------------------+
    |  Finalizando e Exibindo PDF    |
    +-------------------------------*/ 
    oPacking:EndPage() // Finaliza a Pagina
    /*--------------------------------------------------------------*/
    oPacking:Preview() // Mostra um preview do PDF

Return( Nil )
