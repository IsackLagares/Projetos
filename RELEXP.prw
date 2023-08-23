/*
+----------------------------------------------------------------------------+
!                         FICHA TECNICA DO PROGRAMA                          !
+----------------------------------------------------------------------------+
!Modulo            ! Faturamento                                             !
+------------------+---------------------------------------------------------+
!Nome              ! RELEXP()                                                !
+------------------+---------------------------------------------------------+
!                  ! 1. Gerar Relatorio INVOICE                              !
!Descricao         !                                                         !
!                  ! 2. Gerar Relatorio PACKING LIST                         !
+------------------+---------------------------------------------------------+
!Autor             ! TPR System - Isack Lagares                      	     !
+------------------+---------------------------------------------------------+
!Cliente           ! Implacil de Bortoli                                     !
+------------------+---------------------------------------------------------+
!Data de Criacao   ! 19/05/2023                                              !
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
#INCLUDE 'COLORS.CH'
#INCLUDE "RPTDEF.CH"
#INCLUDE "FWPRINTSETUP.CH"

/*-------------------+
| Declarando a Cor   |
+-------------------*/
Static nCorC := RGB( 211, 211, 211 )

/*-------------------+
| Quebra de Linha    |
+-------------------*/
#DEFINE ENTER ( CHR(13) + CHR(10) )

*=======================*
User Function RELEXP()
*=======================*

    /*------------------------------+
    |  Declaracao de Variaveis      |
    +------------------------------*/
    Local  aParamBox := {} 
    Local  lErro     := .F. 
    Static nLineEnd := 830 // Quantidade de Linhas por Pagina
    Static nLine    := 215 // Linha de Partida( a partir do layout construido ) 


    /*-------------------------------------+
    | Declarando Parametros para Consulta  |
    +-------------------------------------*/
    aAdd( aParamBox, { 1, "Documento de Saída", Space(9),,, "SF2",, 40 , .T. } )
    aAdd( aParamBox, { 1, "Série"             , Space(3),,,,,       10 , .T. } )
    aAdd( aParamBox, { 1, "Entrega Terceiros" , Space(6),,, "SA1",, 40 , .F. } )
    aAdd( aParamBox, { 1, "Loja"              , Space(2),,,,,       10 , .F. } )


    /*-------------------------------------------------------------+
    | Enquanto Documento nao Existir Retorna a Tela de Parametros  |
    +-------------------------------------------------------------*/
    While lErro == .F.
        /*------------------------------+
        | Exibindo Tela de Parametros   |
        +------------------------------*/
        If ParamBox( aParamBox, "Parâmetros Invoice e Packing List" )  

            /*----------------------------------+
            | Validacao se Documento Existe     |
            +----------------------------------*/
            DbSelectArea( "SF2" ) // Seleciona a Tabela de Cabecalhos da NF
            /*-----------------------------------------------------------------------------------------------------------*/
            SF2 -> ( DbSetOrder( 1 ) ) // Indicando Indice( F2_FILIAL + F2_DOC + F2_SERIE ) 
            /*-----------------------------------------------------------------------------------------------------------*/
            SF2 -> ( DbGoTop() ) // Posiciona no Primeiro Registro
            /*-----------------------------------------------------------------------------------------------------------*/
            // Valida se o Documento Existe - Se nao Existir Exibi uma Mensagem de Atencao e Retorna a Tela de Parametros
            If SF2 -> ( DbSeek( FWxFilial( "SF2" ) + MV_PAR01 + AllTrim( MV_PAR02 ) ) )
                lErro := .T. // Retorna Verdadeiro
                /*-------------------------------------------*/
                SF2 -> ( DbCloseArea() ) // Fecha a Tabela


                /*--------------------------------------+
                | Tela de Processamento e Finalizacao   |
                +--------------------------------------*/
                MsgRun( "Gerando Relatório..."   , "Relatório Invoice", { || GERINVO() } )
                MsgRun( "Gerando Packing List...", "Packing List"     , { || GERPACK() } )
                FWAlertSuccess( "Relatórios Concluídos", "Invoice e Packing List" )
            Else
                /*---------------+
                | Tela de Erro   |
                +---------------*/
                FWAlertWarning( "Documento de Saída não encontrado", "Atenção" ) 
            EndIf
        Else 
            /*-----------------+
            | Botao Cancelar   |
            +-----------------*/
            Exit
        EndIf
    EndDo
Return( Nil )


*===========================*
Static Function GERINVO() 
*===========================*

    /*-------------------------------+
    |  Declaracao de Variaveis       |
    +-------------------------------*/
    Local  aArea        := GetArea()                                        // Dados dos Alias abertos em memoria
    Static cDiretorio   := ""                                               // Diretorio para salvar relatorio
    Local  cTitulo      := "Selecione a Pasta para Salvar o Relatório"      // Titulo da Janela de Diretorio 
    Static cLogo        := GetSrvProfString( "Startpath", "" ) + "LOGO.BMP" // Diretorio da logo
    Local  cQuery       := ""                                               // Consulta SQL Cabecalho e Itens da NF
    Local  cQry         := ""                                               // Consulta SQL para o Banco
    Static aHeader      := {}                                               // Dados Cliente/Pedido
    Local  nFrete       := 0                                                // Valor Frete( F2_FRETE )
    Local  nValMerc     := 0                                                // Valor Mercadoria( F2_VALMERC )
    Local  cBanco       := ""                                               // Codigo Banco( C5_BANCO )
    Static cIncoTerm   := ""                                                // Icoterm( C5_XICOTER )
    Local  nValBrut     := 0                                                // Valor Bruto( F2_VALBRUT )
    Static aDadosEmp    := {}                                               // Dados da Empresa
    Local  aBanco       := {}                                               // Dados do Banco
    Local  aTerceiro    := {}                                               // Dados de Terceiros
    Local  cArquivo     := ""                                               // Arquivo PDF
    Local  cPathPDF     := ""                                               // Pasta Temporaria do Sistema
    Local  nQuant       := 0                                                // Quantidade de Itens
    Static nPesoL       := 0                                                // Peso Liquido( C5_PESOL )
    Static nPBruto      := 0                                                // Peso Bruto( C5_PBRUTO )
    Static cTransp      := ""                                               // Transportadora( C5_TRANSP )
    Static cCondPag     := ""                                               // Condicao Pagamento( C5_CONDPAG )
    Static cDesembarque := ""                                               // Aeroporto Desembarque( C5_XDESEM )
    Static cEmbarque    := ""                                               // Aeroporto Embarque( C5_XEMBARQ )


    /*-------------------------------+
    |  Declaracao de Objetos         |
    +-------------------------------*/
    Local  oInvoice  // Grafico Invoice 
    Static oFont     // Fonte 
    Static oFontBold // Fonte 

    /*---------------------------------+
    | SELECT Tabela SF2                |
    +---------------------------------*/
    cQuery := "SELECT"      + ENTER
    cQuery += "F2_DOC,"     + ENTER
    cQuery += "F2_SERIE,"   + ENTER
    cQuery += "F2_VALMERC," + ENTER 
    cQuery += "F2_FRETE,"   + ENTER 
    cQuery += "F2_VALBRUT," + ENTER


    /*---------------------------------+
    | SELECT Tabela SD2                |
    +---------------------------------*/
    cQuery += "D2_COD,"     + ENTER 
    cQuery += "D2_PEDIDO,"  + ENTER 
    cQuery += "D2_DTVALID," + ENTER 
    cQuery += "D2_CLIENTE," + ENTER 
    cQuery += "D2_LOTECTL," + ENTER
    cQuery += "D2_EMISSAO," + ENTER 
    cQuery += "D2_LOJA,"    + ENTER 
    cQuery += "D2_QUANT,"   + ENTER
    cQuery += "D2_UM,"      + ENTER 
    cQuery += "D2_PRCVEN,"  + ENTER 
    cQuery += "D2_TOTAL,"   + ENTER

    /*---------------------------------+
    | SELECT Tabela SA1                |
    +---------------------------------*/
    cQuery += "A1_NREDUZ,"  + ENTER
    cQuery += "A1_BAIRRO,"  + ENTER 
    cQuery += "A1_END,"     + ENTER
    cQuery += "A1_EST,"     + ENTER 
    cQuery += "A1_CGC,"     + ENTER
    cQuery += "A1_CEP,"     + ENTER 
    cQuery += "A1_TEL,"     + ENTER 
    cQuery += "A1_PAIS,"    + ENTER 
    cQuery += "A1_MUN,"     + ENTER
    cQuery += "A1_DDD,"     + ENTER

    /*---------------------------------+
    | SELECT Tabela SB1                |
    +---------------------------------*/
    cQuery += "B1_PESO,"   + ENTER
    cQuery += "B1_DESC,"   + ENTER
    cQuery += "B1_ANVISA," + ENTER 

    /*---------------------------------+
    | SELECT Tabela SC5                |
    +---------------------------------*/  
    cQuery += "C5_MOEDA,"   + ENTER   
    cQuery += "C5_BANCO,"   + ENTER 
    cQuery += "C5_XINCO,"   + ENTER
    cQuery += "C5_PESOL,"   + ENTER 
    cQuery += "C5_XEMBARQ," + ENTER 
    cQuery += "C5_XDESEM,"   + ENTER
    cQuery += "C5_CONDPAG," + ENTER 
    cQuery += "C5_TRANSP,"  + ENTER 
    cQuery += "C5_PBRUTO"   + ENTER 

    /*----------------+
    | FROM            |
    +----------------*/
    cQuery += "FROM"                                   + ENTER 
    cQuery += + RetSqlName( "SD2" ) + " SD2 (NOLOCK) " + ENTER


    /*----------------+
    | INNER JOIN      |
    +----------------*/
    cQuery += "INNER JOIN " + RetSqlName( "SF2" ) + " (NOLOCK) AS SF2 ON D2_DOC = F2_DOC AND D2_SERIE = F2_SERIE"   + ENTER
    cQuery += "INNER JOIN " + RetSqlName( "SB1" ) + " (NOLOCK) AS SB1 ON D2_COD = B1_COD"                           + ENTER
    cQuery += "INNER JOIN " + RetSqlName( "SA1" ) + " (NOLOCK) AS SA1 ON D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA" + ENTER
    cQuery += "INNER JOIN " + RetSqlName( "SC5" ) + " (NOLOCK) AS SC5 ON D2_PEDIDO = C5_NUM"                        + ENTER


    /*----------------+
    | WHERE           |
    +----------------*/
    cQuery += "WHERE "                                     + ENTER 
    cQuery += "D2_DOC = '"+ MV_PAR01 +"' AND"              + ENTER
    cQuery += "D2_SERIE = '"+ AllTrim( MV_PAR02 ) +"' AND" + ENTER 
    cQuery += "SD2.D_E_L_E_T_ = '' "                       + ENTER


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
    

    /*----------------------------------------------+
    |  Armazenando Informacoes do Cliente/Pedido    |
    +----------------------------------------------*/
    aAdd( aHeader, { TMP -> D2_PEDIDO                                                      , ; // 01 - Numero do Pedido 
		     TMP -> D2_EMISSAO                                                     , ; // 02 - Data de Emissao
                     TMP -> C5_MOEDA                                                       , ; // 03 - Moeda de Conversao
        	     AllTrim( TMP -> A1_END )                                              , ; // 04 - Endereco Cliente
		     AllTrim( TMP -> A1_BAIRRO )                                           , ; // 05 - Bairro Cliente
		     AllTrim( TMP -> A1_MUN )                                              , ; // 06 - Municipio Cliente
		     TMP -> A1_EST                                                         , ; // 07 - Estado Cliente
	             TMP -> A1_CEP                                                         , ; // 08 - CEP Cliente
		     AllTrim( TMP -> A1_DDD )                                              , ; // 09 - DDD Cliente
		     AllTrim( TMP -> A1_TEL )                                              , ; // 10 - Telefone Cliente
		     TMP -> A1_NREDUZ                                                      , ; // 11 - Nome Cliente
                     Posicione( "SYA", 1, FWxFilial( "SYA" ) + TMP -> A1_PAIS, "YA_DESCR" );   // 12 - Pais Cliente
                    } )


    /*----------------------------------------------+
    |  Armazenando Informacoes de Terceiros         |
    +----------------------------------------------*/
    If !Empty( MV_PAR03 )
        aAdd( aTerceiro, { Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_NREDUZ" ) , ; // 01 - Nome Terceiro
                           Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_END"    ) , ; // 02 - Endereco Terceiro
                           Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_CEP"    ) , ; // 03 - CEP Terceiro
                           Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_EST"    ) , ; // 04 - Estado Terceiro
                           Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_DDD"    ) , ; // 05 - DDD Terceiro
                           Posicione( "SA1", 1, FWxFilial( "SA1" ) + MV_PAR03 + MV_PAR04, "A1_TEL"    )   ; // 06 - Telefone Terceiro
                        } )
    EndIf


    /*-------------------------------------------+
    |  Armazenando Informacoes da Empresa        |
    +-------------------------------------------*/
    aDadosEmp := FWSM0Util():GetSM0Data()


    /*-------------------------------------------+
    |  Armazenando Informacoes de Cabecalho      |
    +-------------------------------------------*/
    nFrete       := TMP -> F2_FRETE
    nValMerc     := TMP -> F2_VALMERC
    nValBrut     := TMP -> F2_VALBRUT
    cBanco       := TMP -> C5_BANCO
    cIncoTerm    := TMP -> C5_XINCO
    nPesoL       := Transform( TMP -> C5_PESOL , "@E 999,999.9999" )
    nPBruto      := Transform( TMP -> C5_PBRUTO, "@E 999,999.9999" )
    cTransp      := Posicione( "SA4", 1, FWxFilial( "SA4" ) + TMP -> C5_TRANSP , "A4_NREDUZ" )
    cCondPag     := Posicione( "SE4", 1, FWxFilial( "SE4" ) + TMP -> C5_CONDPAG, "E4_COND"   )
    cEmbarque    := TMP -> C5_XEMBARQ
    cDesembarque := TMP -> C5_XDESEM 



    /*--------------------------------+
    |  Consulta SQL para o Banco      |
    +--------------------------------*/
    cQry := "SELECT "      + ENTER 
    cQry += "A6_COD, "     + ENTER 
    cQry += "A6_NOME, "    + ENTER 
    cQry += "A6_AGENCIA, " + ENTER 
    cQry += "A6_NUMCON, "  + ENTER
    cQry += "A6_XSWIFT, "  + ENTER
    cQry += "A6_XIBAN "    + ENTER   
    /*---------------------------------------------------------*/
    cQry += "FROM " + ENTER 
    cQry += + RetSqlName( "SA6" ) + " SA6 (NOLOCK) " + ENTER
    /*---------------------------------------------------------*/
    cQry += "WHERE "                       + ENTER
    cQry += "A6_COD = '"+ cBanco +"' AND " + ENTER
    cQry += "SA6.D_E_L_E_T_ = '' "


    /*------------------------------------------+ 
    | Gera um Alias Temporario                  |
    +------------------------------------------*/
    If Select( "QRY" ) > 0 // Valida se o Alias esta aberto
        QRY -> ( DbCloseArea() )
    EndIf
    /*------------------------------------------------------------------*/
    TCQuery cQry New Alias "QRY" //  Gera um Alias a partir da Query
    /*------------------------------------------------------------------*/
    DbSelectArea( "QRY" ) // Seleciona o Alias Gerado


    /*------------------------------------------+
    |  Armazenando Informacoes do Banco         |
    +------------------------------------------*/
    aAdd( aBanco, { AllTrim( QRY -> A6_NOME    ), ; // 01 - Nome Banco
                    AllTrim( QRY -> A6_AGENCIA ), ; // 02 - Numero Agencia
                    AllTrim( QRY -> A6_NUMCON  ), ; // 03 - C/C 
                    AllTrim( QRY -> A6_XIBAN   ), ; // 04 - IBAN
                    AllTrim( QRY -> A6_XSWIFT  )  ; // 05 - SWIFT    
                } )
    /*------------------------------------------------------------------*/
    QRY -> ( DbCloseArea() ) // Fecha Alias Temporario 


    /*-------------------------------------------+
    |  Definindo Dados da Impressao              |
    +-------------------------------------------*/
	oBrush := TBrush():New( , nCorC ) // Definindo a Cor da Tinta
    /*-----------------------------------------------------------------------------------------------------------------------*/
    cArquivo := "Invoice_" + aHeader[1][1] + "_" + dToS( Date() ) + ".pdf" // Nome do Arquivo
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
    oFontBold := TFont():New( "Arial" ,, 07,, .T.,,,,, .F., .F. ) // Com Negrito
    /*---------------------------------------------------------------------------*/
    oFont     := TFont():New( "Arial" ,, 07,, .F.,,,,, .F., .F. ) // Sem Negrito


    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Box )     |
    +-------------------------------------------*/
	oInvoice:FillRect( { 020, 015, 030, 580 }, oBrush )
	oInvoice:Line( 020, 015, 020, 580 )
	oInvoice:Line( 020, 015, 030, 015 )
	oInvoice:Line( 020, 580, 030, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 030, 015, 065, 385 )
	oInvoice:Box( 030, 385, 045, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 045, 385, 055, 580 )
	oInvoice:Box( 055, 385, 065, 580 )
    /*------------------------------------------------------*/
	oInvoice:FillRect( { 065, 015, 075, 385 }, oBrush )
	oInvoice:FillRect( { 065, 385, 075, 580 }, oBrush )
	oInvoice:Line( 065, 015, 075, 015 )
	oInvoice:Line( 065, 015, 065, 580 )
	oInvoice:Line( 065, 385, 075, 385 )
	oInvoice:Line( 065, 580, 075, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 075, 015, 085, 385 )
	oInvoice:Box( 075, 385, 085, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 085, 015, 095, 385 )
	oInvoice:Box( 085, 385, 095, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 095, 015, 105, 385 )
	oInvoice:Box( 095, 385, 105, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 105, 015, 115, 385 )
	oInvoice:Box( 105, 385, 115, 580 )
    /*------------------------------------------------------*/
	oInvoice:FillRect( { 115, 015, 125, 385 }, oBrush )
	oInvoice:FillRect( { 115, 385, 125, 580 }, oBrush )
	oInvoice:Line( 115, 015, 115, 580 )
	oInvoice:Line( 115, 015, 125, 015 )
	oInvoice:Line( 115, 385, 125, 385 )
	oInvoice:Line( 115, 580, 125, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 125, 015, 135, 385 )
	oInvoice:Box( 125, 385, 135, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 135, 015, 145, 385 )
	oInvoice:Box( 135, 385, 145, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 145, 015, 155, 385 )
	oInvoice:Box( 145, 385, 155, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 155, 015, 165, 385 )
	oInvoice:Box( 155, 385, 165, 580 )
    /*------------------------------------------------------*/
	oInvoice:FillRect( { 165, 015, 175, 580 }, oBrush )
	oInvoice:Line( 165, 015, 165, 580 )
	oInvoice:Line( 165, 015, 175, 015 )
	oInvoice:Line( 165, 580, 175, 580 )
    /*------------------------------------------------------*/
	oInvoice:Box( 175, 015, 215, 055 )
	oInvoice:Box( 175, 055, 215, 095 )
	oInvoice:Box( 175, 095, 215, 125 )
	oInvoice:Box( 175, 125, 215, 155 )
	oInvoice:Box( 175, 155, 215, 185 )
	oInvoice:Box( 175, 185, 215, 225 )
	oInvoice:Box( 175, 225, 215, 385 )
	oInvoice:Box( 175, 385, 215, 475 )
	oInvoice:Box( 175, 475, 215, 580 )

    /*-----------------------+
    |  Impressao da Logo     |
    +-----------------------*/
    oInvoice:SayBitMap( 032, 200, cLogo, 45, 30 )


    /*-------------------------------------------+
    |  Definindo layout da Impressao ( Say )     |
    +-------------------------------------------*/
    oInvoice:Say( 028, 016, "Date / Fecha: " + Right( aHeader[1][2], 2 ) + "/" + SubStr( aHeader[1][2], 5, 2 ) + "/" + Left( aHeader[1][2], 4 ), oFontBold )
	oInvoice:Say( 028, 235, "COMMERCIAL INVOICE / Factura Comercial"				                                       , oFontBold )
	oInvoice:Say( 038, 016, "Company Name / Nombre de la compañia"  			                                               , oFont     )
	oInvoice:Say( 040, 386, "Invoice / Factura Comercial: # " + aHeader[1][1]		 	                                       , oFont     )
    oInvoice:Say( 053, 386, AllTrim( cCondPag ) + " dias después del la recepcion"                                                             , oFont     )
	oInvoice:Say( 063, 386, "Air Waybill / Guía Aérea: "            				                                       , oFont     )
	oInvoice:Say( 073, 155, "Ship From / Envio de"                  				                                       , oFontBold )
	oInvoice:Say( 073, 455, "Ship to / Envio para"                  			                                               , oFontBold )
	oInvoice:Say( 083, 386, "Name / Nombre: " + aHeader[1][11]            				                                       , oFont     )
	oInvoice:Say( 083, 016, "Name / Nombre: " + AllTrim( aDadosEmp[4][2] ) 			                                               , oFont     )
	oInvoice:Say( 093, 016, "Adress / Dirección: " + AllTrim( aDadosEmp[14][2] )		                                               , oFont     )
	oInvoice:Say( 093, 386, "Adress / Dirección: " + aHeader[1][4]       				                                       , oFont     )
	oInvoice:Say( 103, 016, "City, State, Zip / Ciudad, C. Postal: " + AllTrim( aDadosEmp[26][2] ) + " / " + AllTrim( aDadosEmp[25][2] )   , oFont     )
	oInvoice:Say( 103, 386, "City, State, Zip / Ciudad, C. Postal: " + aHeader[1][8] + " / " + aHeader[1][7]                               , oFont 	   )
	oInvoice:Say( 113, 016, "Phone / Teléfono: " + AllTrim( aDadosEmp[6][2] )                                                              , oFont     )
	oInvoice:Say( 113, 386, "Phone / Teléfono: (" + aHeader[1][9] + ") " + Transform( aHeader[1][10], "@E 9999-9999" )                     , oFont     )


    /*-------------------------------------------+
    |  Validacao de Entrega a Terceiros          |
    +-------------------------------------------*/
    If !Empty( aTerceiro )
	    oInvoice:Say( 123, 115, "Third Party Shipment / Envío a Tercera Persona"		                         , oFontBold )
	    oInvoice:Say( 133, 016, "Name / Nombre: " + AllTrim( aTerceiro[1][1] )               		         , oFont     )
	    oInvoice:Say( 143, 016, "Adress / Dirección: " + AllTrim( aTerceiro[1][2] ) 				 , oFont     )
	    oInvoice:Say( 153, 016, "City, State, Zip / Ciudad, C. Postal: " + aTerceiro[1][3] + "/ " + aTerceiro[1][4]  , oFont     )
	    oInvoice:Say( 163, 016, "Phone / Teléfono: (" + AllTrim( aTerceiro[1][5] ) + ")" + AllTrim( aTerceiro[1][6] ), oFont     )  
    Else
        oInvoice:Say( 123, 115, "Third Party Shipment / Envío a Tercera Persona", oFontBold )
	    oInvoice:Say( 133, 016, "Name / Nombre: "                  		, oFont     )
	    oInvoice:Say( 143, 016, "Adress / Dirección: "                  	, oFont     )
	    oInvoice:Say( 153, 016, "City, State, Zip / Ciudad, C. Postal: "    , oFont     )
	    oInvoice:Say( 163, 016, "Phone / Teléfono: "                        , oFont     )
    EndIf
    

    /*-------------------------------------------+
    |  Preenchimento do CheckBox                 |
    +-------------------------------------------*/
    Do Case 
        Case cIncoTerm == "CIP" 
	        oInvoice:Say( 123, 440, "Check One / Selecione Uno"                     		      , oFontBold )
	        oInvoice:Say( 133, 386, "[ x ] CIF Country of Export / País de Exportación: " + aHeader[1][12], oFont     )
	        oInvoice:Say( 143, 386, "[  ] FOB Country of Manufacture / País de Fabricación: "             , oFont     )
	        oInvoice:Say( 153, 386, "[  ] CPT Country of Destination / País de Destino: "                 , oFont     )
        Case cIncoTerm == "FOT"
            oInvoice:Say( 123, 440, "Check One / Selecione Uno"                     		                   , oFontBold )
	        oInvoice:Say( 133, 386, "[  ] CIF Country of Export / País de Exportación: "                       , oFont     )
	        oInvoice:Say( 143, 386, "[ x ] FOB Country of Manufacture / País de Fabricación: " + aHeader[1][12], oFont     )
	        oInvoice:Say( 153, 386, "[  ] CPT Country of Destination / País de Destino: "                      , oFont     )
        Case cIncoTerm == "CPT"
            oInvoice:Say( 123, 440, "Check One / Selecione Uno"                     		               , oFontBold )
	        oInvoice:Say( 133, 386, "[  ] CIF Country of Export / País de Exportación: "                   , oFont     )
	        oInvoice:Say( 143, 386, "[  ] FOB Country of Manufacture / País de Fabricación: "              , oFont     )
	        oInvoice:Say( 153, 386, "[ x ] CPT Country of Destination / País de Destino: " + aHeader[1][12], oFont     )
        OtherWise 
            oInvoice:Say( 123, 440, "Check One / Selecione Uno"                     		 , oFontBold )
	        oInvoice:Say( 133, 386, "[  ] CIF Country of Export / País de Exportación: "     , oFont     )
	        oInvoice:Say( 143, 386, "[  ] FOB Country of Manufacture / País de Fabricación: ", oFont     )
	        oInvoice:Say( 153, 386, "[  ] CPT Country of Destination / País de Destino: "    , oFont     )
    EndCase


    /*----------------------------------+
    |  Preenchimento da Moeda           |
    +----------------------------------*/
    Do Case
        Case aHeader[1][3] == 1 
	        oInvoice:Say( 163, 386, "Currency / Moneda: REAL", oFont )
        Case aHeader[1][3] == 2
            oInvoice:Say( 163, 386, "Currency / Moneda: DOLAR", oFont )
        Case aHeader[1][3] == 3
            oInvoice:Say( 163, 386, "Currency / Moneda: EURO" , oFont )
        OtherWise
            oInvoice:Say( 163, 386, "Currency / Moneda: "     , oFont )
    EndCase


	oInvoice:Say( 173, 205, "Package Information / Información del (los) Paquete(s)" , oFontBold )
	oInvoice:Say( 183, 028, "Qty / "                                                 , oFontBold )
	oInvoice:Say( 213, 022, "Cantidad"                                               , oFontBold )
	oInvoice:Say( 198, 068, "COD"                                                    , oFontBold )
	oInvoice:Say( 198, 105, "UN"                                                     , oFontBold )
	oInvoice:Say( 198, 135, "CP"                                                     , oFontBold )
	oInvoice:Say( 183, 190, "Unit Value "                                            , oFontBold )
	oInvoice:Say( 198, 195, "/ Valor "                                               , oFontBold )
	oInvoice:Say( 213, 193, "Unitario "                                              , oFontBold )
	oInvoice:Say( 183, 160, "No. Of "                                                , oFontBold )
	oInvoice:Say( 193, 156, "pkgs / No"                                              , oFontBold )
	oInvoice:Say( 203, 165, "De"                                                     , oFontBold )
	oInvoice:Say( 213, 156, "Paquetes"                                               , oFontBold )
	oInvoice:Say( 198, 229, "Commodity Description / Descripción del Producto"       , oFontBold )
	oInvoice:Say( 198, 410, "Weight / Peso "                                         , oFontBold )
	oInvoice:Say( 198, 490, "Total Value / Valor Total "                             , oFontBold )
    

    /*-------------------------------------------+
    |  Impressao dos Itens                       |
    +-------------------------------------------*/
    TMP -> ( DbGoTop() ) // Posiciona o Primeiro Registro da Tabela
    /*--------------------------------------------------------------*/
    While !TMP -> ( Eof() ) // Enquanto não for o Final da Tabela
        /*-------------------------------+
        |  Validar o fim da Pagina       |
        +-------------------------------*/
        BRKLINE() // Funcao para validar o final da pagina
        

        /*-----------------------+
        |  Layout Itens          |
        +-----------------------*/
        oInvoice:Box( nLine, 015, nLine + 10, 055 )
        oInvoice:Box( nLine, 055, nLine + 10, 095 )
        oInvoice:Box( nLine, 095, nLine + 10, 125 )
        oInvoice:Box( nLine, 125, nLine + 10, 155 )
        oInvoice:Box( nLine, 155, nLine + 10, 185 )
        oInvoice:Box( nLine, 185, nLine + 10, 225 )
        oInvoice:Box( nLine, 225, nLine + 10, 385 )
        oInvoice:Box( nLine, 385, nLine + 10, 475 )
        oInvoice:Box( nLine, 475, nLine + 10, 580 )
        /*-------------------------------------------------------------------------------------------*/
        oInvoice:Say( nLine + 8, 031, cValToChar( TMP -> D2_QUANT )                      , oFont )
        oInvoice:Say( nLine + 8, 068, TMP -> D2_COD                                      , oFont )
        oInvoice:Say( nLine + 8, 105, TMP -> D2_UM                                       , oFont )
        oInvoice:Say( nLine + 8, 135, "CP"                                               , oFont )
        oInvoice:Say( nLine + 8, 166, cValToChar( TMP -> D2_QUANT )                      , oFont )
        oInvoice:Say( nLine + 8, 186, Transform( TMP -> D2_PRCVEN, "@E 9,999,999.99"    ), oFont )
        oInvoice:Say( nLine + 8, 226, TMP -> B1_DESC                                     , oFont )

        If Empty( TMP -> B1_PESO )
            oInvoice:Say( nLine + 8, 386, " - ", oFont )
        Else 
            oInvoice:Say( nLine + 8, 386, Transform( TMP -> B1_PESO, "@E 999,999.9999" )   , oFont )
        EndIf 

        oInvoice:Say( nLine + 8, 476, "$" + Transform( TMP -> D2_TOTAL, "@E 9,999,999.99" ), oFont )

        /*--------------------------------------------+
        |  Armezenando dados para variaveis auxiliar  |
        +--------------------------------------------*/
        nLine  := nLine + 10 // Controle de Linha
        /*------------------------------------------------------------------------*/
        nQuant := nQuant + TMP -> D2_QUANT // Controle da Quantidade de Produto


        /*-----------------------+
        |  Proximo Registro      |
        +-----------------------*/
        TMP -> ( DbSkip() )
    EndDo
    

    /*-------------------------------+
    |  Validar o fim da Pagina       |
    +-------------------------------*/
    If nLine >= nLineEnd - 60 
        oInvoice:EndPage() // Finaliza a Pagina
       /*----------------------------------------------*/
        oInvoice:StartPage() // Inicia uma nova pagina
       /*----------------------------------------------*/
        nLine := 020 // Linha Inicial 
    EndIf

    /*-----------------------+
    |  Layout Rodape         |
    +-----------------------*/
    oInvoice:Box( nLine + 20, 385, nLine + 30, 475 )
    oInvoice:Box( nLine + 30, 385, nLine + 40, 475 )
    oInvoice:Box( nLine + 40, 385, nLine + 50, 475 )
    oInvoice:Box( nLine + 50, 385, nLine + 60, 475 )
    /*------------------------------------------------------------------------------------*/
    oInvoice:Box( nLine + 20, 475, nLine + 30, 580 )
    oInvoice:Box( nLine + 30, 475, nLine + 40, 580 )
    oInvoice:Box( nLine + 40, 475, nLine + 50, 580 )
    oInvoice:Box( nLine + 50, 475, nLine + 60, 580 )
    /*------------------------------------------------------------------------------------*/
    oInvoice:Say( nLine + 28, 386, "Total Produtos "       , oFontBold )
    oInvoice:Say( nLine + 38, 386, "Quantidade de Produtos", oFontBold )
    oInvoice:Say( nLine + 48, 386, "Freight "              , oFontBold )
    oInvoice:Say( nLine + 58, 386, "Total "                , oFontBold )
    /*---------------------------------------------------------------------------------------*/
    oInvoice:Say( nLine + 28, 477, "$ " + Transform( nValMerc, "@E 9,999,999.99" ), oFont )
    oInvoice:Say( nLine + 38, 477, cValToChar( nQuant )                           , oFont )

    If nFrete == 0
       oInvoice:Say( nLine + 48, 477, " - ", oFont )
    Else 
        oInvoice:Say( nLine + 48, 477, "$" + Transform( nFrete, "@E 9,999,999.99" ), oFont )
    EndIf 

    oInvoice:Say( nLine + 58, 477, "$" + Transform( nValBrut, "@E 9,999,999.99" ), oFont )
    

    /*------------------------------------+
    |  Imprimindo Informações do Banco    |
    +------------------------------------*/        
    oInvoice:Box( nLine + 20, 015, nLine + 070, 200 )
    /*-----------------------------------------------------------------*/
    oInvoice:Say( nLine + 28, 016, "BANCO: "   + aBanco[1][1], oFont )
    oInvoice:Say( nLine + 38, 016, "AGÊNCIA: " + aBanco[1][2], oFont )
    oInvoice:Say( nLine + 48, 016, "C/C: "     + aBanco[1][3], oFont )
    oInvoice:Say( nLine + 58, 016, "IBAN: "    + aBanco[1][4], oFont )
    oInvoice:Say( nLine + 68, 016, "SWIFT: "   + aBanco[1][5], oFont )


    /*--------------+
    |  Assinatura   |
    +--------------*/ 
    If nLine + 120 >= nLineEnd 
        oInvoice:EndPage() // Finaliza a Pagina
        /*----------------------------------------------*/
        oInvoice:StartPage() // Inicia uma Nova Pagina
        /*----------------------------------------------*/
        nLine := 020 
        /*----------------------------------------------*/
        oInvoice:Line( nLine, 180, nLine, 400 )
        oInvoice:Say( nLine, 230, "Implacil de Bortoli Material Odontologico S.A.", oFont )
    Else
        oInvoice:Line( nLine + 108, 180, nLine + 108, 400 )
        oInvoice:Say( nLine + 115, 230, "Implacil de Bortoli Material Odontologico S.A.", oFont )
    EndIf


    /*-------------------------------+
    |  Finalizando e Exibindo PDF    |
    +-------------------------------*/ 
    oInvoice:EndPage() // Finaliza a Pagina
    /*--------------------------------------------------------------*/
    oInvoice:Preview() // Mostra um preview do PDF
    /*--------------------------------------------------------------*/
    RestArea( aArea ) // Restaura dados armazenados
Return( Nil )


*=========================*
Static Function BRKLINE() 
*=========================*
    /*-------------------------------------------+
    |  Validacao para finalizar a pagina         |
    +-------------------------------------------*/
    If nLine >= nLineEnd - 10
        oInvoice:EndPage() // Finaliza a Pagina
        /*----------------------------------------------*/
        oInvoice:StartPage() // Inicia uma Nova Pagina
        /*----------------------------------------------*/
        nLine := 020 // Linha Inicial
    EndIf
Return( Nil )


*=========================*
Static Function GERPACK() 
*=========================*
    /*-------------------------------+
    |  Declaracao de Variaveis       |
    +-------------------------------*/
    Local cDtValid := "" // Data Validade( D2_DTVALID )


    /*-------------------------------+
    |  Declaracao de Objetos         |
    +-------------------------------*/
    Local oPacking // Grafico Packing List
    Local oFontBold07 // Fonte 


    /*------------------------+
    |  Redeclarando Fonte     |
    +------------------------*/
    oFontBold   := TFont():New( "Arial" ,, 14,, .T.,,,,, .F., .F. ) // Com Negrito - Tamanho 14
    oFontBold07 := TFont():New( "Arial" ,, 07,, .T.,,,,, .F., .F. ) // Com Negrito - Tamanho 07

    /*-------------------------------------------+
    |  Definindo Dados da Impressao              |
    +-------------------------------------------*/
    cArquivo := "PackingList_" + aHeader[1][1] + "_" + dToS( Date() ) + ".pdf" // Nome do Arquivo
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


    /*---------------------------------+
    |  Definindo layout da Impressao   |
    +---------------------------------*/
    oPacking:SayBitMap( 060, 050, cLogo, 55, 40 )
    /*--------------------------------------------------------------*/
    oPacking:Say( 070, 270, "PACKING LIST", oFontBold )
    /*--------------------------------------------------------------*/
    oPacking:Box( 120, 050, 180, 290 )
    /*--------------------------------------------------------------*/
    oPacking:Box( 190, 050, 250, 290 )
    /*--------------------------------------------------------------*/
    oPacking:Box( 120, 340, 250, 550 )
    /*--------------------------------------------------------------*/
    oPacking:Box( 260, 050, 270, 550 )
    oPacking:Say( 268, 150, "UNITARY MERCHANDISES, AMOUNTS, PRICES AND DESCRIBED TOTAL PRICES IN ATTACHED LEVES", oFontBold07 )
    /*--------------------------------------------------------------*/
    oPacking:Box( 280, 050, 305, 250 )
    oPacking:Box( 280, 250, 305, 290 )
    oPacking:Box( 280, 290, 305, 340 )
    oPacking:Box( 280, 340, 305, 390 )
    oPacking:Box( 280, 390, 305, 510 )
    oPacking:Box( 280, 510, 305, 550 )
    /*------------------------------------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( 128, 051, aDadosEmp[5][2]                                                                                   , oFontBold07 )
    oPacking:Say( 138, 051, AllTrim( aDadosEmp[14][2] ) + " - " + aDadosEmp[16][2]                                            , oFont       )
    oPacking:Say( 148, 051, AllTrim( aDadosEmp[17][2] ) + " - " + aDadosEmp[18][2] + " - BRAZIL - CEP " + aDadosEmp[26][2]    , oFont       )  
    oPacking:Say( 158, 051, "CNPJ: " + aDadosEmp[10][2] + " - Phone: +" + aDadosEmp[6][2]                                     , oFont       )
    oPacking:Say( 168, 051, "Airport of Desparture: " + AllTrim( cEmbarque )                                                  , oFont       )  
    oPacking:Say( 178, 051, cIncoTerm + " - " + AllTrim( aDadosEmp[17][2] ) + " - " + aDadosEmp[18][2] + " - BRAZIL"          , oFont       )
    /*------------------------------------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( 198, 051, "SHIP TO: "                                                   , oFontBold07 )
    oPacking:Say( 208, 051, aHeader[1][11]                                                , oFont       )
    oPacking:Say( 218, 051, "Adress / Dirección: " + aHeader[1][4]                        , oFont       )
    oPacking:Say( 228, 051, "Providencia - " + aHeader[1][12]                             , oFont       )
    oPacking:Say( 238, 051, "Phone / Teléfono: ( " + aHeader[1][9] + ") " + aHeader[1][10], oFont       )
    oPacking:Say( 248, 051, "Airport of Destination: " + AllTrim( cDesembarque )          , oFont       )
    /*------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( 128, 341, "INVOICE"                                                           , oFontBold07 )
    oPacking:Say( 138, 341, "Number " + aHeader[1][1]                                           , oFont       )
    oPacking:Say( 148, 341, "Paymente " + AllTrim( cCondPag ) + " dias después del la recepcion", oFont       )
    oPacking:Say( 158, 341, "Transportation: " + cTransp                                        , oFont       )
    oPacking:Say( 168, 341, "Mark: " + aDadosEmp[4][2]                                          , oFont       )
    oPacking:Say( 178, 341, "Gross Weight: " + nPBruto                                          , oFont       )
    oPacking:Say( 188, 341, "Net Weight: "   + nPesoL                                           , oFont       )
    /*-------------------------------------------------------------------------------------------------------------*/
    oPacking:Say( 295, 075, "Commodity Description / Descripción del Producto", oFontBold07 )
    oPacking:Say( 288, 265, "Qty /"                                           , oFontBold07 )
    oPacking:Say( 298, 255, "Cantidad"                                        , oFontBold07 )
    oPacking:Say( 288, 305, "Batch"                                           , oFontBold07 )
    oPacking:Say( 298, 303, "Number"                                          , oFontBold07 )
    oPacking:Say( 295, 348, "Expire Date"                                     , oFontBold07 )
    oPacking:Say( 295, 430, "Registration"                                    , oFontBold07 )
    oPacking:Say( 295, 515, "Item Code"                                       , oFontBold07 )


    /*-------------------------------+
    |  Redeclarando Linha Inical     |
    +-------------------------------*/
    nLine := 305


    /*-------------------------------------------+
    |  Impressao dos Itens                       |
    +-------------------------------------------*/
    TMP -> ( DbGoTop() ) // Posiciona o Primeiro Registro da Tabela
    /*--------------------------------------------------------------*/
    While !TMP -> ( Eof() ) // Enquanto não for o Final da Tabela
        cDtValid := TMP -> D2_DTVALID
    

        /*-------------------------------+
        |  Validar o fim da Pagina       |
        +-------------------------------*/
        BRKLINE() // Funcao para validar o final da pagina


        /*-----------------------+
        |  Layout Itens          |
        +-----------------------*/
        oPacking:Box( nLine, 050, nLine + 10, 250 )
        oPacking:Box( nLine, 250, nLine + 10, 290 )
        oPacking:Box( nLine, 290, nLine + 10, 340 )
        oPacking:Box( nLine, 340, nLine + 10, 390 )
        oPacking:Box( nLine, 390, nLine + 10, 510 )
        oPacking:Box( nLine, 510, nLine + 10, 550 )
        /*-----------------------------------------------------------------------------------------*/
        oPacking:Say( nLine + 8, 051, TMP -> B1_DESC               , oFont )        
        oPacking:Say( nLine + 8, 266, cValToChar( TMP -> D2_QUANT ), oFont )
        oPacking:Say( nLine + 8, 298, TMP -> D2_LOTECTL            , oFont )
        oPacking:Say( nLine + 8, 350, Right( cDtValid, 2 ) + "/" + SubStr( cDtValid, 5, 2 ) + "/" + Left( cDtValid, 4 ), oFont )
        oPacking:Say( nLine + 8, 430, TMP -> B1_ANVISA             , oFont )
        oPacking:Say( nLine + 8, 521, TMP -> D2_COD                , oFont )
   

        /*--------------------------------------------+
        |  Armezenando dados para variaveis auxiliar  |
        +--------------------------------------------*/
        nLine  := nLine + 10 // Controle de Linha
        

        /*-----------------------+
        |  Proximo Registro      |
        +-----------------------*/
        TMP -> ( DbSkip() )
    EndDo


    /*-------------------------------+
    |  Finalizando e Exibindo PDF    |
    +-------------------------------*/ 
    oPacking:EndPage() // Finaliza a Pagina
    /*--------------------------------------------------------------*/
    oPacking:Preview() // Mostra um preview do PDF
    /*--------------------------------------------------------------*/
    TMP -> ( DbCloseArea() ) // Fecha o Alias Temporario 
Return( Nil )
