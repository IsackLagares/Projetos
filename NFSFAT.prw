#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'TOTVS.CH'

/*
///////////////////////////////////////////////////////////////////////////////////////
//|---------------------------------------------------------------------------------|//
//| PROGRAMA	|   NFSFAT  | AUTOR | Isack R. Lagares | DATA | 16/02/2023	    |//
//|---------------------------------------------------------------------------------|//
//| DESCRICAO	| Faturar através de um Documento de Entrada                        |//
//|---------------------------------------------------------------------------------|//
//| USO		    | ExecAuto                                                      |//
//|---------------------------------------------------------------------------------|//
//| OBS		    | Colocar a observação	         	                    |//
//|---------------------------------------------------------------------------------|//
///////////////////////////////////////////////////////////////////////////////////////
*/

User Function NFSFAT()

    MsgRun( "Lendo Tabela...", "Processo", { || PvNFs() } )

Return ( Nil )

Static Function PvNFs()
    Local cDocF1        := SF1 -> F1_DOC       
    Local cSerieNfe     := SF1 -> F1_SERIE
    Local cSerieNfs     := "1"
    Local cPedVen       := GetSxeNum( "SC5", "C5_NUM", 1 )
    Local cDocF2        := GetSxeNum( "SF2", "F2_DOC", 1 )
    Local aHeader       := {}
    Local aItems        := {}
    Local aLine         := {}
    Local aBloqueio     := {}
    Local aPvlNfs       := {}
    Local cRet          := ""
    Private lMsErroAuto := .F.

    /*--------------------------------------------------------------+
    |   Iniciando a Geração e Liberação do Pedido de Venda          |
    +--------------------------------------------------------------*/
    // Monta arrays do cabeçalho
    aAdd( aHeader, { "C5_NUM",     cPedVen,  Nil } )
    aAdd( aHeader, { "C5_TIPO",    "B",      Nil } )
    aAdd( aHeader, { "C5_CLIENTE", "000510", Nil } ) 
    aAdd( aHeader, { "C5_LOJACLI", "01",     Nil } )
    aAdd( aHeader, { "C5_CONDPAG", "L01",    Nil } )
    aAdd( aHeader, { "C5_NATUREZ", "101020", Nil } )

    DbSelectArea( "SD1" )
    SD1 -> ( DBSetOrder(1) ) // D1_FILIAL + D1_DOC + D1_SERIE
    If SD1 -> ( DBSeek( FWxFilial( "SD1" ) + cDocF1 + cSerieNfe ) )

        // Enquanto não for o final da tabela  
        While !SD1 -> ( EoF() ) .AND. SD1 -> D1_DOC == cDocF1 .AND. SD1 -> D1_SERIE == cSerieNfe 
                // Monta arrays dos itens 
                aLine := {}
                aAdd( aLine, { "C6_ITEM",    SD1 -> D1_ITEMPC, Nil } )
                aAdd( aLine, { "C6_PRODUTO", SD1 -> D1_COD,    Nil } )
                aAdd( aLine, { "C6_QTDVEN",  SD1 -> D1_QUANT,  Nil } )
                aAdd( aLine, { "C6_DESC",    SD1 -> D1_DESC,   Nil } )
                aAdd( aLine, { "C6_PRCVEN",  SD1 -> D1_VUNIT,  Nil } )
                aAdd( aLine, { "C6_VALOR",   SD1 -> D1_TOTAL,  Nil } )
                aAdd( aLine, { "C6_TES",     "556",            Nil } )
                aAdd( aLine, { "C6_UM",      SD1 -> D1_UM,     Nil } )
                aAdd( aLine, { "C6_CF",      "5901",           Nil } )
                aAdd( aLine, { "C6_QTDLIB",  SD1 -> D1_QUANT,  Nil } )

                aAdd( aItems, aLine )

            SD1 -> ( DbSkip() )   
        EndDo

        SD1 -> ( DbCloseArea() )

        // Rotina Autmática 
        MsExecAuto( { | x, y, z |    MATA410( x, y, z ) }, aHeader, aItems, 3 )
                    
        If !lMsErroAuto
            FWAlertSuccess( "Pedido de Venda gerado com sucesso!", "Finalizado" )

            /*--------------------------------------------------------------+
            |  Inicia o Procedimento de Liberação do Faturamento            |
            +--------------------------------------------------------------*/
            SX5->( DbSetOrder( 1 ) )
            If SX5 -> ( DbSeek( xFilial( "SX5" ) + PadR( '01', Len( SX5 -> X5_TABELA ) ) + PadR( cSerieNfs, 3 ) ) )
                SC5 -> ( DbSetOrder( 1 ) )
                SC5 -> ( DbSeek( FWxFilial( "SC5" ) + cPedVen ) ) // C5_FILIAL + C5_NUM

                Ma410LbNfs( 2, @aPvlNfs, @aBloqueio )
                Ma410LbNfs( 1, @aPvlNfs, @aBloqueio )

                cRet := MaPvlNfs( aPvlNfs,   ;
                                  cSerieNfs, ;
                                  .F.,       ;
                                  .F.,       ;
                                  .F.,       ;
                                  .F.,       ;
                                  .F.,       ; 
                                  3  ,       ;  
                                  1  ,       ;
                                  .F.,       ;
                                  .F.,,,,,,  ;
                                  SC5 -> C5_EMISSAO )
                    
                If !Empty( cRet )
                    MsgInfo( "NOTA FISCAL: " + cDocF2 + " - " + cSerieNfs )
                EndIF

                FWAlertSuccess( "Processo Finalizado", "Finalizado" )
            Else 
                FWAlertError( "Erro ao gerar o Documento de Saída", "Erro" ) 
            EndIf
        Else
            FWAlertError( "Erro em gerar Pedido de Venda", "Erro")
            MostraErro()
        EndIf
    Else
        FWAlertError( "Erro em gerar o Documento de Saída", "Atenção" ) 
    EndIf

Return ( Nil )
