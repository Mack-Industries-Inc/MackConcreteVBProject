Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Transactions
Imports System.Data
Imports Mongoose.Core.Common
Imports Mongoose.Core.DataAccess
Imports Mongoose.IDO
Imports Mongoose.IDO.Protocol
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports Mongoose.IDO.DataAccess
Imports Mongoose.Core

Namespace ue_Mack_SLOrderInvoicingCreditMemoReprintReport
    ' <IDOExtensionClass( "ue_Mack_SLOrderInvoicingCreditMemoReprintReport" )>
    Partial Public Class ue_Mack_SLOrderInvoicingCreditMemoReprintReport : Inherits IDOExtensionClass
        Public Function ue_Rpt_OrderInvoicingCreditMemoReprintSp(ByVal pSessionIDChar As String, ByVal InvType As String, ByVal Mode As String, ByVal StartInvNum As String, ByVal EndInvNum As String,
                                 ByVal StartOrderNum As String, ByVal EndOrderNum As String, ByVal StartInvDate As String, ByVal EndInvDate As String, ByVal StartCustNum As String, ByVal EndCustNum As String,
                                 ByVal PrintItemCustomerItem As String, ByVal TransToDomCurr As String, ByVal InvCred As String, ByVal PrintSerialNumbers As String, ByVal PrintPlanItemMaterial As String,
                                 ByVal PrintConfigurationDetail As String, ByVal PrintEuro As String, ByVal PrintCustomerNotes As String, ByVal PrintOrderNotes As String, ByVal PrintOrderLineNotes As String,
                                 ByVal PrintOrderBlanketLineNotes As String, ByVal PrintProgressiveBillingNotes As String, ByVal PrintInternalNotes As String, ByVal PrintExternalNotes As String, ByVal PrintItemOverview As String, ByVal DisplayHeader As String,
                                 ByVal PrintLineReleaseDescription As String, ByVal PrintStandardOrderText As String, ByVal PrintBillToNotes As String, ByVal LangCode As String, ByVal BGSessionId As String,
                                 ByVal PrintDiscountAmt As String, ByVal PrintLotNumbers As String, ByVal pSite As String, ByVal CalledFrom As String, ByVal InvoicBuilderProcessID As String, ByVal StartBuilderInvNum As String,
                                 ByVal EndBuilderInvNum As String, ByVal pPrintDrawingNumber As String, ByVal pPrintDeliveryIncoTerms As String, ByVal pPrintTax As String, ByVal pPrintEUDetails As String, ByVal pPrintCurrCode As String,
                                 ByVal pPrintHeaderOnAllPages As String, ByVal StartReprintInvDateOffset As String, ByVal EndReprintInvDateOffset As String
                                 ) As DataTable

            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim oRptCustomFieldsResponse As LoadCollectionResponseData
            Dim sFilter As String
            Dim sPropertiesList As String
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim sRecordFilter As String
            Dim oCoitem As LoadCollectionResponseData
            Dim oCo As LoadCollectionResponseData
            Dim oShipment As LoadCollectionResponseData
            Dim oInv As LoadCollectionResponseData
            Dim sResultSetColumnsList As String
            Dim invokeResponse As InvokeResponseData
            Dim StrCoLine As String
            Dim CoLine As String
            Dim CoRelease As String

            'EXEC dbo.ApplyDateOffsetSp @pStartDueDate OUTPUT, @pStartDueDateOffset, 0
            invokeResponse = Me.Context.Commands.Invoke("SLJobTrans", "ApplyDateOffsetSp", StartInvDate, StartReprintInvDateOffset, "0")
            StartInvDate = invokeResponse.Parameters(0).Value.ToString
            invokeResponse = Me.Context.Commands.Invoke("SLJobTrans", "ApplyDateOffsetSp", EndInvDate, EndReprintInvDateOffset, "1")
            EndInvDate = invokeResponse.Parameters(0).Value.ToString

            Dim ReportInputParametersArray As String() = New List(Of String) From {pSessionIDChar, InvType, Mode, StartInvNum, EndInvNum,
                                 StartOrderNum, EndOrderNum, StartInvDate, EndInvDate, StartCustNum, EndCustNum,
                                 PrintItemCustomerItem, TransToDomCurr, InvCred, PrintSerialNumbers, PrintPlanItemMaterial,
                                 PrintConfigurationDetail, PrintEuro, PrintCustomerNotes, PrintOrderNotes, PrintOrderLineNotes,
                                 PrintOrderBlanketLineNotes, PrintProgressiveBillingNotes, PrintInternalNotes, PrintExternalNotes, PrintItemOverview, DisplayHeader,
                                 PrintLineReleaseDescription, PrintStandardOrderText, PrintBillToNotes, LangCode, BGSessionId,
                                 PrintDiscountAmt, PrintLotNumbers, pSite, CalledFrom, InvoicBuilderProcessID, StartBuilderInvNum,
                                 EndBuilderInvNum, pPrintDrawingNumber, pPrintDeliveryIncoTerms, pPrintTax, pPrintEUDetails, pPrintCurrCode, pPrintHeaderOnAllPages}.ToArray


            sResultSetColumnsList = "TxType,InvNum,InvMemoNum,CustNum,CoNum,InvLine,InvSite,InvDate,InvSlsman,InvDescription,InvTaxNumLbl1,InvTaxNum1,InvTaxNumLbl2,"
            sResultSetColumnsList = sResultSetColumnsList & "InvTaxNum2,InvCustTaxNumLbl1,InvCustTaxNum1,InvCustShiptoTaxNum1,InvCustTaxNumLbl2,InvCustTaxNum2,InvCustShiptoTaxNum2,InvCurrCode,"
            sResultSetColumnsList = sResultSetColumnsList & "InvCustSeq,InvStrCustSeq,InvFaxNum,InvLongCustPo,InvShortCustPo,InvPkgs,InvWeight,InvShipvia,InvTerms,TermsPct,TaxAmtLabel1,TaxAmtLabel2,"
            sResultSetColumnsList = sResultSetColumnsList & "InvSaleAmt,InvDiscAmt,InvNetAmt,InvCoText1,InvMiscCharges,InvCoText2,InvFreight,InvCoText3,InvSalesTax,InvSalesTax2,InvPrepaidAmt,InvTotal,"
            sResultSetColumnsList = sResultSetColumnsList & "InvPrintEuro,InvEuroTotal,OurAddress,BillToAddress,ShipToAddress,DropShipAddress,TaxCodeLbl,InvTaxCode,TaxCodeELbl,TaxCodeE,TaxRate,TaxBasis,"
            sResultSetColumnsList = sResultSetColumnsList & "ExtendedTax,CurrCode,CurrDescription,CoLine,StrCoLine,QtyConv,Qty,Back,CprPrice,ExtPrice,ShipDat,Contact,CosigneeName,CosigneeAddr1,CosigneeAddr2,"
            sResultSetColumnsList = sResultSetColumnsList & "CosigneeAddr3,CosigneeAddr4,CosigneeCity,CosigneeState,CosigneeZip,CosigneeCountry,AmtOldPrice,AmtNewPrice,Item1,Item2,ItemDesc,ItemOverview,CoitemUM,"
            sResultSetColumnsList = sResultSetColumnsList & "XInvItemDoLine,XInvItemDoSeq,CoitemShipDate,XInvItemCustPo,TaxItemLabel,TaxCode,TaxCodeDescription,OrdNum,Lcr,SerNum,AmtProgExPrice,LineRelease,"
            sResultSetColumnsList = sResultSetColumnsList & "DoNum,ConfigId,CompId,CompName,ConfigQty,Price,AttrName,AttrValue,InvProSeq,InvProDescription,Subtotal,Credit,FBOMFeatStr,FBOMQtyConv,FBOMUM,"
            sResultSetColumnsList = sResultSetColumnsList & "FBOMDescription,CustomerRowPointer,CustomerNotesFlag,CoRowPointer,CoNotesFlag,CoitemRowPointer,CoitemNotesFlag,CoBlnRowPointer,CoBlnNotesFlag,"
            sResultSetColumnsList = sResultSetColumnsList & "CoitemLineReleaseDes,ProgbillRowPointer,ProgbillNotesFlag,RptKey,BillToRowPointer,BillToNotesFlag,PackNum1,PackNum2,PackNum3,PackNum4,PackNum5,PackNum6,"
            sResultSetColumnsList = sResultSetColumnsList & "PackNum7,PackNum8,PrintTaxInvoice,OrigInvNum,ReasonText,TaxSystemRate1,TaxSystemRate2,TaxSystem1Enabled,TaxSystem2Enabled,TaxMode1,TaxMode2,AmtTotal,"
            sResultSetColumnsList = sResultSetColumnsList & "PriceWithoutTax,IncludeTaxInPrice,TaxAmt,TaxAmt2,UseMultiDueDates,MultiDueInvSeq,MultiDueDate,MultiDuePercent,MultiDueAmount,ApplyToInvNum,DomPriceFormat,"
            sResultSetColumnsList = sResultSetColumnsList & "DomPricePlaces,DomAmountFormat,DomAmountPlaces,Type,TermsDiscountAmt,TaxDiscAllow1,TaxDiscAllow2,KitComponent,KitComponentDesc,KitQtyRequired,KitUM,"
            sResultSetColumnsList = sResultSetColumnsList & "KitFlag,OrderError,QtyUnitFormat,PlacesQtyUnit,LotNum,TTxType,TCoNum,TRptKey,TParmsCompany,TParmsAddr1,TParmsAddr2,TParmsZip,TParmsCity1,TParmsCity2,"
            sResultSetColumnsList = sResultSetColumnsList & "TArinvAmount1,TArinvAmount2,TArinvInvDate,TArinvDueDate,TCustNum,TInvNum,TBankNumber,TBranchCode,TBankAcctNo1,TBankAcctNo2,TBankAddr1,TBankAddr2,"
            sResultSetColumnsList = sResultSetColumnsList & "TCustaddrName,TCustaddrAddr1,TCustaddrZip,TCustaddrCity,TCustdrftDraftNum,DemandSitePO,Url,EmailAddr,OfficeAddrFooter,DomAmtTotFormat,ReturnFlag,"
            sResultSetColumnsList = sResultSetColumnsList & "UseLongName,ShowHeader,PromotionCode,Surcharge,ItemContent,ParmsSingleLineAddress,DrawingNbr,Delterm,EcCode,Origin,CommCode,Description,EndUser,DueDate,"
            sResultSetColumnsList = sResultSetColumnsList & "TBankName,TBankTransitNum,TBankAcctNo,ShipmentId,RefLine,RefRelease"
            ',ue_CustPo,ue_JobName,ue_Structure,ue_Infobar,ue_SpecID

            oResultDataTable = New DataTable()
            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SL.SLOrderInvoicingCreditMemoReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_OrderInvoicingCreditMemoSp"

            For iIndex = 0 To ReportInputParametersArray.Length - 1
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest)
            If oRptStdResultResponse.Items.Count > 0 Then
                sResultSetColumnsList = sResultSetColumnsList & ",ue_CustPo,ue_JobName,ue_Structure,ue_Infobar,ue_SpecID"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")

                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                Next iIndex

                ' looks like here we are adding the data from the method being called into a data table variable 
                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    oResultDataRow = oResultDataTable.NewRow()
                    For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 6
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                    Next iColumnIndex
                    oResultDataTable.Rows.Add(oResultDataRow)
                Next iRowIndex

                For Each oRow As DataRow In oResultDataTable.Rows
                    sRecordFilter = String.Format("CoNum = {0}", SqlLiteral.Format(oRow.Item("CoNum")))
                    oCo = Me.Context.Commands.LoadCollection("SLCos", "coUf_JobProjName, CustPo", sRecordFilter, "", 0)
                    If oCo.Items.Count > 0 Then
                        oRow.Item("ue_CustPo") = oCo(0, "CustPo")
                        oRow.Item("ue_JobName") = oCo(0, "coUf_JobProjName")
                    End If

                    sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = 0", SqlLiteral.Format(oRow.Item("CoNum")), SqlLiteral.Format(oRow.Item("CoLine")))
                    oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "coiUf_Structure,coiUf_SpecCoLineID,Description", sRecordFilter, "", 0)
                    If oCoitem.Items.Count > 0 Then
                        oRow.Item("ue_Structure") = oCoitem(0, "coiUf_Structure")
                        oRow.Item("ue_SpecID") = oCoitem(0, "coiUf_SpecCoLineID")
                        oRow.Item("Description") = oCoitem(0, "Description")
                    End If

                    sRecordFilter = String.Format("RefNum = {0} and RefLineSuf = {1} and RefRelease = 0", SqlLiteral.Format(oRow.Item("CoNum")), SqlLiteral.Format(oRow.Item("CoLine")))
                    oShipment = Me.Context.Commands.LoadCollection("SLShipmentLines", "ShipmentId", sRecordFilter, "", 0)
                    If oShipment.Items.Count > 0 Then
                        oRow.Item("ShipmentId") = oShipment(0, "ShipmentId")
                    End If

                    sRecordFilter = String.Format("InvNum = {0}", SqlLiteral.Format(oRow.Item("InvNum")))
                    oInv = Me.Context.Commands.LoadCollection("SLInvHdrs", "TaxCode1", sRecordFilter, "", 0)
                    If oInv.Items.Count > 0 Then
                        oRow.Item("InvTaxCode") = oInv(0, "TaxCode1")
                    End If

                Next oRow
            End If


            Return oResultDataTable    'return the data table

        End Function



    End Class
End Namespace