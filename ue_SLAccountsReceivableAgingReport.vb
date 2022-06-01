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

Namespace ue_SLAccountsReceivableAgingReport
    ' <IDOExtensionClass( "SLAccountsReceivableAgingReport" )>
    Partial Public Class SLAccountsReceivableAgingReport : Inherits IDOExtensionClass

        Public Function LoadAccountsReceivableAging(ByVal AgingDate As String, ByVal CutoffDate As String, ByVal AgingDateOffset As String, ByVal CutoffDateOffset As String,
                                ByVal StateCycle As String, ByVal ShowActive As String, ByVal BegSlsman As String, ByVal EndSlsman As String, ByVal CustomerStarting As String, ByVal CustomerEnding As String,
                                ByVal NameStarting As String, ByVal NameEnding As String, ByVal CurCodeStarting As String, ByVal CurCodeEnding As String, ByVal PrZeroBal As String,
                                ByVal CreditHold As String, ByVal PrCreditBal As String, ByVal SumToCorp As String, ByVal TransDomCurr As String, ByVal UseHistRate As String,
                                ByVal PrOpenItem As String, ByVal PrOpenPay As String, ByVal HidePaid As String, ByVal SortByCurr As String, ByVal ArSortBy As String, ByVal AgeBuckets As String,
                                ByVal InvDue As String, ByVal AgeDays1 As String, ByVal AgeDesc1 As String, ByVal AgeDays2 As String, ByVal AgeDesc2 As String,
                                ByVal AgeDays3 As String, ByVal AgeDesc3 As String, ByVal AgeDays4 As String, ByVal AgeDesc4 As String, ByVal AgeDays5 As String, ByVal AgeDesc5 As String,
                                ByVal SiteGroup As String, ByVal DisplayHeader As String, ByVal ConsolidateCustomers As String, ByVal IncludeEstCurrGainLossAmtsInTotals As String, ByVal Site As String, ByVal SessionID As String) As DataTable

            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCustomer As LoadCollectionResponseData
            Dim oCo As LoadCollectionResponseData
            Dim oCustAddr As LoadCollectionResponseData

            Dim sRecordFilter As String


            Dim ReportInputParametersArray As String() = New List(Of String) From {AgingDate, CutoffDate, AgingDateOffset, CutoffDateOffset, StateCycle, ShowActive,
                           BegSlsman, EndSlsman, CustomerStarting, CustomerEnding, NameStarting, NameEnding, CurCodeStarting, CurCodeEnding, PrZeroBal, CreditHold,
                           PrCreditBal, SumToCorp, TransDomCurr, UseHistRate, PrOpenItem, PrOpenPay, HidePaid, SortByCurr, ArSortBy, AgeBuckets, InvDue,
                           AgeDays1, AgeDesc1, AgeDays2, AgeDesc2, AgeDays3, AgeDesc3, AgeDays4, AgeDesc4, AgeDays5, AgeDesc5, SiteGroup, DisplayHeader,
                           ConsolidateCustomers, IncludeEstCurrGainLossAmtsInTotals, Site, SessionID}.ToArray

            Dim sResultSetColumnsList As String
            sResultSetColumnsList = "TcSortCurrCode,CurrencyFormat,CurrencyPlaces,TotCurrencyFormat,TotCurrencyPlaces,TcSortBy,TcCustNum,TcCustName,TcCity,TcState,"
            sResultSetColumnsList = sResultSetColumnsList & "TcSite,TcSiteName,TcContact,TcPhone,TcTempTermsCode,TcCustType,TcCreditLimit,TcCredhold,TcCurrCode,TcArtranType,"
            sResultSetColumnsList = sResultSetColumnsList & "StdCh,TcArtranInvSeq,TcArtranDate,TcArtranDueDate,TcAmtTran,TcArtranExchRate,TcArtranCurrCode,CustAmtTran,TcAmtTemp,PAgeDesc,"
            sResultSetColumnsList = sResultSetColumnsList & "PAgeDescNum,TcApprovalStatus,StdCh1,TcCustCurrCode,OrderByDate,ApplyToInv,TcArTranIvDate,TcArTranChkSeq,InvNum,TotalDays,"
            sResultSetColumnsList = sResultSetColumnsList & "THPaymentNumber,Seq,Group1,Group2,Group3,Gp3SiteTotalOriginal,Gp3SiteTotalAgeDesc1,Gp3SiteTotalAgeDesc2,Gp3SiteTotalAgeDesc3,Gp3SiteTotalAgeDesc4,"
            sResultSetColumnsList = sResultSetColumnsList & "Gp3SiteTotalAgeDesc5,Gp3CustomerTotalOriginal,Gp3CustomerTotalAgeDesc1,Gp3CustomerTotalAgeDesc2,Gp3CustomerTotalAgeDesc3,Gp3CustomerTotalAgeDesc4,Gp3CustomerTotalAgeDesc5,Gp2CustomerTotalOriginal,Gp2CustomerTotalAgeDesc1,Gp2CustomerTotalAgeDesc2,"
            sResultSetColumnsList = sResultSetColumnsList & "Gp2CustomerTotalAgeDesc3,Gp2CustomerTotalAgeDesc4,Gp2CustomerTotalAgeDesc5,Gp2SiteTotalOriginal,Gp2SiteTotalAgeDesc1,Gp2SiteTotalAgeDesc2,Gp2SiteTotalAgeDesc3,Gp2SiteTotalAgeDesc4,Gp2SiteTotalAgeDesc5,TotalOriginal,"
            sResultSetColumnsList = sResultSetColumnsList & "TotalAgeDesc1,TotalAgeDesc2,TotalAgeDesc3,TotalAgeDesc4,TotalAgeDesc5,GrandTotalOriginal,GrandTotalAgeDesc1,GrandTotalAgeDesc2,GrandTotalAgeDesc3,GrandTotalAgeDesc4,"
            sResultSetColumnsList = sResultSetColumnsList & "GrandTotalAgeDesc5,IsCurrCodeDistinct"


            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SL.SLAccountsReceivableAgingReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_AccountsReceivableAgingSp"

            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",CustPo,ShipToCity,Bucket1Amt,Bucket2Amt,Bucket3Amt,Bucket4Amt,Bucket5Amt"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(", ")  'sResultSetColumnsList.Split(CType(",", Char()))                    

                'Add columns to Datatable corresponding each column out of the result set.
                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                Next iIndex

                'Now Prepare the data table with Data. Loop on Standard Result Rows and add data rows and data to the data table.
                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    oResultDataRow = oResultDataTable.NewRow()
                    For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 8 'i.e. Ignoring the 9 custom fields. This will be populated later.
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                    Next iColumnIndex
                    oResultDataTable.Rows.Add(oResultDataRow)
                Next iRowIndex

                'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
                For Each oRow As DataRow In oResultDataTable.Rows
                    If String.Equals(oRow.Item("PAgeDescNum").ToString, "1") Then
                        oRow.Item("Bucket1Amt") = oRow.Item("TcAmtTemp")
                    End If
                    If String.Equals(oRow.Item("PAgeDescNum").ToString, "2") Then
                        oRow.Item("Bucket2Amt") = oRow.Item("TcAmtTemp")
                    End If
                    If String.Equals(oRow.Item("PAgeDescNum").ToString, "3") Then
                        oRow.Item("Bucket3Amt") = oRow.Item("TcAmtTemp")
                    End If
                    If String.Equals(oRow.Item("PAgeDescNum").ToString, "4") Then
                        oRow.Item("Bucket4Amt") = oRow.Item("TcAmtTemp")
                    End If
                    If String.Equals(oRow.Item("PAgeDescNum").ToString, "5") Then
                        oRow.Item("Bucket5Amt") = oRow.Item("TcAmtTemp")
                    End If

                    sRecordFilter = String.Format("InvNum = {0}", SqlLiteral.Format(oRow.Item("InvNum")))
                    oCustomer = Me.Context.Commands.LoadCollection("SLInvHdrs", "CustPo", sRecordFilter, "", 0)
                    If oCustomer.Items.Count = 1 Then
                        oRow.Item("CustPo") = oCustomer(0, "CustPo").Value
                    End If
                    oCustomer = Nothing
                    sRecordFilter = String.Format("InvNum = {0}", SqlLiteral.Format(oRow.Item("InvNum")))
                    oCustomer = Me.Context.Commands.LoadCollection("SLInvItemAlls", "CoNum", sRecordFilter, "", 0)
                    If oCustomer.Items.Count >= 1 Then
                        If Not String.IsNullOrEmpty(oCustomer(0, "CoNum").Value) Then
                            sRecordFilter = String.Format("CoNum = {0}", SqlLiteral.Format(oCustomer(0, "CoNum").Value))
                            oCo = Me.Context.Commands.LoadCollection("SLCos", "CustNum, CustSeq", sRecordFilter, "", 0)
                            If oCo.Items.Count >= 1 Then

                                If Not String.IsNullOrEmpty(oCo(0, "CustNum").Value) Then
                                    sRecordFilter = String.Format("CustNum = {0} and CustSeq = {1}", SqlLiteral.Format(oCo(0, "CustNum").Value), SqlLiteral.Format(oCo(0, "CustSeq").Value))
                                    'row.Item("ShipToCity") = sRecordFilter
                                    oCustAddr = Me.Context.Commands.LoadCollection("SLCustAddrs", "City", sRecordFilter, "", 0)
                                    If oCustAddr.Items.Count >= 1 Then
                                        oRow.Item("ShipToCity") = oCustAddr(0, "City").Value
                                    End If
                                End If
                            End If
                        End If
                    End If

                    oCustomer = Nothing
                    oCo = Nothing
                    oCustAddr = Nothing

                Next oRow

            End If


            LoadAccountsReceivableAging = oResultDataTable

        End Function



    End Class
End Namespace