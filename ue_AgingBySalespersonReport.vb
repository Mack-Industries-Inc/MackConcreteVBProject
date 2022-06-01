Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data

Imports Mongoose.Core.Common
Imports Mongoose.Core.DataAccess
Imports Mongoose.IDO
Imports Mongoose.IDO.Protocol
Imports Microsoft.VisualBasic
Imports Mongoose.Scripting
Imports Mongoose.Core
Namespace ue_AgingBySalespersonReport
    ' <IDOExtensionClass( "ue_AgingBySalespersonReport" )>
    Partial Public Class ue_AgingBySalespersonReport : Inherits IDOExtensionClass
        Public Function ue_Rpt_AccountsReceivableAgingSp(ByVal AgingDate As String, ByVal CutoffDate As String, ByVal AgingDateOffset As String, ByVal CutoffDateOffset As String,
                                ByVal StateCycle As String, ByVal ShowActive As String, ByVal BegSlsman As String, ByVal EndSlsman As String, ByVal CustomerStarting As String, ByVal CustomerEnding As String,
                                ByVal NameStarting As String, ByVal NameEnding As String, ByVal CurCodeStarting As String, ByVal CurCodeEnding As String, ByVal PrZeroBal As String,
                                ByVal CreditHold As String, ByVal PrCreditBal As String, ByVal SumToCorp As String, ByVal TransDomCurr As String, ByVal UseHistRate As String,
                                ByVal PrOpenItem As String, ByVal PrOpenPay As String, ByVal HidePaid As String, ByVal SortByCurr As String, ByVal ArSortBy As String, ByVal AgeBuckets As String,
                                ByVal InvDue As String, ByVal AgeDays1 As String, ByVal AgeDesc1 As String, ByVal AgeDays2 As String, ByVal AgeDesc2 As String,
                                ByVal AgeDays3 As String, ByVal AgeDesc3 As String, ByVal AgeDays4 As String, ByVal AgeDesc4 As String, ByVal AgeDays5 As String, ByVal AgeDesc5 As String,
                                ByVal SiteGroup As String, ByVal DisplayHeader As String, ByVal ConsolidateCustomers As String, ByVal IncludeEstCurrGainLossAmtsInTotals As String, ByVal pSite As String, ByVal pProcessId As String
                                ) As DataTable

            Dim sResultSetColumnsList As String
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim sRecordFilter As String
            Dim oInvHdr As LoadCollectionResponseData
            Dim FinalResultDataTable As DataTable = New DataTable
            Dim CoNum As String
            Dim CustNum As String
            Dim CustSeq As String
            Dim oCustomer As LoadCollectionResponseData

            Dim ReportInputParametersArray As String() = New List(Of String) From {
            AgingDate, CutoffDate, AgingDateOffset, CutoffDateOffset, StateCycle, ShowActive, BegSlsman, EndSlsman, CustomerStarting, CustomerEnding,
            NameStarting, NameEnding, CurCodeStarting, CurCodeEnding, PrZeroBal, CreditHold, PrCreditBal, SumToCorp, TransDomCurr, UseHistRate,
            PrOpenItem, PrOpenPay, HidePaid, SortByCurr, ArSortBy, AgeBuckets, InvDue, AgeDays1, AgeDesc1, AgeDays2, AgeDesc2, AgeDays3, AgeDesc3,
            AgeDays4, AgeDesc4, AgeDays5, AgeDesc5, SiteGroup, DisplayHeader, ConsolidateCustomers, IncludeEstCurrGainLossAmtsInTotals,
            pSite, pProcessId}.ToArray

            sResultSetColumnsList = "TcSortCurrCode,CurrencyFormat,CurrencyPlaces,TotCurrencyFormat,TotCurrencyPlaces,TcSortBy,TcCustNum,TcCustName,TcCity,TcState,TcSite,"
            sResultSetColumnsList = sResultSetColumnsList & "TcSiteName,TcContact,TcPhone,TcTempTermsCode,TcCustType,TcCreditLimit,TcCredhold,TcCurrCode,TcArtranType,StdCh,TcArtranInvSeq,TcArtranDate,"
            sResultSetColumnsList = sResultSetColumnsList & "TcArtranDueDate,TcAmtTran,TcArtranExchRate,TcArtranCurrCode,CustAmtTran,TcAmtTemp,PAgeDesc,PAgeDescNum,TcApprovalStatus,StdCh1,TcCustCurrCode,"
            sResultSetColumnsList = sResultSetColumnsList & "OrderByDate,ApplyToInv,TcArTranIvDate,TcArTranChkSeq,InvNum,TotalDays,THPaymentNumber,Seq,Group1,Group2,Group3,"
            sResultSetColumnsList = sResultSetColumnsList & "GrandTotalAgeDesc1,Gp3SiteTotalAgeDesc1,Gp3SiteTotalAgeDesc2,Gp3SiteTotalAgeDesc3,Gp3SiteTotalAgeDesc4,Gp3SiteTotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "Gp3CustomerTotalOriginal,Gp3CustomerTotalAgeDesc1,Gp3CustomerTotalAgeDesc2,Gp3CustomerTotalAgeDesc3,Gp3CustomerTotalAgeDesc4,Gp3CustomerTotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "Gp2CustomerTotalOriginal,Gp2CustomerTotalAgeDesc1,Gp2CustomerTotalAgeDesc2,Gp2CustomerTotalAgeDesc3,Gp2CustomerTotalAgeDesc4,Gp2CustomerTotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "Gp2SiteTotalOriginal,Gp2SiteTotalAgeDesc1,Gp2SiteTotalAgeDesc2,Gp2SiteTotalAgeDesc3,Gp2SiteTotalAgeDesc4,Gp2SiteTotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "TotalOriginal,TotalAgeDesc1,TotalAgeDesc2,TotalAgeDesc3,TotalAgeDesc4,TotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "GrandTotalOriginal,GrandTotalAgeDesc1,GrandTotalAgeDesc2,GrandTotalAgeDesc3,GrandTotalAgeDesc4,GrandTotalAgeDesc5,"
            sResultSetColumnsList = sResultSetColumnsList & "IsCurrCodeDistinct,ProcessId,TcArTranTypeGroup,TcAmtTempAge1,TcAmtTempAge2,TcAmtTempAge3,TcAmtTempAge4,TcAmtTempAge5"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SL.SLAccountsReceivableAgingReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_AccountsReceivableAgingSp"

            For iIndex = 0 To ReportInputParametersArray.Length - 1
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",Slsman,CoNum,Bucket1Desc,Bucket2Desc,Bucket3Desc,Bucket4Desc,Bucket5Desc"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

                'Add columns to Datatable corresponding each column out of the result set.
                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    If iIndex = 45 Then
                        oResultDataTable.Columns.Add("GrandTotalDesc")
                    Else
                        oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                    End If
                Next iIndex

                'Now Prepare the data table with Data. Loop on Standard Result Rows and add data rows and data to the data table.
                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    oResultDataRow = oResultDataTable.NewRow()
                    For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 8 'i.e. Ignoring the 1 custom fields. This will be populated later.
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                    Next iColumnIndex
                    oResultDataTable.Rows.Add(oResultDataRow)
                Next iRowIndex

                'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
                For Each oRow As DataRow In oResultDataTable.Rows
                    sRecordFilter = String.Format("InvNum = {0}", SqlLiteral.Format(oRow.Item("InvNum")))
                    oInvHdr = Me.Context.Commands.LoadCollection("SLInvHdrs", "CoNum,CustNum,CustSeq", sRecordFilter, "", 0)
                    If oInvHdr.Items.Count > 0 Then
                        CoNum = oInvHdr(0, "CoNum").Value
                        oRow.Item("CoNum") = CoNum
                        CustNum = oInvHdr(0, "CustNum").Value
                        CustSeq = oInvHdr(0, "CustSeq").Value
                        sRecordFilter = String.Format("CoNum = {0}", SqlLiteral.Format(CoNum))
                        oCustomer = Me.Context.Commands.LoadCollection("SLCos", "Slsman", sRecordFilter, "", 0)
                        If oCustomer.Items.Count > 0 Then
                            oRow.Item("Slsman") = oCustomer(0, "Slsman").Value
                        End If
                    End If

                    oRow.Item("Bucket1Desc") = AgeDesc1
                    oRow.Item("Bucket2Desc") = AgeDesc2
                    oRow.Item("Bucket3Desc") = AgeDesc3
                    oRow.Item("Bucket4Desc") = AgeDesc4
                    oRow.Item("Bucket5Desc") = AgeDesc5
                Next oRow
            End If
            'Sort by Slsman
            If oResultDataTable.Rows.Count > 0 Then
                Dim dv As DataView = oResultDataTable.DefaultView
                dv.Sort = "Slsman"
                FinalResultDataTable = dv.ToTable()
            Else
                FinalResultDataTable = oResultDataTable
            End If

            Return FinalResultDataTable
        End Function

    End Class
End Namespace

