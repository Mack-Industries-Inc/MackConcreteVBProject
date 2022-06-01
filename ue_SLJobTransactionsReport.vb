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


Namespace ue_SLJobTransactionsReport
    ' <IDOExtensionClass( "ue_SLJobTransactionsReport" )>
    Partial Public Class ue_SLJobTransactionsReport : Inherits IDOExtensionClass
        Public Function ue_Rpt_JobTransactionsSp(ByVal TransactionType As String, ByVal PayType As String, ByVal Posted As String, ByVal EmployeeType As String,
                                ByVal ShowDetail As String, ByVal BackflushTransaction As String, ByVal EmployeeStarting As String, ByVal EmployeeEnding As String,
                                ByVal JobStarting As String, ByVal JobEnding As String, ByVal SuffixStarting As String, ByVal SuffixEnding As String,
                                ByVal TransactionDateStarting As String, ByVal TransactionDateEnding As String,
                                ByVal TransactionNumberStarting As String, ByVal TransactionNumberEnding As String,
                                ByVal ShiftStarting As String, ByVal ShiftEnding As String,
                                ByVal ReasonStarting As String, ByVal ReasonEnding As String,
                                ByVal UserInitialStarting As String, ByVal UserInitialEnding As String,
                                ByVal ResourceStarting As String, ByVal ResourceEnding As String,
                                ByVal SortByTEJ As String, ByVal TransactionDateStartingOffset As String, ByVal TransactionDateEndingOffset As String,
                                ByVal ViewCost As String, ByVal DisplayHeader As String, ByVal PMessageLanguage As String,
                                ByVal pSite As String, ByVal BGUser As String) As DataTable

            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oItem As LoadCollectionResponseData
            Dim Dept As String

            Dim ReportInputParametersArray As String() = New List(Of String) From {TransactionType, PayType, Posted, EmployeeType, ShowDetail,
                BackflushTransaction, EmployeeStarting, EmployeeEnding, JobStarting, JobEnding, SuffixStarting, SuffixEnding, TransactionDateStarting,
                TransactionDateEnding, TransactionNumberStarting, TransactionNumberEnding, ShiftStarting, ShiftEnding, ReasonStarting, ReasonEnding,
                UserInitialStarting, UserInitialEnding, ResourceStarting, ResourceEnding, SortByTEJ, TransactionDateStartingOffset, TransactionDateEndingOffset,
                ViewCost, DisplayHeader, PMessageLanguage, pSite, BGUser}.ToArray

            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String
            Dim oJob As LoadCollectionResponseData

            sResultSetColumnsList = "TranNum,Posted,TransDate,TypeDesc,EmpNum,JobRate,Shift,Job,Suffix,OperNum,IndCode,Backflush,"
            sResultSetColumnsList = sResultSetColumnsList & "ReasonCode,Hours,Tot,Completed,Scrapped,MoveTo,LocOper,UserCode,CloseJob,CompleteOp,StartTime,"
            sResultSetColumnsList = sResultSetColumnsList & "EndTime,PayType,PayRate,JobCostRate,TotalCost,ToLoc,EmployeeName,Item,GroupField,ItemDesc,RESID,"
            sResultSetColumnsList = sResultSetColumnsList & "DESCR,TranNumString"


            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SLJobTransactionsReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_JobTransactionsSp"


            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",Dept"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

                'Add columns to Datatable corresponding each column out of the result set.
                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                Next iIndex

                'Now Prepare the data table with Data. Loop on Standard Result Rows and add data rows and data to the data table.
                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    oResultDataRow = oResultDataTable.NewRow()
                    For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 2 'i.e. Ignoring the 1 custom fields. This will be populated later.
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                    Next iColumnIndex
                    oResultDataTable.Rows.Add(oResultDataRow)
                Next iRowIndex

                'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
                For Each oRow As DataRow In oResultDataTable.Rows
                    Dept = ""
                    sRecordFilter = String.Format("Job = {0} and Suffix = {1} and OperNum = {2}", SqlLiteral.Format(oRow.Item("Job")), SqlLiteral.Format(oRow.Item("Suffix")), SqlLiteral.Format(oRow.Item("OperNum")))
                    oJob = Me.Context.Commands.LoadCollection("SLJobRoutes", "WcDept", sRecordFilter, "", 0)
                    If oJob.Items.Count = 1 Then
                        Dept = oJob(0, "WcDept").Value
                    End If

                    oRow.Item("Dept") = Dept

                Next oRow

            End If

            ue_Rpt_JobTransactionsSp = oResultDataTable

        End Function

    End Class
End Namespace