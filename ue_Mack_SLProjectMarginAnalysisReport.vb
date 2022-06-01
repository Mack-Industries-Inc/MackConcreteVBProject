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

Namespace ue_Mack_SLProjectMarginAnalysisReport
    ' <IDOExtensionClass( "ue_Mack_SLProjectMarginAnalysisReport" )>
    Partial Public Class ue_Mack_SLProjectMarginAnalysisReport : Inherits IDOExtensionClass
        Public Function Mack_Rpt_ProjectMarginAnalysisSp(ByVal StartingProject As String, ByVal EndingProject As String, ByVal DisplayHeader As String, ByVal pSite As String,
                                ByVal pStatus As String) As DataTable

            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oProject As LoadCollectionResponseData
            Dim oCoitem As LoadCollectionResponseData
            Dim oSlsman As LoadCollectionResponseData
            Dim sRecordFilter As String
            Dim DAmount As Double
            Dim CostToComplete As Double
            Dim CostToDate As Double

            Dim ReportInputParametersArray As String() = New List(Of String) From {StartingProject, EndingProject, DisplayHeader, pSite}.ToArray

            Dim sResultSetColumnsList As String
            sResultSetColumnsList = "Seq,ProjNum,CostToDate,CostToComplete,ProjRev,ProjectedMargin,ProjectedPercentage,RevPtd,CurrentMargin,"
            sResultSetColumnsList = sResultSetColumnsList & "CurrentPercentage,ActVsProjected,ProjRowPointer"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SL.SLProjectMarginAnalysisReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = "ProjNum"
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_ProjectMarginAnalysisSp"

            For iIndex = 0 To ReportInputParametersArray.Length - 1
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",DAmount,Slsman,SlsmanName,Budget,Forecast"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

                'Add columns to Datatable corresponding each column out of the result set.
                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                Next iIndex

                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    sRecordFilter = String.Format("ProjNum = {0}", SqlLiteral.Format(oRptStdResultResponse.Items(iRowIndex).PropertyValues(1).ToString()))
                    oProject = Me.Context.Commands.LoadCollection("SLProjs", "Stat", sRecordFilter, "", 0)
                    If oProject.Items.Count = 1 Then
                        If pStatus.Contains(oProject(0, "Stat").ToString()) Then
                            oResultDataRow = oResultDataTable.NewRow()
                            For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 6
                                oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                            Next iColumnIndex
                            oResultDataTable.Rows.Add(oResultDataRow)
                        End If
                    End If
                Next iRowIndex

                For Each oRow As DataRow In oResultDataTable.Rows
                    sRecordFilter = String.Format("CoNum = {0} and Type in ('P','D','C')", SqlLiteral.Format(oRow.Item("ProjNum").ToString))
                    oCoitem = Me.Context.Commands.LoadCollection("SLArtrans", "Type, CoNum, Amount", sRecordFilter, "", 0)
                    DAmount = 0
                    If oCoitem.Items.Count <> 0 Then
                        For iIndex = 0 To oCoitem.Items.Count - 1
                            If String.Equals(oCoitem(iIndex, "Type").ToString, "P") OrElse String.Equals(oCoitem(iIndex, "Type").ToString, "C") Then
                                DAmount = DAmount + oCoitem(iIndex, "Amount").GetValue(Of Double)()
                            Else
                                DAmount = DAmount - oCoitem(iIndex, "Amount").GetValue(Of Double)()
                            End If
                        Next
                    End If
                    oRow.Item("DAmount") = DAmount

                    Dim ProjRev As Double = Convert.ToDouble(oRow.Item("ProjRev"))

                    sRecordFilter = String.Format("ProjNum = {0}", SqlLiteral.Format(oRow.Item("ProjNum").ToString))
                    oProject = Me.Context.Commands.LoadCollection("SLProjs", "proUf_Slsman,DerTotPlanCost,DerTotFcstCost", sRecordFilter, "", 0)
                    If oProject.Items.Count = 1 Then
                        Dim Slsman As String = oProject(0, "proUf_Slsman").ToString
                        Dim Budget As Double = oProject(0, "DerTotPlanCost").GetValue(Of Double)
                        Dim ProjectedMargin As Double = ProjRev - Budget
                        Dim ProjectedPercentage As Double = 0
                        If ProjRev <> 0 Then
                            ProjectedPercentage = (ProjRev - Budget) / ProjRev * 100
                        End If

                        CostToDate = Convert.ToDouble(oRow.Item("CostToDate"))
                        CostToComplete = Budget - CostToDate

                        oRow.Item("Slsman") = Slsman
                        oRow.Item("Budget") = Convert.ToString(Budget)
                        oRow.Item("Forecast") = oProject(0, "DerTotFcstCost")
                        oRow.Item("ProjectedMargin") = Convert.ToString(ProjectedMargin)
                        oRow.Item("ProjectedPercentage") = Convert.ToString(ProjectedPercentage)
                        oRow.Item("CostToComplete") = Convert.ToString(CostToComplete)

                        sRecordFilter = String.Format("Slsman = {0}", SqlLiteral.Format(Slsman))
                        oSlsman = Me.Context.Commands.LoadCollection("SLSlsmans", "DerSalesmanName", sRecordFilter, "", 0)
                        If oSlsman.Items.Count = 1 Then
                            oRow.Item("SlsmanName") = oSlsman(0, "DerSalesmanName")
                        End If
                    End If
                Next oRow

            End If

            Mack_Rpt_ProjectMarginAnalysisSp = oResultDataTable

        End Function

    End Class
End Namespace