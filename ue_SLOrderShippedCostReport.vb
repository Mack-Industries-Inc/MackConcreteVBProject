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

Namespace ue_SLOrderShippedCostReport
    ' <IDOExtensionClass( "ue_SLOrderShippedCostReport" )>
    Partial Public Class ue_SLOrderShippedCostReport : Inherits IDOExtensionClass
        Public Function ue_Rpt_OrderShippedCostSp(ByVal TranslateToDomestic As String, ByVal OrderStarting As String, ByVal OrderEnding As String,
                                               ByVal CustomerStarting As String, ByVal CustomerEnding As String, ByVal ItemStarting As String, ByVal ItemEnding As String,
                                               ByVal DateShippedStarting As String, ByVal DateShippedEnding As String, ByVal DateShippedStartingOffset As String, ByVal DateShippedEndingOffset As String,
                                               ByVal DisplayReportHeader As String, ByVal pSite As String) As DataTable
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim ReportInputParametersArray As String() = New List(Of String) From {TranslateToDomestic, OrderStarting, OrderEnding, CustomerStarting, CustomerEnding,
                ItemStarting, ItemEnding, DateShippedStarting, DateShippedEnding, DateShippedStartingOffset, DateShippedEndingOffset, DisplayReportHeader, pSite}.ToArray
            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String
            Dim oCo As LoadCollectionResponseData
            Dim oCustomer As LoadCollectionResponseData
            Dim oCoitem As LoadCollectionResponseData
            Dim oShip As LoadCollectionResponseData

            sResultSetColumnsList = "CONumber,CustomerNumber,OrderDate,CoStatus,CustomerName,TotalPrice,CoLine,CoRelease,"
            sResultSetColumnsList = sResultSetColumnsList & "Item,UM,QtyOrdered,COItemPrice,DiscountPrice,SumQtyShipped,SumLaborCost,SumMatlCost,"
            sResultSetColumnsList = sResultSetColumnsList & "SumFOvhdCost,SumVOvhdCost,SumOutCost,QtyUnitFormat,PlacesQtyUnit,TotalCost,ExtendedDiscPrice,"
            sResultSetColumnsList = sResultSetColumnsList & "Margin,MixedLineRelease"


            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SLOrderShippedCostReport"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "Rpt_OrderShippedCostSp"


            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",CY,JobName,ItemDesc,ProdCode,BillToName,ShipDate"
                Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

                'Add columns to Datatable corresponding each column out of the result set.
                For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                    oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
                Next iIndex

                'Now Prepare the data table with Data. Loop on Standard Result Rows and add data rows and data to the data table.
                For iRowIndex As Integer = 0 To oRptStdResultResponse.Items.Count - 1
                    oResultDataRow = oResultDataTable.NewRow()
                    For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 7 'i.e. Ignoring the 1 custom fields. This will be populated later.
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oRptStdResultResponse.Items(iRowIndex).PropertyValues(iColumnIndex).ToString()
                    Next iColumnIndex
                    oResultDataTable.Rows.Add(oResultDataRow)
                Next iRowIndex

                'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
                For Each oRow As DataRow In oResultDataTable.Rows
                    CY = 0
                    sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(oRow.Item("Item")))
                    oItem = Me.Context.Commands.LoadCollection("SLItems", "WeightUnits, UnitWeight, ProductCode", sRecordFilter, "", 0)
                    If oItem.Items.Count = 1 Then
                        If String.Equals(oRow.Item("UM"), "EA") Then

                            If String.Equals(oItem(0, "WeightUnits").Value, "CY") Then
                                CY = Convert.ToDouble(oItem(0, "UnitWeight").Value) * Convert.ToDouble(oRow.Item("SumQtyShipped"))
                            End If
                        End If
                        oRow.Item("ProdCode") = oItem(0, "ProductCode").Value
                    End If
                    oRow.Item("CY") = CY

                    sRecordFilter = String.Format("CoNum = {0}", SqlLiteral.Format(oRow.Item("CONumber")))
                    oCo = Me.Context.Commands.LoadCollection("SLCos", "coUf_JobProjName", sRecordFilter, "", 0)
                    If oCo.Items.Count = 1 Then
                        oRow.Item("JobName") = oCo(0, "coUf_JobProjName").Value
                    End If

                    sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = {2}", SqlLiteral.Format(oRow.Item("CONumber")),
                SqlLiteral.Format(oRow.Item("CoLine")), SqlLiteral.Format(oRow.Item("CoRelease")))
                    oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "Description", sRecordFilter, "", 0)
                    If oCoitem.Items.Count = 1 Then
                        oRow.Item("ItemDesc") = oCoitem(0, "Description").Value
                    End If

                    sRecordFilter = String.Format("CustNum = {0} and CustSeq = 0", SqlLiteral.Format(oRow.Item("CustomerNumber")))
                    oCustomer = Me.Context.Commands.LoadCollection("SLCustAddrs", "Name", sRecordFilter, "", 0)
                    If oCustomer.Items.Count = 1 Then
                        oRow.Item("BillToName") = oCustomer(0, "Name").Value
                    End If

                    sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = {2}", SqlLiteral.Format(oRow.Item("CONumber")),
                                                  SqlLiteral.Format(oRow.Item("CoLine")), SqlLiteral.Format(oRow.Item("CoRelease")))
                    oShip = Me.Context.Commands.LoadCollection("SLCoShips", "ShipDate", sRecordFilter, "ShipDate Desc", 0)
                    If oShip.Items.Count = 1 Then
                        oRow.Item("ShipDate") = oShip(0, "ShipDate").Value
                    End If
                Next oRow

            End If


            ue_Rpt_OrderShippedCostSp = oResultDataTable

        End Function

    End Class
End Namespace