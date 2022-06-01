'IDO
'System.dll, System.Linq.dll, System.Core.dll, System.Transactions.dll, System.Net.dll, System.Web.dll, System.Net.Http.dll, System.Data.dll, System.Data.SqlClient.dll, System.IO.dll, Mongoose.IDO.DataAccess.dll
'Make sure change this line
'For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 2 'i.e. Ignoring the 1 custom fields. This will be populated later. Make sure change this line
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

Namespace ue_Mack_SLExecutives
    ' <IDOExtensionClass( "ue_Mack_SLExecutives" )>
    Partial Public Class ue_Mack_SLExecutives : Inherits IDOExtensionClass
        Public Structure ShipmentStruct
            Public Property ShipSite As String
            Public Property CoNum As String
            Public Property CoLine As String
            Public Property CoRelease As String
            Public Property Item As String
            Public Property ProductCode As String
            Public Property Slsman As String
            Public Property CY As Double
            Public Property ItemDescription As String
        End Structure

        Public Function Mack_CLM_ExecutiveShipmentRevenueSp(ByVal View As String, ByVal SiteGroup As String, ByVal DateStarting As String, ByVal DateEnding As String,
                                ByVal FilterString As String) As DataTable
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oFinalResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim AllView As String = "A"
            Dim ReportInputParametersArray As String() = New List(Of String) From {AllView, SiteGroup, DateStarting, DateEnding, FilterString}.ToArray
            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String
            Dim lstShipmentStruct As List(Of ShipmentStruct) = New List(Of ShipmentStruct)
            Dim ShipSite As String
            Dim CoNum As String
            Dim CoLine As String
            Dim CoRelease As String
            Dim Item As String
            Dim ProductCode As String
            Dim oShipmentStruct As ShipmentStruct = Nothing
            Dim ItemDescription As String

            oResultDataTable = GetShipmentRevenueSp(AllView, SiteGroup, DateStarting, DateEnding, FilterString)

            'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
            For Each oRow As DataRow In oResultDataTable.Rows
                sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = {2}", SqlLiteral.Format(oRow.Item("CoNum")), SqlLiteral.Format(oRow.Item("CoLine")), SqlLiteral.Format(oRow.Item("CoRelease")))
                oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "UM, QtyOrdered, UnitWeight", sRecordFilter, "", 0)
                CY = 0
                If oCoitem.Items.Count = 1 Then
                    If String.Equals(oCoitem(0, "UM").Value, "EA") Then
                        sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(oRow.Item("Item")))
                        oItem = Me.Context.Commands.LoadCollection("SLItems", "WeightUnits, UnitWeight", sRecordFilter, "", 0)
                        If oItem.Items.Count = 1 Then
                            If String.Equals(oItem(0, "WeightUnits").Value, "CY") Then
                                CY = Convert.ToDouble(oCoitem(0, "QtyOrdered").Value) * Convert.ToDouble(oItem(0, "UnitWeight").Value)
                            End If
                        End If
                    End If
                End If
                oRow.Item("CY") = CY
                ShipSite = oRow.Item("ShipSite").ToString
                CoNum = oRow.Item("CoNum").ToString
                CoLine = oRow.Item("CoLine").ToString
                CoRelease = oRow.Item("CoRelease").ToString
                Item = oRow.Item("Item").ToString
                ProductCode = oRow.Item("DerProductCode").ToString
                ItemDescription = oRow.Item("ItDescription").ToString

                lstShipmentStruct.Add(New ShipmentStruct With {.ShipSite = ShipSite, .CoNum = CoNum, .CoLine = CoLine, .CoRelease = CoRelease,
                                      .Item = Item, .ProductCode = ProductCode, .CY = CY, .ItemDescription = ItemDescription})
            Next oRow

            If String.Equals(View, "A") Then
                oFinalResultDataTable = oResultDataTable
            Else
                oFinalResultDataTable = GetShipmentRevenueSp(View, SiteGroup, DateStarting, DateEnding, FilterString)
            End If

            If lstShipmentStruct.Any And String.Equals(View, "S") Then
                Dim SgroupShipment = lstShipmentStruct.GroupBy(Function(group) group.ShipSite).Select(Function(group) New ShipmentStruct With {.ShipSite = group.Key,
                                  .CY = group.Sum(Function(sum) sum.CY)}).ToList

                For Each oRow As DataRow In oFinalResultDataTable.Rows
                    oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ShipSite, oRow.Item("ShipSite").ToString)).FirstOrDefault
                    oRow.Item("CY") = oShipmentStruct.CY
                Next oRow
            End If

            If lstShipmentStruct.Any And String.Equals(View, "P") Then
                Dim SgroupShipment = lstShipmentStruct.GroupBy(Function(group) group.ProductCode).Select(Function(group) New ShipmentStruct With {.ProductCode = group.Key,
                                  .CY = group.Sum(Function(sum) sum.CY)}).ToList

                For Each oRow As DataRow In oFinalResultDataTable.Rows
                    oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ProductCode, oRow.Item("DerProductCode").ToString)).FirstOrDefault
                    oRow.Item("CY") = oShipmentStruct.CY
                Next oRow
            End If

            If lstShipmentStruct.Any And String.Equals(View, "I") Then
                'Dim SgroupShipment = lstShipmentStruct.GroupBy(Function(group) group.Item).Select(Function(group) New ShipmentStruct With {.Item = group.Key,
                '.CY = group.Sum(Function(sum) sum.CY)}).ToList

                Dim SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.Item, Key group.ItemDescription}).
                        Select(Function(group) New ShipmentStruct With {.Item = group.Key.Item, .ItemDescription = group.Key.ItemDescription,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                For Each oRow As DataRow In oFinalResultDataTable.Rows
                    'oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.Item, oRow.Item("Item").ToString)).FirstOrDefault
                    oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.Item, oRow.Item("Item").ToString) AndAlso
                                                                   String.Equals(db.ItemDescription, oRow.Item("ItDescription").ToString)).FirstOrDefault

                    oRow.Item("CY") = oShipmentStruct.CY
                Next oRow
            End If

            Mack_CLM_ExecutiveShipmentRevenueSp = oFinalResultDataTable

        End Function

        Public Function GetShipmentRevenueSp(ByVal View As String, ByVal SiteGroup As String, ByVal DateStarting As String, ByVal DateEnding As String,
                                ByVal FilterString As String) As DataTable
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim ReportInputParametersArray As String() = New List(Of String) From {View, SiteGroup, DateStarting, DateEnding, FilterString}.ToArray
            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String


            sResultSetColumnsList = "CoOrigSite,CoNum,CoLine,CoRelease,Item,ItDescription,DerProductCode,ShipSite,DerTotCost,DerNetPrice,"
            sResultSetColumnsList = sResultSetColumnsList & "DerMargin,CoCustNum,Adr0Name,UbGrandTotCost,UbGrandTotPrice,UbGrandTotMargin,DomCurrCode"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SLExecutives"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "CLM_ExecutiveShipmentRevenueSp"


            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)

            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",CY"
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
                If String.Equals(View, "A") Then
                    For Each oRow As DataRow In oResultDataTable.Rows
                        sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = {2}", SqlLiteral.Format(oRow.Item("CoNum")), SqlLiteral.Format(oRow.Item("CoLine")), SqlLiteral.Format(oRow.Item("CoRelease")))
                        oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "UM, QtyOrdered, UnitWeight", sRecordFilter, "", 0)
                        CY = 0
                        If oCoitem.Items.Count = 1 Then
                            If String.Equals(oCoitem(0, "UM").Value, "EA") Then
                                sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(oRow.Item("Item")))
                                oItem = Me.Context.Commands.LoadCollection("SLItems", "WeightUnits, UnitWeight", sRecordFilter, "", 0)
                                If oItem.Items.Count = 1 Then
                                    If String.Equals(oItem(0, "WeightUnits").Value, "CY") Then
                                        CY = Convert.ToDouble(oCoitem(0, "QtyOrdered").Value) * Convert.ToDouble(oItem(0, "UnitWeight").Value)
                                    End If
                                End If
                            End If
                        End If
                        oRow.Item("CY") = CY
                    Next oRow
                End If
            End If

            GetShipmentRevenueSp = oResultDataTable

        End Function

        Public Function Mack_CLM_ExecutiveLateOrdersSp(ByVal SiteGroup As String, ByVal DateStarting As String, ByVal DateEnding As String, ByVal FilterString As String) As DataTable
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim ReportInputParametersArray As String() = New List(Of String) From {SiteGroup, DateStarting, DateEnding, FilterString}.ToArray
            Dim sRecordFilter As String

            Dim sResultSetColumnsList As String
            sResultSetColumnsList = "CoOrigSite,CoNum,CoLine,CoRelease,DueDate,Item,ItDescription,DerProductCode,ShipSite,"
            sResultSetColumnsList = sResultSetColumnsList & "DerTotCost,DerNetPrice,DerMargin,CoCustNum,Adr0Name,UbGrandTotPrice,DomCurrCode"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SLExecutives"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "CLM_ExecutiveLateOrdersSp"


            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)
            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",CY"
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
                    sRecordFilter = String.Format("CoNum = {0} and CoLine = {1} and CoRelease = {2}", SqlLiteral.Format(oRow.Item("CoNum")), SqlLiteral.Format(oRow.Item("CoLine")), SqlLiteral.Format(oRow.Item("CoRelease")))
                    oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "UM, QtyOrdered, UnitWeight", sRecordFilter, "", 0)
                    CY = 0
                    If oCoitem.Items.Count = 1 Then
                        If String.Equals(oCoitem(0, "UM").Value, "EA") Then
                            sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(oRow.Item("Item")))
                            oItem = Me.Context.Commands.LoadCollection("SLItems", "WeightUnits, UnitWeight", sRecordFilter, "", 0)
                            If oItem.Items.Count = 1 Then
                                If String.Equals(oItem(0, "WeightUnits").Value, "CY") Then
                                    CY = Convert.ToDouble(oCoitem(0, "QtyOrdered").Value) * Convert.ToDouble(oItem(0, "UnitWeight").Value)
                                End If
                            End If
                        End If
                    End If
                    oRow.Item("CY") = CY
                Next oRow

            End If

            Mack_CLM_ExecutiveLateOrdersSp = oResultDataTable

        End Function


        Public Function GetOrderBookingsSp(ByVal View As String, ByVal Detail As String, ByVal SiteGroup As String, ByVal DateStarting As String, ByVal DateEnding As String,
                                       ByVal FilterString As String) As DataTable
            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim ReportInputParametersArray As String() = New List(Of String) From {View, Detail, SiteGroup, DateStarting, DateEnding, FilterString}.ToArray
            Dim sRecordFilter As String

            Dim sResultSetColumnsList As String
            sResultSetColumnsList = "CoOrigSite,CoNum,Item,ItDescription,DerProductCode,ShipSite,DerTotCost,DerExtPrice,UbTotDiscount,"
            sResultSetColumnsList = sResultSetColumnsList & "DerNetPrice,DerMargin,CoCustNum,Adr0Name,CoSlsman,UbGrandTotPrice,UbGrandTotMargin,DomCurrCode"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.

            oRptStdResultRequest = New LoadCollectionRequestData()
            oRptStdResultRequest.IDOName = "SLExecutives"
            oRptStdResultRequest.RecordCap = 0
            oRptStdResultRequest.OrderBy = ""
            oRptStdResultRequest.PropertyList.SetProperties(sResultSetColumnsList)
            oRptStdResultRequest.CustomLoadMethod = New CustomLoadMethod()
            oRptStdResultRequest.CustomLoadMethod.Name = "CLM_ExecutiveOrderBookingsSp"


            For iIndex = 0 To ReportInputParametersArray.Length - 1 'Adding all the inputs to Request Object so that these will be passed to CLM to be invoked
                oRptStdResultRequest.CustomLoadMethod.Parameters.Add(ReportInputParametersArray(iIndex))
            Next iIndex

            'Call Standard Report First. And then let us prepare custom data.
            oRptStdResultResponse = Me.Context.Commands.LoadCollection(oRptStdResultRequest) 'IDOClient.LoadCollection(requestLoadCol)
            If oRptStdResultResponse.Items.Count > 0 Then 'i.e. If Standard Report Comes up with some data.
                'Appending custom columns to std columns list
                sResultSetColumnsList = sResultSetColumnsList & ",CY"
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


            End If

            GetOrderBookingsSp = oResultDataTable

        End Function

        Public Function Mack_CLM_ExecutiveOrderBookingsSp(ByVal View As String, ByVal Detail As String, ByVal SiteGroup As String, ByVal DateStarting As String, ByVal DateEnding As String,
                                        ByVal FilterString As String) As DataTable

            Dim oRptStdResultRequest As LoadCollectionRequestData
            Dim oRptStdResultResponse As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oFinalResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim CY As Double
            Dim ReportInputParametersArray As String() = New List(Of String) From {View, Detail, SiteGroup, DateStarting, DateEnding, FilterString}.ToArray
            Dim sRecordFilter As String
            Dim AllView As String = "S"
            Dim lstShipmentStruct As List(Of ShipmentStruct) = New List(Of ShipmentStruct)
            Dim ShipSite As String
            Dim CoNum As String
            Dim CoLine As String
            Dim CoRelease As String
            Dim Item As String
            Dim ProductCode As String
            Dim Slsman As String
            Dim oShipmentStruct As ShipmentStruct = Nothing
            Dim SgroupShipment As List(Of ShipmentStruct) = Nothing

            If String.Equals(View, "T") Then
                AllView = "T"
            Else
                AllView = "S"
            End If

            oResultDataTable = GetOrderBookingsSp(AllView, "1", SiteGroup, DateStarting, DateEnding, FilterString)

            If oResultDataTable IsNot Nothing Then
                'Now data table is already populated with Standard Data. Now Let us populate custom fields data also.
                For Each oRow As DataRow In oResultDataTable.Rows
                    sRecordFilter = String.Format("CoNum = {0}", SqlLiteral.Format(oRow.Item("CoNum")))
                    oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", "Item, UM, QtyOrdered, UnitWeight, CoSlsman, ShipSite", sRecordFilter, "", 0)
                    If oCoitem.Items.Count > 0 Then
                        For iIndex = 0 To oCoitem.Items.Count - 1
                            CY = 0
                            ProductCode = ""
                            sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(oCoitem(iIndex, "Item").Value))
                            oItem = Me.Context.Commands.LoadCollection("SLItems", "WeightUnits, UnitWeight, ProductCode", sRecordFilter, "", 0)
                            If String.Equals(oCoitem(iIndex, "UM").Value, "EA") Then
                                If oItem.Items.Count = 1 Then
                                    If String.Equals(oItem(0, "WeightUnits").Value, "CY") Then
                                        CY = Convert.ToDouble(oCoitem(iIndex, "QtyOrdered").Value) * Convert.ToDouble(oItem(0, "UnitWeight").Value)
                                    End If
                                    ProductCode = oItem(0, "ProductCode").Value
                                End If
                            End If

                            oRow.Item("CY") = CY
                            ShipSite = oCoitem(iIndex, "ShipSite").Value
                            CoNum = oRow.Item("CoNum").ToString
                            Item = oCoitem(iIndex, "Item").Value
                            Slsman = oCoitem(iIndex, "CoSlsman").Value

                            lstShipmentStruct.Add(New ShipmentStruct With {.ShipSite = ShipSite, .CoNum = CoNum,
                                          .Item = Item, .ProductCode = ProductCode, .CY = CY, .Slsman = Slsman})
                        Next
                    End If


                Next oRow
                If Not String.Equals(Detail, "1") Then
                    oFinalResultDataTable = GetOrderBookingsSp(View, Detail, SiteGroup, DateStarting, DateEnding, FilterString)
                Else
                    If String.Equals(View, "P") Then
                        oFinalResultDataTable = GetOrderBookingsSp(View, Detail, SiteGroup, DateStarting, DateEnding, FilterString)
                    Else
                        oFinalResultDataTable = oResultDataTable
                    End If
                End If

                If lstShipmentStruct.Any And String.Equals(View, "S") Then
                    If String.Equals(Detail, "1") Then
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.Slsman, Key group.CoNum}).
                        Select(Function(group) New ShipmentStruct With {.Slsman = group.Key.Slsman, .CoNum = group.Key.CoNum,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.Slsman, oRow.Item("CoSlsman").ToString) AndAlso
                                                                   String.Equals(db.CoNum, oRow.Item("CoNum").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY
                        Next oRow
                    Else
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.Slsman}).
                        Select(Function(group) New ShipmentStruct With {.Slsman = group.Key.Slsman,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.Slsman, oRow.Item("CoSlsman").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY
                        Next oRow
                    End If
                End If

                If lstShipmentStruct.Any And String.Equals(View, "T") Then
                    If String.Equals(Detail, "1") Then
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.ShipSite, Key group.CoNum}).
                        Select(Function(group) New ShipmentStruct With {.ShipSite = group.Key.ShipSite, .CoNum = group.Key.CoNum,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ShipSite, oRow.Item("ShipSite").ToString) AndAlso
                                                                   String.Equals(db.CoNum, oRow.Item("CoNum").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY
                        Next oRow
                    Else
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.ShipSite}).
                           Select(Function(group) New ShipmentStruct With {.ShipSite = group.Key.ShipSite,
                                         .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ShipSite, oRow.Item("ShipSite").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY
                        Next oRow
                    End If
                End If

                If lstShipmentStruct.Any And String.Equals(View, "P") Then
                    If String.Equals(Detail, "1") Then
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.ProductCode, Key group.Item}).
                        Select(Function(group) New ShipmentStruct With {.ProductCode = group.Key.ProductCode, .Item = group.Key.Item,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ProductCode, oRow.Item("DerProductCode").ToString) AndAlso
                                                                   String.Equals(db.Item, oRow.Item("Item").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY

                        Next oRow
                    Else
                        SgroupShipment = lstShipmentStruct.GroupBy(Function(group) New With {Key group.ProductCode}).
                        Select(Function(group) New ShipmentStruct With {.ProductCode = group.Key.ProductCode,
                                      .CY = group.Sum(Function(sum) sum.CY)}).ToList

                        For Each oRow As DataRow In oFinalResultDataTable.Rows
                            oShipmentStruct = SgroupShipment.Where(Function(db) String.Equals(db.ProductCode, oRow.Item("DerProductCode").ToString)).FirstOrDefault
                            oRow.Item("CY") = oShipmentStruct.CY

                        Next oRow

                    End If
                End If

            End If

            Mack_CLM_ExecutiveOrderBookingsSp = oFinalResultDataTable
        End Function
    End Class
End Namespace