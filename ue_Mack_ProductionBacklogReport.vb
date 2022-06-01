Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
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

Namespace ue_Mack_ProductionBacklogReport
    ' <IDOExtensionClass( "ue_Mack_ProductionBacklogReport" )>
    Partial Public Class ue_Mack_ProductionBacklogReport : Inherits IDOExtensionClass
        Public Function ue_Mack_rpt_ProductionBacklogSp(ByVal pStartItem As String, ByVal pEndItem As String, pSummary As String, ByVal pSite As String) As DataTable
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oItem As LoadCollectionResponseData
            Dim oItemWhse As LoadCollectionResponseData
            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String
            Dim Item As String
            Dim Description As String
            Dim ProductCode As String
            Dim QtyAllocjob As Double
            Dim Whse As String
            Dim QtyAllocCo As Double
            Dim QtyOnHand As Double
            Dim NetQty As Double
            Dim NetQty2 As Double
            Dim QtyInv As Double
            Dim QtyReorder As Double
            Dim NetYTDS As Double
            Dim UnitWeight As Double
            Dim QtyRsvdCo As Double
            Dim SiteRef As String

            pStartItem = If(pStartItem, "   ")
            pEndItem = If(pEndItem, "ｰｰｰｰｰｰ")

            sResultSetColumnsList = "Item,Description,ProductCode,ReqJob,ReqCo,QtyOnHand,NetReq,YtdsReq,Inventory,SafetyStock,InvReq,Site"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

            'Add columns to Datatable corresponding each column out of the result set.
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
            Next iIndex

            sRecordFilter = String.Format("product_code like 'C%' and p_m_t_code = 'M' and DerMetricProd <> 0 and item between {0} and {1} and site_ref = {2}", SqlLiteral.Format(pStartItem), SqlLiteral.Format(pEndItem), SqlLiteral.Format(pSite))
            'oResultDataRow = oResultDataTable.NewRow()
            'oResultDataRow(1) = sRecordFilter
            'oResultDataTable.Rows.Add(oResultDataRow)

            Dim OutputFields As String = "item, description, product_code, qty_allocjob, unit_weight, site_ref"
            oItem = Me.Context.Commands.LoadCollection("ue_item_msts", OutputFields, sRecordFilter, "product_code, item", 0)
            If oItem.Items.Count > 0 Then
                For iIndex = 0 To oItem.Items.Count - 1
                    Item = oItem(iIndex, "item").ToString
                    Description = oItem(iIndex, "description").ToString
                    ProductCode = oItem(iIndex, "product_code").ToString
                    QtyAllocjob = oItem(iIndex, "qty_allocjob").GetValue(Of Double)()
                    UnitWeight = oItem(iIndex, "unit_weight").GetValue(Of Double)()
                    SiteRef = oItem(iIndex, "site_ref").ToString
                    sRecordFilter = String.Format("Item ={0} and SiteRef = {1}", SqlLiteral.Format(Item), SqlLiteral.Format(SiteRef))
                    oItemWhse = Me.Context.Commands.LoadCollection("SLItemwhseAlls", "Whse, QtyAllocCo, QtyOnHand, QtyReorder, QtyRsvdCo", sRecordFilter, "Item", 0)
                    If oItemWhse.Items.Count > 0 Then
                        For iIndex2 = 0 To oItemWhse.Items.Count - 1
                            Whse = oItemWhse(iIndex2, "Whse").ToString
                            QtyAllocCo = oItemWhse(iIndex2, "QtyAllocCo").GetValue(Of Double)()
                            QtyOnHand = oItemWhse(iIndex2, "QtyOnHand").GetValue(Of Double)()
                            QtyReorder = oItemWhse(iIndex2, "QtyReorder").GetValue(Of Double)()
                            QtyRsvdCo = oItemWhse(iIndex2, "QtyRsvdCo").GetValue(Of Double)()

                            If QtyAllocjob <> 0 Or QtyAllocCo <> 0 Or QtyOnHand <> 0 Then
                                NetQty = QtyAllocCo - QtyOnHand
                                If QtyOnHand > QtyAllocCo Then
                                    QtyInv = QtyOnHand - QtyAllocCo
                                Else
                                    QtyInv = 0
                                End If

                                NetQty2 = QtyReorder + NetQty

                                If NetQty > 0 Then
                                    NetYTDS = NetQty * UnitWeight
                                Else
                                    NetYTDS = 0
                                End If

                                If QtyAllocCo > 0 Or QtyOnHand < 0 Or NetQty > 0 Or NetQty2 > 0 Then

                                    Dim ReqCo As Double = QtyAllocCo - QtyRsvdCo
                                    Dim OnHand As Double = QtyOnHand - QtyRsvdCo

                                    oResultDataRow = oResultDataTable.NewRow()
                                    oResultDataRow(0) = Item
                                    oResultDataRow(1) = Description
                                    oResultDataRow(2) = ProductCode
                                    oResultDataRow(3) = QtyAllocjob.ToString
                                    oResultDataRow(4) = ReqCo.ToString
                                    oResultDataRow(5) = OnHand.ToString
                                    oResultDataRow(6) = NetQty.ToString
                                    oResultDataRow(7) = NetYTDS.ToString
                                    oResultDataRow(8) = QtyInv.ToString
                                    oResultDataRow(9) = QtyReorder.ToString
                                    'oResultDataRow(9) = SiteRef
                                    If NetQty2 > 0 Then
                                        oResultDataRow(10) = NetQty2.ToString
                                    End If
                                    oResultDataTable.Rows.Add(oResultDataRow)
                                End If

                            End If

                        Next
                    End If

                Next
            End If
            Return oResultDataTable
        End Function

    End Class
End Namespace