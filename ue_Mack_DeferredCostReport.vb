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

Namespace ue_Mack_DeferredCostReport
    ' <IDOExtensionClass( "ue_Mack_DeferredCostReport" )>
    Partial Public Class ue_Mack_DeferredCostReport : Inherits IDOExtensionClass
        Public Function ue_Mack_rpt_DeferredCostReportSp(ByVal pShipDate As String) As DataTable
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim oCoitem As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim oPoitem As LoadCollectionResponseData
            Dim sRecordFilter As String
            Dim sResultSetColumnsList As String
            Dim CoNum As String
            Dim CoLine As String
            Dim CoRelease As String
            Dim sStructure As String
            Dim Item As String
            Dim Description As String
            Dim EndUserType As String
            Dim ShipDate As String
            Dim ItemUnitCost As Double
            Dim PoNum As String
            Dim PoLine As String
            Dim PoRelease As String
            Dim RefType As String
            Dim QtyOrdered As Double
            Dim TotalCost As Double


            sResultSetColumnsList = "EUT,CoNum,CoLine,Structure,Item,Description,Qty,UnitCost,ExtCost,Date"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

            'Add columns to Datatable corresponding each column out of the result set.
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
            Next iIndex

            sRecordFilter = String.Format("(ShipDate > {0} or ShipDate is NULL) and CoType='R' and CoStat = 'O' and ISNULL(ue_ItemDeferred,0) = 0 and ISNULL(coiUf_Structure, '') <> '' and Stat <> 'C'", SqlLiteral.Format(pShipDate))
            Dim OutputFields As String = "CoNum,CoLine,CoRelease,Item,Description,coiUf_Structure,ue_CoEndUserType,ShipDate,RefType,RefNum,RefLineSuf,RefRelease,QtyOrdered,Price"
            oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", OutputFields, sRecordFilter, "CoNum, CoLine, CoRelease", 0)
            If oCoitem.Items.Count > 0 Then
                For iIndex = 0 To oCoitem.Items.Count - 1
                    CoNum = oCoitem(iIndex, "CoNum").ToString
                    CoLine = oCoitem(iIndex, "CoLine").ToString
                    CoRelease = oCoitem(iIndex, "CoRelease").ToString
                    sStructure = oCoitem(iIndex, "coiUf_Structure").ToString
                    Item = oCoitem(iIndex, "Item").ToString
                    Description = oCoitem(iIndex, "Description").ToString
                    EndUserType = oCoitem(iIndex, "ue_CoEndUserType").ToString
                    ShipDate = oCoitem(iIndex, "ShipDate").ToString

                    RefType = oCoitem(iIndex, "RefType").ToString
                    PoNum = oCoitem(iIndex, "RefNum").ToString
                    PoLine = oCoitem(iIndex, "RefLineSuf").ToString
                    PoRelease = oCoitem(iIndex, "RefRelease").ToString

                    QtyOrdered = oCoitem(iIndex, "QtyOrdered").GetValue(Of Double)()
                    Description = sStructure & "^" & Description
                    sRecordFilter = String.Format("Item = {0}", SqlLiteral.Format(Item))
                    oItem = Me.Context.Commands.LoadCollection("SLItems", "UnitCost", sRecordFilter, "", 0)
                    If oItem.Items.Count > 0 Then
                        ItemUnitCost = oItem(0, "UnitCost").GetValue(Of Double)()
                    End If

                    If (String.Equals(RefType, "P") And Not String.Equals(PoNum, "")) Then
                        sRecordFilter = String.Format("PoNum = {0} and PoLine = {1} and PoRelease = {2}", SqlLiteral.Format(PoNum), SqlLiteral.Format(PoLine), SqlLiteral.Format(PoRelease))
                        oPoitem = Me.Context.Commands.LoadCollection("SLPoItems", "ItemCostConv", sRecordFilter, "", 0)
                        If oPoitem.Items.Count > 0 Then
                            ItemUnitCost = oPoitem(0, "ItemCostConv").GetValue(Of Double)()
                        End If
                    End If

                    TotalCost = ItemUnitCost * QtyOrdered

                    oResultDataRow = oResultDataTable.NewRow()
                    oResultDataRow(0) = EndUserType
                    oResultDataRow(1) = CoNum
                    oResultDataRow(2) = CoLine
                    oResultDataRow(3) = sStructure
                    oResultDataRow(4) = Item
                    oResultDataRow(5) = Description
                    oResultDataRow(6) = QtyOrdered
                    oResultDataRow(7) = ItemUnitCost.ToString
                    oResultDataRow(8) = TotalCost.ToString
                    oResultDataRow(9) = ShipDate

                    oResultDataTable.Rows.Add(oResultDataRow)


                Next
            End If

            ue_Mack_rpt_DeferredCostReportSp = oResultDataTable
        End Function


    End Class
End Namespace