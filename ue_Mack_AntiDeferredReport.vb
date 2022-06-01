Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Transactions
Imports System.Data
Imports System.Runtime.InteropServices
Imports Mongoose.Core.Common
Imports Mongoose.Core.DataAccess
Imports Mongoose.IDO
Imports Mongoose.IDO.Protocol
Imports Mongoose.IDO.DataAccess

Namespace ue_Mack_AntiDeferredReport
    ' <IDOExtensionClass( "ue_Mack_AntiDeferredReport" )>
    Partial Public Class ue_Mack_AntiDeferredReport : Inherits IDOExtensionClass

        Public Function ue_Mack_AntiDeferred(ByVal pShipCutoffDate As String, pSite As String) As DataTable
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
            Dim TotalPrice As Double
            Dim Price As Double

            sResultSetColumnsList = "EUT,CoNum,CoLine,Structure,Item,Description,Qty,UnitCost,ExtCost,Date"

            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    

            'Add columns to Datatable corresponding each column out of the result set.
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
            Next iIndex

            sRecordFilter = String.Format("(ShipDate > {0} or ShipDate is NULL) and CoType='R' and CoStat = 'O' and ISNULL(ue_ItemDeferred,0) = 1 and ISNULL(coiUf_Structure, '') <> '' and Stat <> 'C'", SqlLiteral.Format(pShipCutoffDate))
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
                    Price = oCoitem(iIndex, "Price").GetValue(Of Double)()

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
                    TotalPrice = Price * QtyOrdered

                    oResultDataRow = oResultDataTable.NewRow()
                    oResultDataRow(0) = EndUserType
                    oResultDataRow(1) = CoNum
                    oResultDataRow(2) = CoLine
                    oResultDataRow(3) = sStructure
                    oResultDataRow(4) = Item
                    oResultDataRow(5) = Description
                    oResultDataRow(6) = QtyOrdered
                    oResultDataRow(7) = TotalCost.ToString
                    oResultDataRow(8) = TotalPrice.ToString
                    'oResultDataRow(9) = ShipDate

                    oResultDataTable.Rows.Add(oResultDataRow)


                Next
            End If

            Return oResultDataTable
        End Function

        Public Function ue_Mack_AntiDeferred_OLD(ByVal pShipCutoffDate As String, pSite As String) As IDataReader

            Dim results As DataTable = New DataTable("Mack_AntiDeferred")
            Dim appDB As ApplicationDB = Me.CreateApplicationDB()
            Dim cmd As IDbCommand = appDB.Connection.CreateCommand()
            Dim parm_Site, parm_ShipCutoffDate As IDbDataParameter

            cmd.CommandType = CommandType.Text
            cmd.CommandText = "

   EXEC SetSiteSp @pSite, NULL 

   Declare
   @CoNum   nvarchar(100)
   ,@CoLine  int
   ,@CoRelease  int
   ,@Item   nvarchar(100)
   ,@RefTyp  nvarchar(10)
   ,@RefNum  nvarchar(100)
   ,@RefLine  int
   ,@RefRelease int
   ,@UnitCost  decimal(19,8)
   ,@ItemDesc  nvarchar(100)
   ,@Structure  nvarchar(100)
   ,@QtyShip  decimal(19,8)
   ,@Price   decimal(19,8)
   ,@tUnitCost  decimal(19,8)

   Create Table #Result
   (CoNum   nvarchar(100)
   ,CoLine   int
   ,Item   nvarchar(100)
   ,ItemDesc  nvarchar(100)
   ,Structure  nvarchar(100)
   ,EndUserType nvarchar(100)
   ,QtyShip  decimal(19,8)
   ,UnitCost  decimal(19,8)
   ,Price   decimal(19,8)
   )

   insert into #Result (CoNum, CoLine, Item, ItemDesc, Structure, QtyShip, UnitCost, Price, EndUserType)
   select coitem.co_num, coitem.co_line, coitem.item, item.description, coitem.Uf_Structure
   ,coitem.qty_shipped
   ,case when ISNULL(poitem.po_num, '') = '' then item.unit_cost else poitem.item_cost_conv end
   ,coitem.price_conv
   ,co.end_user_type
   from coitem
   inner join co on coitem.co_num = co.co_num
   inner join item on coitem.item = item.item
   left join poitem on po_num = coitem.ref_num and po_line = coitem.ref_line_suf and po_release = coitem.ref_release
   where co.type = 'R' and co.stat = 'O'
   and coitem.qty_shipped > 0
   and coitem.ship_date > @ShipDate
   --and (charindex('base' , item.description) = 0 and charindex('box' , item.description) = 0)

   select r.EndUserType, r.CoNum, r.CoLine, r.Structure, r.Item, r.ItemDesc
   ,r.QtyShip
   ,r.UnitCost * r.QtyShip as ExtCost
   ,r.Price * r.QtyShip as ExtPrice
   from #Result r
   where ISNULL(r.Structure, '') <> '' and
   exists(select 1 from coitem 
   inner join item on coitem.item = item.item
   where coitem.co_num = r.CoNum and coitem.co_line <> r.CoLine
   and (coitem.stat = 'O' or coitem.ship_date > @ShipDate)
   and  coitem.qty_ordered <> 0
   and (charindex('base' , item.description) > 0 or charindex('box' , item.description) > 0)
   and (coitem.qty_shipped < coitem.qty_ordered_conv or coitem.ship_date > @ShipDate)
   and coitem.qty_ordered <> 0 and coitem.price <> 0
   )

   "

            parm_ShipCutoffDate = cmd.CreateParameter()
            parm_ShipCutoffDate.ParameterName = "@ShipDate"
            parm_ShipCutoffDate.Value = If(CObj(pShipCutoffDate), DBNull.Value)

            parm_Site = cmd.CreateParameter()
            parm_Site.ParameterName = "@pSite"
            parm_Site.Value = pSite

            cmd.Parameters.Add(parm_Site)
            cmd.Parameters.Add(parm_ShipCutoffDate)

            cmd.Connection = appDB.Connection
            Dim reader As IDataReader = cmd.ExecuteReader()
            results.Load(reader)

            Return results.CreateDataReader()


        End Function

    End Class
End Namespace