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

Namespace ue_Mack_IntercompanyCOPricing
    ' <IDOExtensionClass( "ue_Mack_IntercompanyCOPricing" )>
    Partial Public Class ue_Mack_IntercompanyCOPricing : Inherits IDOExtensionClass
        Public Function Mack_UpdateCoPrices(ByVal Action As String, ByVal pCoNum As String, ByVal pStartLine As String, ByVal pEndLine As String, ByVal pMargin As Decimal) As DataTable
            Dim sRecordFilter As String
            Dim oCoitem As LoadCollectionResponseData
            Dim oJob As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim CoNum As String
            Dim CoLine As String
            Dim CoRelease As String
            Dim Item As String
            Dim RefType As String
            Dim RefNum As String
            Dim RefLineSuf As String
            Dim RefRelease As String
            Dim bomjobM As String
            Dim mjob2 As String
            Dim msuf2 As String
            Dim mitem1 As String
            Dim mitem2 As String
            Dim bomoper As Integer
            Dim bomseq As Integer
            Dim sResultSetColumnsList As String
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim QtyOrderedConv As Decimal
            Dim Price As Decimal
            Dim ExtendedPrice As Decimal
            Dim cototal As Decimal
            Dim NewCost As Decimal
            Dim updateRequest As UpdateCollectionRequestData
            Dim UpdateItem As IDOUpdateItem

            sResultSetColumnsList = "CoNum,CoLine,Qty,Price,ExtendedPrice,Cost,TotalCost,NewCost"
            oResultDataTable = New DataTable() 'Only table creation for now. Columns and Data will be added later.
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")  'sResultSetColumnsList.Split(CType(",", Char()))                    
            'Add columns to Datatable corresponding each column out of the result set.
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                oResultDataTable.Columns.Add(ResultSetColumnsListArray(iIndex))
            Next iIndex

            sRecordFilter = String.Format("CoNum = {0} and CoLine between {1} and {2} and CoStat = 'O' and Stat = 'O'", SqlLiteral.Format(pCoNum), SqlLiteral.Format(pStartLine), SqlLiteral.Format(pEndLine))
            Dim OutputFields As String = "CoNum,CoLine,CoRelease,Item,RefType,RefNum,RefLineSuf,RefRelease,QtyOrderedConv,Price"
            oCoitem = Me.Context.Commands.LoadCollection("SLCoitems", OutputFields, sRecordFilter, "", 0)
            If oCoitem.Items.Count > 0 Then
                'oResultDataRow = oResultDataTable.NewRow()
                'oResultDataRow(0) = pCoNum
                'oResultDataRow(1) = oCoitem.Items.Count.ToString
                'oResultDataTable.Rows.Add(oResultDataRow)

                For iIndex = 0 To oCoitem.Items.Count - 1
                    CoNum = oCoitem(iIndex, "CoNum").ToString
                    CoLine = oCoitem(iIndex, "CoLine").ToString
                    CoRelease = oCoitem(iIndex, "CoRelease").ToString
                    Item = oCoitem(iIndex, "Item").ToString
                    RefType = oCoitem(iIndex, "RefType").ToString
                    RefNum = oCoitem(iIndex, "RefNum").ToString
                    RefLineSuf = oCoitem(iIndex, "RefLineSuf").ToString
                    RefRelease = oCoitem(iIndex, "RefRelease").ToString
                    QtyOrderedConv = oCoitem(iIndex, "QtyOrderedConv").GetValue(Of Decimal)
                    Price = oCoitem(iIndex, "Price").GetValue(Of Decimal)
                    ExtendedPrice = QtyOrderedConv * Price
                    cototal = 0
                    If String.Equals(RefType, "J") Then
                        sRecordFilter = String.Format("Job = {0} and Suffix = {1}", SqlLiteral.Format(RefNum), SqlLiteral.Format(RefLineSuf))
                        OutputFields = "Job,Suffix"
                        oJob = Me.Context.Commands.LoadCollection("SLJobs", OutputFields, sRecordFilter, "", 0)
                        If oJob.Items.Count > 0 Then
                            bomjobM = oJob(0, "Job").ToString
                            mjob2 = oJob(0, "Job").ToString
                            msuf2 = oJob(0, "Suffix").ToString
                            mitem1 = Item
                            mitem2 = Item
                            bomoper = 0
                            bomseq = 0
                            cototal = InterProc(mitem2, mjob2, msuf2, bomoper, bomseq)
                        End If
                    End If
                    If String.Equals(RefType, "I") Then
                        sRecordFilter = String.Format("Item={0}", SqlLiteral.Format(Item))
                        oItem = Me.Context.Commands.LoadCollection("SLItems", "UnitCost", sRecordFilter, "", 0)
                        If oItem.Items.Count > 0 Then
                            cototal = oItem(0, "UnitCost").GetValue(Of Decimal)
                        End If
                    End If
                    NewCost = cototal * (1 + (pMargin / 100))

                    oResultDataRow = oResultDataTable.NewRow()
                    oResultDataRow(0) = CoNum
                    oResultDataRow(1) = CoLine
                    oResultDataRow(2) = QtyOrderedConv.ToString
                    oResultDataRow(3) = Price.ToString
                    oResultDataRow(4) = ExtendedPrice.ToString
                    oResultDataRow(5) = cototal.ToString
                    oResultDataRow(6) = NewCost.ToString
                    oResultDataTable.Rows.Add(oResultDataRow)

                    'Update Price
                    If String.Equals(Action, "P") Then
                        updateRequest = New UpdateCollectionRequestData("SLCoitems")
                        UpdateItem = New IDOUpdateItem(UpdateAction.Update, oCoitem.Items(iIndex).ItemID)
                        UpdateItem.Properties.Add("Price", NewCost, True)
                        UpdateItem.Properties.Add("PriceConv", NewCost, True)
                        updateRequest.Items.Add(UpdateItem)
                        Me.Context.Commands.UpdateCollection(updateRequest)

                    End If
                Next

            End If

            Return oResultDataTable

        End Function


        Function InterProc(ByRef Item As String, ByVal Job As String, ByVal Suffix As String, ByVal OperNum As Integer, ByVal Seq As Integer) As Decimal
            Dim sRecordFilter As String
            Dim OutputFields As String
            Dim oJob As LoadCollectionResponseData
            Dim oJobroute As LoadCollectionResponseData
            Dim oJobmatl As LoadCollectionResponseData
            Dim oWC As LoadCollectionResponseData
            Dim oDept As LoadCollectionResponseData
            Dim oItem As LoadCollectionResponseData
            Dim iIndex As Integer
            Dim RunRateLbr As Decimal
            Dim Dept As String
            Dim FixovhdRate As Decimal
            Dim bomqty As Decimal
            Dim JshRunTicksMch As Decimal
            Dim JshRunTicksLbr As Decimal
            Dim mchqty As Decimal
            Dim lbrqty As Decimal
            Dim lbrcost As Decimal
            Dim totlbr As Decimal
            Dim ovhdcost As Decimal
            Dim totovhd As Decimal
            Dim bomitem As String
            Dim mitemcost As Decimal
            Dim bomqty2 As Decimal
            Dim matlcost As Decimal
            Dim totmatl As Decimal
            Dim tottotal As Decimal
            Dim mitem2 As String
            Dim bomoper As Integer
            Dim bomseq As Integer
            Dim subtotal As Decimal

            sRecordFilter = String.Format("Job = {0} and Suffix = {1} and Item = {2}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix), SqlLiteral.Format(Item))
            oJob = Me.Context.Commands.LoadCollection("SLJobs", "Job, Suffix", sRecordFilter, "", 0)
            If oJob.Items.Count > 0 Then
                sRecordFilter = String.Format("Job = {0} and Suffix = {1}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix))
                OutputFields = "JshRunTicksMch,JshRunTicksLbr,Wc"
                oJobroute = Me.Context.Commands.LoadCollection("SLJobRoutes", OutputFields, sRecordFilter, "", 0)
                If oJobroute.Items.Count > 0 Then
                    For iIndex = 0 To oJobroute.Items.Count - 1
                        JshRunTicksMch = oJobroute(iIndex, "JshRunTicksMch").GetValue(Of Decimal)
                        JshRunTicksLbr = oJobroute(iIndex, "JshRunTicksLbr").GetValue(Of Decimal)

                        sRecordFilter = String.Format("Wc = {0}", SqlLiteral.Format(oJobroute(iIndex, "Wc").ToString))
                        oWC = Me.Context.Commands.LoadCollection("SLWcs", "RunRateLbr,Dept", sRecordFilter, "", 0)
                        RunRateLbr = oWC(0, "RunRateLbr").GetValue(Of Decimal)
                        Dept = oWC(0, "Dept").GetValue(Of String)

                        sRecordFilter = String.Format("Dept = {0}", SqlLiteral.Format(Dept))
                        oDept = Me.Context.Commands.LoadCollection("SLDepts", "FixovhdRate", sRecordFilter, "", 0)
                        FixovhdRate = oDept(0, "FixovhdRate").GetValue(Of Decimal)

                        sRecordFilter = String.Format("Job={0} and Suffix={1} and OperNum={2} and Sequence={3} and Item = {4}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix), SqlLiteral.Format(OperNum), SqlLiteral.Format(Seq), SqlLiteral.Format(Item))
                        oJobmatl = Me.Context.Commands.LoadCollection("SLJobmatls", "MatlQty", sRecordFilter, "", 0)
                        If oJobmatl.Items.Count > 0 Then
                            bomqty = oJobmatl(0, "MatlQty").GetValue(Of Decimal)
                        Else
                            bomqty = 1
                        End If

                        mchqty = JshRunTicksMch / 100 * bomqty
                        lbrqty = JshRunTicksLbr / 100 * bomqty
                        lbrcost = lbrqty * RunRateLbr
                        totlbr = totlbr + lbrcost
                        ovhdcost = lbrqty * FixovhdRate
                        totovhd = totovhd + ovhdcost
                    Next
                End If

                sRecordFilter = String.Format("Job={0} and Suffix={1}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix))
                oJobmatl = Me.Context.Commands.LoadCollection("SLJobmatls", "MatlQty,MatlQtyConv,Cost,Item,UM", sRecordFilter, "", 0)
                If oJobmatl.Items.Count > 0 Then
                    For iIndex = 0 To oJobmatl.Items.Count - 1
                        bomitem = oJobmatl(iIndex, "Item").GetValue(Of String)
                        bomqty2 = oJobmatl(iIndex, "MatlQtyConv").GetValue(Of Decimal)
                        mitemcost = oJobmatl(iIndex, "Cost").GetValue(Of Decimal)

                        sRecordFilter = String.Format("Item={0}", SqlLiteral.Format(oJobmatl(iIndex, "Item").GetValue(Of String)))
                        oItem = Me.Context.Commands.LoadCollection("SLItems", "Description,UM", sRecordFilter, "", 0)
                        If oItem.Items.Count > 0 Then
                            matlcost = mitemcost * bomqty2 * bomqty
                        Else
                            matlcost = mitemcost * bomqty2
                        End If
                        totmatl = totmatl + matlcost
                    Next
                End If
                tottotal = totmatl + totovhd + totlbr


                sRecordFilter = String.Format("Job={0} and Suffix={1}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix))
                oJobmatl = Me.Context.Commands.LoadCollection("SLJobmatls", "Job,Suffix,OperNum,Sequence,Item", sRecordFilter, "", 0)
                If oJobmatl.Items.Count > 0 Then
                    For iIndex = 0 To oJobmatl.Items.Count - 1
                        sRecordFilter = String.Format("Item={0}", SqlLiteral.Format(oJobmatl(iIndex, "Item").GetValue(Of String)))
                        oItem = Me.Context.Commands.LoadCollection("SLItems", "Description,UM", sRecordFilter, "", 0)
                        If oItem.Items.Count > 0 Then
                            mitem2 = oJobmatl(iIndex, "Item").GetValue(Of String)
                            bomoper = oJobmatl(iIndex, "OperNum").GetValue(Of Integer)
                            bomseq = oJobmatl(iIndex, "Sequence").GetValue(Of Integer)
                            subtotal = InterProc(mitem2, Job, Suffix, bomoper, bomseq)
                            tottotal = tottotal + subtotal
                        End If
                    Next
                End If

            End If ' If oJob.Items.Count > 0 Then
            Return tottotal
        End Function
    End Class
End Namespace