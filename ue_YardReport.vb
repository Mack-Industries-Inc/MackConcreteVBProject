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

Namespace ue_YardReport
    ' <IDOExtensionClass( "ue_YardReport" )>
    Partial Public Class ue_YardReport : Inherits IDOExtensionClass
        Public Function Rpt_YardReportSp(ByVal StartingLoc As String, ByVal EndingLoc As String, ByVal pSite As String)
            Dim sFilter As String
            Dim iIndex As Integer
            Dim oResultDataTable As DataTable
            Dim oResultDataRow As DataRow
            Dim dResultTable As DataTable = New DataTable
            Dim sResultSetColumnsList As String
            Dim oItemLoc As LoadCollectionResponseData
            Dim oLotLoc As LoadCollectionResponseData
            Dim oMatl As LoadCollectionResponseData
            Dim oJob As LoadCollectionResponseData
            Dim oCoitem As LoadCollectionResponseData
            Dim oMatl2 As LoadCollectionResponseData

            Dim Item As String
            Dim ItemDesc As String
            Dim Loc As String
            Dim Whse As String
            Dim ItmLotTracked As String
            Dim tItem As String
            Dim iLot As Integer
            Dim Lot As String
            Dim iMatl As Integer
            Dim iMatl2 As Integer
            Dim iJob As Integer
            Dim iCoitem As Integer

            Dim RefType As String
            Dim RefNum As String
            Dim RefLineSuf As String
            Dim RefRelease As String
            Dim CoNum As String
            Dim CoLine As String
            Dim CoRelease As String
            Dim DueDate As String
            Dim CustPo As String
            Dim tDelv As Date
            Dim LotQtyOnHand As String
            Dim Job As String
            Dim Suffix As String
            Dim OrdType As String
            Dim sStructure As String
            Dim QtyOnHand As String
            Dim RcptDate As String
            Dim RcptQty As String
            Dim iCount As Integer

            Dim xYard As New DataTable

            sResultSetColumnsList = "Loc,Whse,Item,Lot,QOH,RcptQty,RcptDate,Job,Suffix,Structure,CoNum,CoLine,CoRelease,DueDate,LotQtyOnHand"
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                xYard.Columns.Add(ResultSetColumnsListArray(iIndex))
            Next iIndex

            sFilter = "QtyOnHand>0"
            If StartingLoc <> "" And EndingLoc <> "" Then
                sFilter = String.Format("Loc between {0} and {1} and QtyOnHand>0", SqlLiteral.Format(StartingLoc), SqlLiteral.Format(EndingLoc))
            End If
            If StartingLoc <> "" And EndingLoc = "" Then
                sFilter = String.Format("Loc >= {0} and QtyOnHand>0", SqlLiteral.Format(StartingLoc))
            End If
            If StartingLoc = "" And EndingLoc <> "" Then
                sFilter = String.Format("Loc <= {0} and QtyOnHand>0", SqlLiteral.Format(EndingLoc))
            End If

            oItemLoc = Context.Commands.LoadCollection("SLItemLocs", "Item, Loc, Whse, ItmLotTracked, ItmDescription, QtyOnHand", sFilter, "Loc, Item", 0)
            For iIndex = 0 To oItemLoc.Items.Count - 1
                Loc = oItemLoc(iIndex, "Loc").ToString
                Whse = oItemLoc(iIndex, "Whse").ToString
                Item = oItemLoc(iIndex, "Item").ToString
                ItmLotTracked = oItemLoc(iIndex, "ItmLotTracked").ToString
                ItemDesc = oItemLoc(iIndex, "ItmDescription").ToString
                QtyOnHand = oItemLoc(iIndex, "QtyOnHand").ToString
                Lot = ""
                CoNum = ""
                CoLine = ""
                CoRelease = ""
                DueDate = ""
                CustPo = ""
                sStructure = ""
                OrdType = ""
                iCount = 0
                sFilter = String.Format("QtyReleased > QtyComplete and Type = 'J' and Stat = 'R' and jobUf_HoleLoc={0} and Whse={1} and Item = {2}", SqlLiteral.Format(Loc), SqlLiteral.Format(Whse), SqlLiteral.Format(Item))
                oJob = Context.Commands.LoadCollection("SLJobs", "Job, Suffix, Item, ue_Uf_Structure, OrdType, OrdNum, OrdLine, OrdRelease, jobUf_HoleLoc, QtyReleased, JschCompdate", sFilter, "", 0)
                For iJob = 0 To oJob.Items.Count - 1
                    'Loc = oJob(iJob, "jobUf_HoleLoc").ToString
                    sStructure = oJob(iJob, "ue_Uf_Structure").ToString
                    Job = oJob(iJob, "Job").ToString
                    Suffix = oJob(iJob, "Suffix").ToString
                    OrdType = oJob(iJob, "OrdType").ToString
                    Item = oJob(iJob, "Item").ToString
                    RcptDate = oJob(iJob, "JschCompdate").ToString
                    RcptQty = oJob(iJob, "QtyReleased").ToString
                    If String.Equals(OrdType, "O") Then
                        CoNum = oJob(iJob, "OrdNum").ToString
                        CoLine = oJob(iJob, "OrdLine").ToString
                        CoRelease = oJob(iJob, "OrdRelease").ToString
                        sFilter = String.Format("CoNum = {0} And CoLine = {1} And CoRelease = {2}", SqlLiteral.Format(CoNum), SqlLiteral.Format(CoLine), SqlLiteral.Format(CoRelease))
                        oCoitem = Context.Commands.LoadCollection("SLCoitems", "DueDate, CoCustPo", sFilter, "", 0)
                        If oCoitem.Items.Count > 0 Then
                            DueDate = oCoitem(0, "DueDate").ToString
                            CustPo = oCoitem(0, "CoCustPo").ToString
                        End If
                    Else
                        CoNum = ""
                        CoLine = ""
                        CoRelease = ""
                        DueDate = ""
                        CustPo = ""
                    End If
                    xYard.Rows.Add(Loc, Whse, Item, "", QtyOnHand, RcptQty, RcptDate, Job, Suffix, sStructure, CoNum, CoLine, CoRelease, DueDate)
                    iCount = iCount + 1
                Next

                'lot
                sFilter = String.Format("Loc={0} and Whse={1} and Item={2} and QtyOnHand > 0", SqlLiteral.Format(Loc), SqlLiteral.Format(Whse), SqlLiteral.Format(Item))
                oLotLoc = Context.Commands.LoadCollection("SLLotLocs", "Lot, QtyOnHand", sFilter, "", 0)
                For iLot = 0 To oLotLoc.Items.Count - 1
                    Lot = oLotLoc(iLot, "Lot").ToString
                    LotQtyOnHand = oLotLoc(iLot, "QtyOnHand").ToString

                    sFilter = String.Format("Loc={0} and Whse={1} and Item={2} and Lot = {3} and RefType='J' and Qty>0 and TransType='F'", SqlLiteral.Format(Loc), SqlLiteral.Format(Whse), SqlLiteral.Format(Item), SqlLiteral.Format(Lot))
                    oMatl = Context.Commands.LoadCollection("SLMatltrans", "RefType, RefNum, RefLineSuf, RefRelease, Lot, Qty, TransDate", sFilter, "TransDate desc", 0)
                    For iMatl = 0 To oMatl.Items.Count - 1
                        Job = oMatl(iMatl, "RefNum").ToString
                        Suffix = oMatl(iMatl, "RefLineSuf").ToString
                        Lot = oMatl(iMatl, "Lot").ToString
                        RcptQty = oMatl(iMatl, "Qty").ToString
                        RcptDate = oMatl(iMatl, "TransDate").ToString
                        CoNum = ""
                        CoLine = ""
                        CoRelease = ""
                        DueDate = ""
                        CustPo = ""
                        sStructure = ""
                        OrdType = ""
                        sFilter = String.Format("Job={0} and Suffix={1}", SqlLiteral.Format(Job), SqlLiteral.Format(Suffix))
                        oJob = Context.Commands.LoadCollection("SLJobs", "Job, Suffix, Item, ue_Uf_Structure, OrdType, OrdNum, OrdLine, OrdRelease, jobUf_HoleLoc, QtyReleased, JschCompdate", sFilter, "", 0)
                        If oJob.Items.Count > 0 Then
                            sStructure = oJob(0, "ue_Uf_Structure").ToString
                            OrdType = oJob(0, "OrdType").ToString
                            If String.Equals(OrdType, "O") Then
                                CoNum = oJob(0, "OrdNum").ToString
                                CoLine = oJob(0, "OrdLine").ToString
                                CoRelease = oJob(0, "OrdRelease").ToString
                                sFilter = String.Format("CoNum = {0} And CoLine = {1} And CoRelease = {2}", SqlLiteral.Format(CoNum), SqlLiteral.Format(CoLine), SqlLiteral.Format(CoRelease))
                                oCoitem = Context.Commands.LoadCollection("SLCoitems", "DueDate, CoCustPo", sFilter, "", 0)
                                If oCoitem.Items.Count > 0 Then
                                    DueDate = oCoitem(0, "DueDate").ToString
                                    CustPo = oCoitem(0, "CoCustPo").ToString
                                End If
                            End If
                        End If
                        xYard.Rows.Add(Loc, Whse, Item, Lot, QtyOnHand, RcptQty, RcptDate, Job, Suffix, sStructure, CoNum, CoLine, CoRelease, DueDate, LotQtyOnHand)
                        iCount = iCount + 1
                    Next
                Next


                If iCount = 0 Then
                    xYard.Rows.Add(Loc, Whse, Item, Lot, QtyOnHand)
                End If
            Next

            Return xYard

        End Function

    End Class
End Namespace

