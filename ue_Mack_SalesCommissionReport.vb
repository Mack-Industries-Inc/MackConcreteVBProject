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

Namespace ue_Mack_SalesCommissionReport
    ' <IDOExtensionClass( "ue_Mack_SalesCommissionReport" )>
    Partial Public Class ue_Mack_SalesCommissionReport : Inherits IDOExtensionClass
        Public Function ue_MACK_Rpt_SalesCommissionbySalespersonSp(ByVal StartSalesman As String, ByVal EndSalesman As String,
          ByVal StartDueDate As String, ByVal EndDueDate As String,
          ByVal StartDateOffset As String, ByVal EndDateOffset As String,
          ByVal pSite As String) As DataTable

            Dim SalesCommissionReport As New DataTable
            Dim sResultSetColumnsList As String
            Dim sFilter As String
            Dim oSlsman As LoadCollectionResponseData
            Dim invokeResponse As InvokeResponseData
            Dim oResultDataRow As DataRow
            Dim Slsman As String
            Dim Margin As Decimal
            Dim TotalMargin As Decimal
            Dim WeightFactor As Decimal
            Dim WeightMargPercent As Decimal
            Dim PretaxPrice As Decimal
            Dim MargPercent As Decimal
            Dim TotalCollections As Decimal
            Dim Collections As Decimal
            Dim WeightDSOFactor As Decimal
            Dim WeightedDSO As Decimal
            Dim DSO As Decimal

            If String.IsNullOrEmpty(StartSalesman) Then
                StartSalesman = "      "
            End If
            If String.IsNullOrEmpty(EndSalesman) Then
                EndSalesman = "ｰｰｰｰｰｰｰｰｰｰ"
            End If
            If String.IsNullOrEmpty(StartDueDate) Then
                StartDueDate = "1753-01-01 00:00:00.000"
            End If
            If String.IsNullOrEmpty(EndDueDate) Then
                EndDueDate = "9999-12-31 23:59:59.997"
            End If

            invokeResponse = Me.Context.Commands.Invoke("SLJobTrans", "ApplyDateOffsetSp", StartDueDate, StartDateOffset, "0")
            StartDueDate = invokeResponse.Parameters(0).Value.ToString
            invokeResponse = Me.Context.Commands.Invoke("SLJobTrans", "ApplyDateOffsetSp", EndDueDate, EndDateOffset, "1")
            EndDueDate = invokeResponse.Parameters(0).Value.ToString

            sResultSetColumnsList = "InvNum,Slsman,SlsmanName,CommCalc,CommDue,DueDate,CommBase,CoNum,CustNum,CommPercent,Tax,TotalPrice,PretaxPrice,Cost,Margin,MargPercent,InvDate,DSO,Collections,SiteRef,WeightingFactor,WeightedMargPercent,WeightDSOFactor,WeightedDSO,DOCost"

            Dim oDataCol As DataColumn = Nothing
            Dim ResultSetColumnsListArray() As String = sResultSetColumnsList.Split(",")
            For iIndex = 0 To ResultSetColumnsListArray.Length - 1
                oDataCol = New DataColumn(ResultSetColumnsListArray(iIndex))
                If String.Equals(ResultSetColumnsListArray(iIndex), "Margin") Or String.Equals(ResultSetColumnsListArray(iIndex), "Collections") Then
                    oDataCol.DataType = System.Type.GetType("System.Decimal")
                End If
                SalesCommissionReport.Columns.Add(oDataCol)
            Next iIndex

            Dim OutputField As String = "ue_InvNum,Slsman,ue_SlsmanName,ue_CommCalc,ue_CommDue,ue_DueDate,ue_CommBase,ue_CoNum,ue_InvCustNum,ue_DerCalcCommPercent,
            ue_DerInvTax,ue_InvPrice,ue_DerPretaxPrice,ue_InvCost,ue_DerMargin,ue_DerMarginPercent,ue_InvDate,ue_DerDSO,ue_DerCollections,SiteRef"

            sFilter = String.Format("Slsman between {0} and {1} and ue_DueDate between {2} and {3}", SqlLiteral.Format(StartSalesman), SqlLiteral.Format(EndSalesman), SqlLiteral.Format(StartDueDate), SqlLiteral.Format(EndDueDate))
            oSlsman = Context.Commands.LoadCollection("ue_Mack_SlsmanCommDueAlls", OutputField, sFilter, "Slsman", 0)

            For iIndex = 0 To oSlsman.Items.Count - 1
                oResultDataRow = SalesCommissionReport.NewRow()
                For iColumnIndex As Integer = 0 To ResultSetColumnsListArray.Length - 6 'i.e. Ignoring the 1 custom fields. This will be populated later.
                    If String.Equals(ResultSetColumnsListArray(iColumnIndex), "Margin") Then
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oSlsman.Items(iIndex).PropertyValues(iColumnIndex).GetValue(Of Decimal)(0.0)
                    Else
                        oResultDataRow(ResultSetColumnsListArray(iColumnIndex)) = oSlsman.Items(iIndex).PropertyValues(iColumnIndex).ToString
                    End If
                Next iColumnIndex
                SalesCommissionReport.Rows.Add(oResultDataRow)
            Next

            For Each oRow As DataRow In SalesCommissionReport.Rows
                Slsman = oRow.Item("Slsman").ToString()
                If String.Equals(oRow.Item("PretaxPrice"), IDONull.Value) Then
                    PretaxPrice = 0
                Else
                    PretaxPrice = Convert.ToDecimal(oRow.Item("PretaxPrice").ToString)
                End If

                If String.Equals(oRow.Item("Margin"), IDONull.Value) Then
                    Margin = 0
                Else
                    Margin = Convert.ToDecimal(oRow.Item("Margin").ToString)
                End If

                If String.Equals(oRow.Item("MargPercent"), IDONull.Value) Then
                    MargPercent = 0
                Else
                    MargPercent = Convert.ToDecimal(oRow.Item("MargPercent").ToString)
                End If

                If String.Equals(oRow.Item("DSO"), IDONull.Value) Then
                    DSO = 0
                Else
                    DSO = Convert.ToDecimal(oRow.Item("DSO").ToString)
                End If

                If String.Equals(oRow.Item("Collections"), IDONull.Value) Then
                    Collections = 0
                Else
                    Collections = Convert.ToDecimal(oRow.Item("Collections").ToString)
                End If

                sFilter = "Slsman='" & Slsman & "'"
                TotalMargin = Convert.ToDecimal(SalesCommissionReport.Compute("Sum(Margin)", sFilter))
                TotalCollections = Convert.ToDecimal(SalesCommissionReport.Compute("Sum(Collections)", sFilter))

                If TotalMargin <> 0 Then
                    WeightFactor = Margin / TotalMargin
                Else
                    WeightFactor = 0
                End If

                If MargPercent < 0 Then
                    WeightMargPercent = WeightFactor * MargPercent * -1
                Else
                    WeightMargPercent = WeightFactor * MargPercent
                End If


                oRow.Item("WeightingFactor") = WeightFactor.ToString
                oRow.Item("WeightedMargPercent") = WeightMargPercent.ToString

                If TotalCollections <> 0 Then
                    WeightDSOFactor = Collections / TotalCollections
                Else
                    WeightDSOFactor = 0
                End If

                WeightedDSO = WeightDSOFactor * DSO

                oRow.Item("WeightDSOFactor") = WeightDSOFactor.ToString
                oRow.Item("WeightedDSO") = WeightedDSO.ToString

            Next

            'oResultDataRow = SalesCommissionReport.NewRow()
            'oResultDataRow("Margin") = TotalMargin.ToString
            'oResultDataRow("SlsmanName") = "Total Margin"
            'SalesCommissionReport.Rows.Add(oResultDataRow)

            Return SalesCommissionReport
        End Function


    End Class
End Namespace