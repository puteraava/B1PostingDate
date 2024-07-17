Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Collections.Generic
Public Class DLVR_SCH__Sub
    Public Shared Function SetRemark(ByRef form As SAPbouiCOM.Form) As String
        Try
            Dim mtx As SAPbouiCOM.Matrix = form.Items.Item("mtx_ds").Specific
            Dim i As Integer
            Dim remarks As String = "SQ#"
            For i = 1 To mtx.RowCount
                If mtx.Columns.Item("check").Cells.Item(i).Specific.Checked Then
                    remarks += mtx.Columns.Item("PtNo").Cells.Item(i).Specific.Value.ToString
                End If
            Next i
            Return remarks
        Catch ex As Exception
            Throw New Exception("SetRemark|" + ex.Message.ToString)
        End Try
    End Function
    Public Shared Sub RefreshRemark(ByRef form As SAPbouiCOM.Form)
        Try
            With form.DataSources.DBDataSources.Item("@ODLS")
                .SetValue("U_IDU_Remarks", 0, DLVR_SCH__Sub.SetRemark(form))
            End With

        Catch ex As Exception
            Throw New Exception("RefreshRemark|" + ex.Message.ToString)
        End Try
    End Sub

    Public Shared Function GetTotal(ByRef form As SAPbouiCOM.Form) As Double
        Try
            Dim mtx As SAPbouiCOM.Matrix = form.Items.Item("mtx_ds").Specific
            Dim i As Integer
            Dim Total As Double = 0
            For i = 1 To mtx.RowCount
                If mtx.Columns.Item("check").Cells.Item(i).Specific.Checked Then
                    Total = Total + Val(mtx.Columns.Item("TtlAmnt").Cells.Item(i).Specific.Value)
                End If
            Next i
            Return Total
        Catch ex As Exception
            Throw New Exception("GetTotal|" + ex.Message.ToString)
        End Try
    End Function

    Public Shared Function GetTotalQty(ByRef form As SAPbouiCOM.Form) As Double
        Try
            Dim mtx As SAPbouiCOM.Matrix = form.Items.Item("mtx_ds").Specific
            Dim i As Integer
            Dim Total As Double = 0
            For i = 1 To mtx.RowCount
                Total = Total + Val(mtx.Columns.Item("Qty").Cells.Item(i).Specific.Value)
            Next i
            Return Total
        Catch ex As Exception
            Throw New Exception("GetTotalQty|" + ex.Message.ToString)
        End Try
    End Function

    Public Shared Function GetTotalChecked(ByRef form As SAPbouiCOM.Form) As Double
        Try
            Dim mtx As SAPbouiCOM.Matrix = form.Items.Item("mtx_ds").Specific
            Dim i As Integer
            Dim Total As Double = 0
            For i = 1 To mtx.RowCount
                If mtx.Columns.Item("check").Cells.Item(i).Specific.Checked Then
                    Total = Total + 1
                End If
            Next i
            Return Total
        Catch ex As Exception
            Throw New Exception("GetTotalChecked|" + ex.Message.ToString)
        End Try
    End Function

    Public Shared Sub RefreshTotalAmount(ByRef form As SAPbouiCOM.Form)
        Try
            With form.DataSources.DBDataSources.Item("@ODLS")
                .SetValue("U_IDU_DocTotal", 0, DLVR_SCH__Sub.GetTotal(form))
            End With

        Catch ex As Exception
            Throw New Exception("RefreshTotalAmount|" + ex.Message.ToString)
        End Try
    End Sub

    Public Shared Function GetGrandTotal(ByRef form As SAPbouiCOM.Form) As Double
        Try
            Dim Total As Double = CDbl(form.Items.Item("DocTotal").Specific.Value)
            Dim LCL As Double = CDbl(form.Items.Item("LCLAmnt").Specific.Value)

            Dim Grand As Double = Total + LCL

            Return Grand
        Catch ex As Exception
            Throw New Exception("GetGrand|" + ex.Message.ToString)
        End Try
    End Function

    Public Shared Sub RefreshGrandTotal(ByRef form As SAPbouiCOM.Form)
        Try
            With form.DataSources.DBDataSources.Item("@ODLS")
                .SetValue("U_IDU_GrandTotal", 0, DLVR_SCH__Sub.GetGrandTotal(form))
            End With

        Catch ex As Exception
            Throw New Exception("RefreshTotalAmount|" + ex.Message.ToString)
        End Try
    End Sub

    Public Shared Function Transferto01(ByRef dsentry As String) As Boolean
        Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
        Dim lRetCode As Integer
        Dim errorCode As Integer
        Dim errorMessage As String = ""
        oDoc = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

        Dim dsdata As Recordset = Control._recordSet("call _IDU_DLS_GETDLSBYID('" & dsentry & "')")
        Try
            dsdata.MoveFirst()

            oDoc.DocDate = dsdata.Fields.Item("U_IDU_DocDate").Value.ToString
            oDoc.TaxDate = dsdata.Fields.Item("U_IDU_DocDate").Value.ToString
            oDoc.FromWarehouse = dsdata.Fields.Item("U_IDU_Whse").Value.ToString
            oDoc.ToWarehouse = "01"
            oDoc.UserFields.Fields.Item("U_IDU_DSNo").Value = dsentry

            For index = 1 To dsdata.RecordCount
                Dim absEntry As String = Control._retRstField("SELECT T0.""AbsEntry"" FROM ""OBIN"" T0 WHERE T0.""BinCode"" = '" & dsdata.Fields.Item("U_IDU_BIN_Loc").Value.ToString & "'")

                If absEntry = "" Then
                    Throw New Exception("Error Code : -1 | Error Message : Bin Location entry not found")
                End If

                oDoc.Lines.ItemCode = dsdata.Fields.Item("U_IDU_ItemCode").Value.ToString
                oDoc.Lines.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                oDoc.Lines.FromWarehouseCode = dsdata.Fields.Item("U_IDU_Whse").Value.ToString
                oDoc.Lines.WarehouseCode = "01"
                oDoc.Lines.UseBaseUnits = BoYesNoEnum.tYES
                oDoc.Lines.UoMEntry = Control._retRstField(String.Format("select ""UomEntry"" from ""OUOM"" where ""UomCode""='" & dsdata.Fields.Item("U_IDU_unitMsr").Value.ToString & "'"))
                'oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse
                'oDoc.Lines.BinAllocations.BinAbsEntry = absEntry
                'oDoc.Lines.BinAllocations.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                'oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0
                'oDoc.Lines.BinAllocations.Add()
                oDoc.Lines.BatchNumbers.BatchNumber = dsdata.Fields.Item("U_IDU_BatchNo").Value.ToString
                oDoc.Lines.BatchNumbers.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                'oDoc.Lines.BatchNumbers.BaseLineNumber = 0
                oDoc.Lines.BatchNumbers.Add()
                oDoc.Lines.Add()


                dsdata.MoveNext()
            Next

            lRetCode = oDoc.Add()
            If lRetCode = 0 Then
                Dim docentry As Integer = B1Connections.diCompany.GetNewObjectKey()
                Dim docnum As String = Control._retRstField("select ""DocNum"" from ""OWTR"" where ""DocEntry""=" & docentry)

                Control._executeQuery("update ""@ODLS"" set ""U_IDU_IT_Number""='" & docnum & "', ""U_IDU_IT_Entry""='" & docentry & "' where ""DocEntry""=" & dsentry)
                Return True
            Else
                B1Connections.diCompany.GetLastError(errorCode, errorMessage)
                Throw New Exception("Error Code : " + errorCode.ToString + " | " + "Error Message : " + errorMessage)
                Return False
            End If


        Catch ex As Exception
            Throw New Exception("Transferto01|" + ex.Message.ToString)
        End Try
    End Function

    Public Shared Function Transferfrom01(ByRef dsentry As String) As Boolean
        Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
        Dim lRetCode As Integer
        Dim errorCode As Integer
        Dim errorMessage As String = ""
        oDoc = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

        Dim dsdata As Recordset = Control._recordSet("call _IDU_DLS_GETDLSBYID('" & dsentry & "')")
        Try
            dsdata.MoveFirst()

            oDoc.DocDate = dsdata.Fields.Item("U_IDU_DocDate").Value.ToString
            oDoc.TaxDate = dsdata.Fields.Item("U_IDU_DocDate").Value.ToString
            oDoc.FromWarehouse = "01"
            oDoc.ToWarehouse = dsdata.Fields.Item("U_IDU_Whse").Value.ToString
            oDoc.UserFields.Fields.Item("U_IDU_DSNo").Value = dsentry

            For index = 1 To dsdata.RecordCount
                Dim absEntry As String = Control._retRstField("SELECT T0.""AbsEntry"" FROM ""OBIN"" T0 WHERE T0.""BinCode"" = '" & dsdata.Fields.Item("U_IDU_BIN_Loc").Value.ToString & "'")

                If absEntry = "" Then
                    Throw New Exception("Error Code : -1 | Error Message : Bin Location entry not found")
                End If

                oDoc.Lines.ItemCode = dsdata.Fields.Item("U_IDU_ItemCode").Value.ToString
                oDoc.Lines.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                oDoc.Lines.FromWarehouseCode = "01"
                oDoc.Lines.WarehouseCode = dsdata.Fields.Item("U_IDU_Whse").Value.ToString
                oDoc.Lines.UseBaseUnits = BoYesNoEnum.tYES
                oDoc.Lines.UoMEntry = Control._retRstField(String.Format("select ""UomEntry"" from ""OUOM"" where ""UomCode""='" & dsdata.Fields.Item("U_IDU_unitMsr").Value.ToString & "'"))
                oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse
                oDoc.Lines.BinAllocations.BinAbsEntry = absEntry
                oDoc.Lines.BinAllocations.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0
                oDoc.Lines.BinAllocations.Add()
                oDoc.Lines.BatchNumbers.BatchNumber = dsdata.Fields.Item("U_IDU_BatchNo").Value.ToString
                oDoc.Lines.BatchNumbers.Quantity = dsdata.Fields.Item("U_IDU_Quantity").Value.ToString
                'oDoc.Lines.BatchNumbers.BaseLineNumber = 0
                oDoc.Lines.BatchNumbers.Add()
                oDoc.Lines.Add()


                dsdata.MoveNext()
            Next

            lRetCode = oDoc.Add()
            If lRetCode = 0 Then
                Dim docentry As Integer = B1Connections.diCompany.GetNewObjectKey()
                Dim docnum As String = Control._retRstField("select ""DocNum"" from ""OWTR"" where ""DocEntry""=" & docentry)

                Control._executeQuery("update ""@ODLS"" set ""U_IDU_IT_Number""='" & docnum & "', ""U_IDU_IT_Entry""='" & docentry & "' where ""DocEntry""=" & dsentry)
                Return True
            Else
                B1Connections.diCompany.GetLastError(errorCode, errorMessage)
                Throw New Exception("Error Code : " + errorCode.ToString + " | " + "Error Message : " + errorMessage)
                Return False
            End If


        Catch ex As Exception
            Throw New Exception("Transferfrom01|" + ex.Message.ToString)
        End Try
    End Function
End Class
