Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Collections.Generic

Public Class APR_DLS__Sub
    Public Shared Sub Loadgrid(ByRef form As SAPbouiCOM.Form)
        Try
            Dim sql As String = "CALL ""_IDU_DLS_LISTAPPROVAL"""

            form.DataSources.DataTables.Item("DT_0").ExecuteQuery(sql)
            Dim grid As Grid = form.Items.Item("Item_0").Specific

            grid.DataTable = form.DataSources.DataTables.Item("DT_0")
            grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            'set editable false
            grid.Columns.Item(0).Editable = False
            grid.Columns.Item(1).Editable = False
            grid.Columns.Item(2).Editable = False
            grid.Columns.Item(3).Editable = False
            grid.Columns.Item(4).Editable = False

        Catch ex As Exception
            Throw New Exception("LoadGrid|" + ex.Message.ToString)
        End Try
    End Sub

    Public Shared Function _getGridRowIndex(ByVal Grid As SAPbouiCOM.Grid, ByVal RowIndex As Integer) As Integer
        Dim i As Integer
        Dim CountIsExpanded As Integer = 0
        For i = 0 To Grid.Rows.Count - 1
            If i = RowIndex Then
                Return RowIndex - CountIsExpanded
            End If
            If Grid.Rows.IsLeaf(i) = False Then
                CountIsExpanded = CountIsExpanded + 1
            End If

        Next i
        Return -1
    End Function

    Public Shared Sub updatePI(ByVal dsentry As String)
        Dim dsdata As Recordset = Control._recordSet("call _IDU_DLS_GETDLSBYID('" & dsentry & "')")
        Dim currentbalance As Double
        Dim qty As Double
        Dim pientry As Integer
        Dim itemcode As String
        dsdata.MoveFirst()
        For index = 1 To dsdata.RecordCount
            pientry = CInt(dsdata.Fields.Item("U_IDU_PIEntry").Value.ToString)
            itemcode = dsdata.Fields.Item("U_IDU_ItemCode").Value.ToString
            qty = CDbl(dsdata.Fields.Item("U_IDU_Quantity").Value.ToString)
            currentbalance = CDbl(Control._retRstField("select ""U_IDU_PI_Bal"" from QUT1 where ""DocEntry""=" & pientry & " and ""ItemCode""='" & itemcode & "' "))
            'update balance
            Dim sql As String = "update ""QUT1"" set ""U_IDU_PI_Bal""=" & currentbalance + qty & " where ""DocEntry""=" & pientry & " and ""ItemCode"" ='" & itemcode & "'"
            Control._executeQuery(sql)

            'open PI
            Control._executeQuery("update OQUT set ""DocStatus""='O' where ""DocEntry""=" & pientry)
            Control._executeQuery("update QUT1 set ""LineStatus""='O' where ""DocEntry""=" & pientry)
            dsdata.MoveNext()
        Next

    End Sub
End Class
