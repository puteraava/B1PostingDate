Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Threading
Imports System.Globalization

Public Class oFMS
    Private _form As String
    Public Property Form() As String
        Get
            Return _form
        End Get
        Set(ByVal value As String)
            _form = value
        End Set
    End Property

    Private _query As String
    Public Property Query() As String
        Get
            Return _query
        End Get
        Set(ByVal value As String)
            _query = value
        End Set
    End Property

    Private _name As String
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property


    Private _itemId As String
    Public Property ItemId() As String
        Get
            Return _itemId
        End Get
        Set(ByVal value As String)
            _itemId = value
        End Set
    End Property

    Private _colId As String
    Public Property ColId() As String
        Get
            Return _colId
        End Get
        Set(ByVal value As String)
            _colId = value
        End Set
    End Property

End Class

Public Class FMS
    Public Shared Sub _ALL()
        _TRANSPORTER()
        _TRUCK()
        _BATCH()
        _BIN()
    End Sub

    Public Shared Sub _TRANSPORTER()
        Dim oFms As New oFMS With {.Name = "Transporter", .Query = "select ""U_IDU_TRANSPORTER_CODE"", ""U_IDU_TRANSPORTER_NAME"", ""Code"", ""Name"" from ""@IDU_TRANSPORTER""", .ItemId = "Transptr", .ColId = "-1", .Form = "DLVR_SCH"}

        Control._createFms(oFms, False, "")
    End Sub

    Public Shared Sub _TRUCK()
        Dim oFms As New oFMS With {.Name = "Truck", .Query = "select ""Code"", ""U_IDU_TRUCK_TYPE"", ""U_IDU_DEFAULT_WEIGHT"",  ""Name"" from ""@IDU_TRUCK""", .ItemId = "Truck", .ColId = "-1", .Form = "DLVR_SCH"}

        Control._createFms(oFms, False, "")
    End Sub

    Public Shared Sub _BATCH()
        Dim oFms As New oFMS With {.Name = "Batch", .Query = "select t2.""DistNumber"",  t4.""BinCode"", t0.""ItemCode"",  t0.""Docdate"",sum(t1.""Quantity"") as 'Quantity' From oitl t0 join itl1 t1 on t0.""LogEntry"" = t1.""LogEntry"" join obtn t2 on t1.""SysNumber"" = t2.""SysNumber"" join obtl t3 on t0.""LogEntry"" = t3.""ITLEntry"" join obin t4 on t3.""BinAbs"" = t4.""AbsEntry"" where t0.""ItemCode""=$[$mtx_ds.ItemCod.0] group by t0.ItemCode, t2.DistNumber, t4.BinCode, t0.Docdate order by t0.DocDate asc", .ItemId = "Batch", .ColId = "-1", .Form = "DLVR_SCH"}

        Control._createFms(oFms, False, "")
    End Sub

    Public Shared Sub _BIN()
        Dim oFms As New oFMS With {.Name = "Bin", .Query = "select t4.""BinCode"", t2.""DistNumber"", t0.""ItemCode"",  t0.""Docdate"",sum(t1.""Quantity"") as 'Quantity' From oitl t0 join itl1 t1 on t0.""LogEntry"" = t1.""LogEntry"" join obtn t2 on t1.""SysNumber"" = t2.""SysNumber"" join OBTL t3 on t0.""LogEntry"" = t3.""ITLEntry"" join obin t4 on t3.""BinAbs"" = t4.""AbsEntry"" where t0.""ItemCode""=$[$mtx_ds.ItemCod.0]  and t2.DistNumber=$[$mtx_ds.Batch.0] group by t0.ItemCode, t2.DistNumber, t4.BinCode, t0.Docdate order by t0.DocDate asc", .ItemId = "BinLoc", .ColId = "Batch", .Form = "DLVR_SCH"}

        Control._createFms(oFms, False, "")
    End Sub
End Class

