Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class APR_DLS__Menu
    Inherits B1XmlFormMenu

    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Public Sub New()
        MyBase.New()
        MenuUID = "APR_DLS"
        'GENERATED CODE
        Me.LoadXml("APR_DLS.xml")
    End Sub

    <B1Listener(BoEventTypes.et_MENU_CLICK, True)> _
    Public Overridable Function OnBeforeMenuClick(ByVal pVal As MenuEvent) As Boolean
        'GENERATED CODE

        Me.LoadForm()
        Dim oForm As Form = B1Connections.theAppl.Forms.ActiveForm

        APR_DLS__Sub.Loadgrid(oForm)

    End Function

    '<B1Listener(BoEventTypes.et_MENU_CLICK, False, {"APR_DLS"})> _
    'Public Overridable Sub OnAfterMenuClick(ByVal pVal As MenuEvent)
    '    Dim Form As Form = B1Connections.theAppl.Forms.ActiveForm
    '    'Dim Form As Form = B1Connections.theAppl.Forms.Item(B1Connections.theAppl.Forms.ActiveForm.UniqueID.ToString)
    '    Dim mtx_ds As SAPbouiCOM.Matrix = Form.Items.Item("mtx_apr").Specific

    '    Dim oRs As Recordset
    '    Dim sql As String = "select * from [@ODLS] where Status='O' and U_IDU_Appr_Status='W'"
    '    oRs = Control._recordSet(sql)

    '    mtx_ds.Clear()

    '    Do Until oRs.EoF
    '        mtx_ds.AddRow(1, mtx_ds.RowCount)

    '        mtx_ds.SetCellWithoutValidation(mtx_ds.RowCount, "DsNo", oRs.Fields.Item("DocNum").Value.ToString)
    '        mtx_ds.SetCellWithoutValidation(mtx_ds.RowCount, "CusCode", oRs.Fields.Item("U_IDU_CardCode").Value.ToString)
    '        mtx_ds.SetCellWithoutValidation(mtx_ds.RowCount, "CusName", oRs.Fields.Item("U_IDU_CardName").Value.ToString)
    '        mtx_ds.SetCellWithoutValidation(mtx_ds.RowCount, "TtlAmnt", oRs.Fields.Item("U_IDU_GrandTotal").Value.ToString)
    '        mtx_ds.SetCellWithoutValidation(mtx_ds.RowCount, "Remark", oRs.Fields.Item("Remark").Value.ToString)
    '        oRs.MoveNext()
    '    Loop

    'End Sub
End Class
