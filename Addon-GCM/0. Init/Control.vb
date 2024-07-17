Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Linq

Public Class Control
    Public Shared Function _retRstField(ByVal ssql As String) As String
        Dim rs As SAPbobsCOM.Recordset
        Dim temp As String
        rs = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        rs.DoQuery(ssql)
        If rs.EoF = True Then
            temp = ""
        Else
            temp = rs.Fields.Item(0).Value.ToString
        End If
        rs = Nothing
        Return temp
    End Function

    Public Shared Sub _executeQuery(ByVal ssql As String)
        Dim RS As SAPbobsCOM.Recordset

        RS = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery(ssql)
        RS = Nothing
    End Sub

    Public Shared Function _recordSet(ByVal ssql As String) As Recordset
        Dim RS As SAPbobsCOM.Recordset

        RS = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery(ssql)

        Return RS
    End Function

    Public Shared Sub _setDocumentNumber(ByVal controller As String, ByVal oForm As Form, ByVal series As String)
        Dim oText As SAPbouiCOM.EditText = oForm.Items.Item(controller).Specific

        oText.Value = oForm.BusinessObject.GetNextSerialNumber(series, "DLVR_SCH")
    End Sub

    Public Shared Sub _setSeries(ByVal controller As String, ByVal oForm As Form, ByVal mode As BoSeriesMode)
        Dim oCombo As SAPbouiCOM.ComboBox = oForm.Items.Item(controller).Specific
        Dim oDate As EditText = oForm.Items.Item("PostDate").Specific

        If oCombo.ValidValues.Count > 0 Then
            For index = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(index, BoSearchKey.psk_Index)
            Next
        End If

        If oDate.Value <> "" Then
            Dim oRs As Recordset = _recordSet("select ""Series"", ""SeriesName"" from nnm1 where ""ObjectCode"" = 'DLVR_SCH' and ""Indicator"" = '" & Date.ParseExact(oDate.Value, "yyyyMMdd", Globalization.CultureInfo.InvariantCulture).Year & "'")

            Do Until oRs.EoF
                oCombo.ValidValues.Add(oRs.Fields.Item(0).Value, oRs.Fields.Item(1).Value)

                oRs.MoveNext()
            Loop
        Else
            oCombo.ValidValues.LoadSeries("DLVR_SCH", mode)
        End If

        oCombo.ExpandType = BoExpandType.et_DescriptionOnly

        oCombo.Select(0, BoSearchKey.psk_Index)
    End Sub

    Public Shared Sub _formStatus(ByRef oForm As Form)
        Dim oMatrix As Matrix = CType(oForm.Items.Item("mtx_ds").Specific, Matrix)
        Dim status As ComboBox = CType(oForm.Items.Item("DocStatus").Specific, ComboBox)
        Dim aprstatus As ComboBox = CType(oForm.Items.Item("AprStatus").Specific, ComboBox)
        Dim Series As ComboBox = CType(oForm.Items.Item("Series").Specific, ComboBox)
        Dim DocStatus As ComboBox = CType(oForm.Items.Item("DocStatus").Specific, ComboBox)

        If oForm.Mode = BoFormMode.fm_ADD_MODE Then

            Control._setSeries("Series", oForm, BoSeriesMode.sf_Add)
            Control._setDocumentNumber("DocNum", oForm, Series.Selected.Value)

            status.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue)
            aprstatus.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oForm.Items.Item("btnSO").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("sonum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("PostDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Transptr").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Truck").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("CustCode").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True)

            oMatrix.Columns.Item("check").Editable = True

        ElseIf (oForm.Mode = BoFormMode.fm_OK_MODE And status.Selected.Value = "O") Then

            oMatrix.Columns.Item("check").Editable = False

            oForm.Items.Item("Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("PostDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Transptr").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Truck").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("CustCode").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)

            oForm.Items.Item("FilterPI").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("FilterGrd").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Item_17").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Item_20").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)

            If aprstatus.Selected.Value = "W" Then
                oForm.Items.Item("btnSO").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
            ElseIf aprstatus.Selected.Value = "R" Then
                oForm.Items.Item("btnSO").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            Else
                oForm.Items.Item("btnSO").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
                oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            End If

        ElseIf (oForm.Mode = BoFormMode.fm_OK_MODE And status.Selected.Value <> "O") Then
            oForm.Items.Item("CustCode").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("CustName").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Transptr").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Truck").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("PostDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("sonum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Remarks").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocTotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Grtotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("AprStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False)

        ElseIf oForm.Mode = BoFormMode.fm_FIND_MODE Then
            Control._setSeries("Series", oForm, BoSeriesMode.sf_View)
            oForm.Items.Item("CustCode").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("CustName").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Transptr").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Truck").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("PostDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("sonum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Remarks").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DocTotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Grtotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("AprStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
            oForm.Items.Item("Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)

            oForm.Items.Item("FilterPI").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("FilterGrd").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Item_17").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Item_20").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False)

        ElseIf oForm.Mode = BoFormMode.fm_UPDATE_MODE And (aprstatus.Selected.Value = "R" Or aprstatus.Selected.Value = "A") Then
            oForm.Items.Item("Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("CustCode").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("CustName").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DelTerm").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Transptr").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Truck").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("truckqty").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("PostDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("sonum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocTotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("LCLAmnt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("Grtotal").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("AprStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("DocStatus").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("SoDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("mtx_ds").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
        End If
    End Sub

    Public Shared Sub _setMenuData(form As Form, status As Boolean)
        'form.EnableMenu("1292", status)
        'form.EnableMenu("1293", status)
    End Sub

    Public Shared Function _chooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.DataTable
        Dim oCflEvent As SAPbouiCOM.ChooseFromListEvent
        Dim oDataTable As SAPbouiCOM.DataTable
        oCflEvent = pVal
        oDataTable = oCflEvent.SelectedObjects

        Return oDataTable
    End Function

    Public Shared Sub _createFms(ByVal oFms As oFMS, ByVal refresh As Boolean, ByVal refreshField As String)
        Try
            'Create Queries Category
            Dim _categoryId As String = Control._retRstField("SELECT ""CategoryId"" FROM OQCN Where ""CatName"" = 'Delivery Schedule'")
            If _categoryId = "" Then
                Dim oCategory As SAPbobsCOM.QueryCategories = B1Connections.diCompany.GetBusinessObject(BoObjectTypes.oQueryCategories)
                oCategory.Name = "Delivery Schedule"
                oCategory.Permissions = "YYYYYYYYYYYYYYY"
                oCategory.Add()
                _categoryId = B1Connections.diCompany.GetNewObjectKey().Split(vbTab).GetValue(0)
            End If

            'Create User Queries
            Dim _queryId As String = Control._retRstField("select ""IntrnalKey"" from OUQR Where ""QName"" = '" & oFms.Name & "' and ""QCategory"" = '" & _categoryId & "'")
            If _queryId = "" Then
                Dim oQuery As SAPbobsCOM.UserQueries = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
                oQuery.Query = oFms.Query
                oQuery.QueryCategory = _categoryId
                oQuery.QueryDescription = oFms.Name
                oQuery.Add()
                _queryId = B1Connections.diCompany.GetNewObjectKey().Split(vbTab).GetValue(0)
            End If

            'Create FMS
            Dim _formatedId = Control._retRstField("select ""IndexID"" from CSHS where ""FormID""='" & oFms.Form & "' AND ""ItemID""='" & oFms.ItemId & "' AND ""ColID""='" & oFms.ColId & "'")
            Dim oFormatted As SAPbobsCOM.FormattedSearches = B1Connections.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
            If _formatedId = "" Then
                oFormatted.FormID = oFms.Form
                oFormatted.ItemID = oFms.ItemId
                oFormatted.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                oFormatted.ColumnID = oFms.ColId
                oFormatted.QueryID = _queryId

                If Not refresh Then
                    oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                Else
                    oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tYES
                    oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.FieldID = refreshField
                End If

                oFormatted.Add()
            Else
                oFormatted.GetByKey(_formatedId)
                oFormatted.FormID = oFms.Form
                oFormatted.ItemID = oFms.ItemId
                oFormatted.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                oFormatted.ColumnID = oFms.ColId
                oFormatted.QueryID = _queryId

                If Not refresh Then
                    oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                Else
                    oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tYES
                    oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                    oFormatted.FieldID = refreshField
                End If

                oFormatted.Update()
            End If
        Catch ex As Exception
            B1Connections.theAppl.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Public Shared Sub _beforeChooseFromCustomer(ByVal oForm As SAPbouiCOM.Form, ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID

        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
        Dim oConds As Conditions = B1Connections.theAppl.CreateObject(BoCreatableObjectType.cot_Conditions)
        oConds = oCFL.GetConditions()
        oConds = Nothing
        oCFL.SetConditions(oConds)
        oConds = oCFL.GetConditions()
        Dim oCondition As Condition = Nothing

        oCondition = oConds.Add()
        oCondition.Alias = "CardType"
        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCondition.CondVal = "C"
        oCFL.SetConditions(oConds)
    End Sub
End Class
