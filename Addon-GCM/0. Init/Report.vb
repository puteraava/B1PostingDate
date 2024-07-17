Imports SAPbobsCOM
Imports B1WizardBase
Imports System.IO
Imports SAPbouiCOM

Public Class Report

    Public Shared Sub _ALL()
        _DLVR_SCH()
    End Sub

    Public Shared Sub _DLVR_SCH()
        Dim newType As ReportType
        Dim newTypeParam As ReportTypeParams = Nothing
        Dim newReportParam As ReportLayoutParams
        Dim rptTypeService As ReportTypesService = DirectCast(B1Connections.diCompany.GetCompanyService.GetBusinessService( _
                ServiceTypes.ReportTypesService), ReportTypesService)

        'ADD NEW REPORT TYPE
        Dim typeCode As String = Control._retRstField("SELECT T0.""CODE"" FROM ""RTYP"" T0 WHERE T0.""MNU_ID"" ='DLVR_SCH'")
        If typeCode = "" Then
            newType = DirectCast(rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportType), ReportType)
            With newType
                .TypeName = "Delivery Schedule"
                .AddonName = "Delivery Schedule"
                .AddonFormType = "Delivery Schedule"
                .MenuID = "DLVR_SCH"
            End With
            newTypeParam = rptTypeService.AddReportType(newType)
            typeCode = newTypeParam.TypeCode
        End If

        'ADD NEW REPORT LAYOUT
        If Control._retRstField(String.Format("SELECT COUNT(""DocCode"") FROM ""RDOC"" T0 WHERE T0.""TypeCode"" = '{0}'", typeCode)) = "0" Then
            Dim rptService As ReportLayoutsService = DirectCast(B1Connections.diCompany.GetCompanyService.GetBusinessService( _
                    ServiceTypes.ReportLayoutsService), ReportLayoutsService)
            Dim newReport As ReportLayout = DirectCast(rptService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout), ReportLayout)
            With newReport
                .Author = B1Connections.diCompany.UserName
                .Category = ReportLayoutCategoryEnum.rlcCrystal
                .Name = "Delivery Schedule"
                .TypeCode = typeCode
            End With
            newReportParam = rptService.AddReportLayout(newReport)

            'SET REPORT LAYOUT TO REPORT TYPE.
            Dim updateTypeParam As SAPbobsCOM.ReportTypeParams
            updateTypeParam = GetReportTypeParams(typeCode)
            newType = rptTypeService.GetReportType(updateTypeParam)
            newType.DefaultReportLayout = newReportParam.LayoutCode
            rptTypeService.UpdateReportType(newType)

            'LINK REPORTY LAYOUT TO CRYSTAL REPORT
            Dim oBlobParams As BlobParams = DirectCast(B1Connections.diCompany.GetCompanyService.GetDataInterface( _
                    CompanyServiceDataInterfaces.csdiBlobParams), BlobParams)
            With oBlobParams
                .Table = "RDOC"
                .Field = "Template"
            End With
            Dim oKeySegment As BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
            With oKeySegment
                .Name = "DocCode"
                .Value = newReportParam.LayoutCode
            End With
        End If
    End Sub

    

    Public Shared Function GetReportTypeParams(ByVal ReportTypeCode As String) As SAPbobsCOM.ReportTypeParams
        Dim oReportTypeService As ReportTypesService = CType(B1Connections.diCompany.GetCompanyService.GetBusinessService(ServiceTypes.ReportTypesService), ReportTypesService)
        Dim oReportTypeParams As ReportTypeParams = Nothing
        Dim oReportType As ReportType = Nothing
        Dim oReportTypesParams As ReportTypesParams = oReportTypeService.GetReportTypeList()

        For i As Integer = 0 To oReportTypesParams.Count - 1
            If (oReportTypesParams.Item(i).TypeCode = ReportTypeCode) Then
                oReportTypeParams = oReportTypesParams.Item(i)
                Exit For
            End If
        Next

        Return oReportTypeParams
    End Function

End Class
