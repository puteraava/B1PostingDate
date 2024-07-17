Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.IO

Public Class SP
    Shared companyName As String = B1Connections.diCompany.CompanyDB

    Public Shared Sub __ALL()
        _IDU_DLS_BALANCEDQTY()
        _IDU_DLS_GET_PI()
        _IDU_DLS_GET_STOCKWHS()
        _IDU_DLS_GETDLSBYID()
        _IDU_DLS_LAYOUTDS()
        _IDU_DLS_LISTAPPROVAL()
    End Sub

    Public Shared Sub _IDU_DLS_GET_STOCKWHS()
        Dim ssql As String
        ssql = "select ""PROCEDURE_NAME"" from ""SYS"".""PROCEDURES"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""PROCEDURE_NAME"" = '_IDU_DLS_GET_STOCKWHS'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_GET_STOCKWHS)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub

    Public Shared Sub _IDU_DLS_GETDLSBYID()
        Dim ssql As String
        ssql = "select ""PROCEDURE_NAME"" from ""SYS"".""PROCEDURES"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""PROCEDURE_NAME"" = '_IDU_DLS_GETDLSBYID'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_GETDLSBYID)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub

    Public Shared Sub _IDU_DLS_LAYOUTDS()
        Dim ssql As String
        ssql = "select ""PROCEDURE_NAME"" from ""SYS"".""PROCEDURES"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""PROCEDURE_NAME"" = '_IDU_DLS_LAYOUTDS'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_LAYOUTDS)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub

    Public Shared Sub _IDU_DLS_LISTAPPROVAL()
        Dim ssql As String
        ssql = "select ""PROCEDURE_NAME"" from ""SYS"".""PROCEDURES"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""PROCEDURE_NAME"" = '_IDU_DLS_LISTAPPROVAL'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_LISTAPPROVAL)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub

    Public Shared Sub _IDU_DLS_GET_PI()
        Dim ssql As String
        ssql = "select ""FUNCTION_NAME"" from ""SYS"".""FUNCTIONS"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""FUNCTION_NAME"" = 'FN_REVERSE'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_GET_PI)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub

    Public Shared Sub _IDU_DLS_BALANCEDQTY()
        Dim ssql As String
        ssql = "select ""FUNCTION_NAME"" from ""SYS"".""FUNCTIONS"" where ""SCHEMA_NAME"" = '" & companyName & "' and ""FUNCTION_NAME"" = 'FN_REPLICATE'"
        If Control._retRstField(ssql) <> "" Then
            Exit Sub
        End If

        Using sr As New StreamReader(System.Windows.Forms.Application.StartupPath + "/" + My.Settings._IDU_DLS_BALANCEDQTY)
            ssql = sr.ReadToEnd()
        End Using
        Control._executeQuery(ssql)
    End Sub
End Class