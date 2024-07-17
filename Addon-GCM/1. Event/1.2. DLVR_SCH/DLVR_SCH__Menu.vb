Option Strict Off
Option Explicit On

Imports B1WizardBase
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class DLVR_SCH__Menu
    Inherits B1XmlFormMenu

    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Public Sub New()
        MyBase.New()
        MenuUID = "DLVR_SCH"
        'GENERATED CODE
        Me.LoadXml("DLVR_SCH.xml")
    End Sub

    <B1Listener(BoEventTypes.et_MENU_CLICK, True)>
    Public Overridable Function OnBeforeMenuClick(ByVal pVal As MenuEvent) As Boolean
        'GENERATED CODE
        Me.LoadForm()
        Dim oForm As Form = B1Connections.theAppl.Forms.Item(B1Connections.theAppl.Forms.ActiveForm.UniqueID.ToString)
        Control._formStatus(oForm)
        Control._setMenuData(oForm, True)
    End Function
End Class
