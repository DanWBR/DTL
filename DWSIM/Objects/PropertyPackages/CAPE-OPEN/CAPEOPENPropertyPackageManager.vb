Imports CapeOpen = CAPEOPEN110
Imports DTL.DTL.SimulationObjects.PropertyPackages
Imports CAPEOPEN110

<System.Serializable()> _
Friend Class CAPEOPENPropertyPackageManager

    Implements ICapeIdentification, ICapeThermoPropertyPackageManager, ICapeUtilities
    Implements IDisposable

    Private _name, _description As String

    Sub New()
        _name = "DWSIM Thermodynamics Library (DTL)"
        _description = "Exposes DTL Property Packages to clients using CAPE-OPEN Thermodynamic Interface Definitions"
    End Sub

    Public Function GetPropertyPackage(ByVal PackageName As String) As Object Implements ICapeThermoPropertyPackageManager.GetPropertyPackage
        Dim pp As PropertyPackage = Nothing
        Select Case PackageName
            Case "Peng-Robinson (PR)"
                pp = New PengRobinsonPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescPengRobinsonPP")
            Case "Peng-Robinson-Stryjek-Vera 2 (PRSV2-M)", "Peng-Robinson-Stryjek-Vera 2 (PRSV2)"
                pp = New PRSV2PropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescPRSV2PP")
            Case "Peng-Robinson-Stryjek-Vera 2 (PRSV2-VL)"
                pp = New PRSV2VLPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescPRSV2VLPP")
            Case "Soave-Redlich-Kwong (SRK)"
                pp = New SRKPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescSoaveRedlichKwongSRK")
            Case "Peng-Robinson / Lee-Kesler (PR/LK)"
                pp = New PengRobinsonLKPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescPRLK")
            Case "UNIFAC"
                pp = New UNIFACPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescUPP")
            Case "UNIFAC-LL"
                pp = New UNIFACLLPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescUPP")
            Case "NRTL"
                pp = New NRTLPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescNRTLPP")
            Case "UNIQUAC"
                pp = New UNIQUACPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescUNIQUACPP")
            Case "Modified UNIFAC (Dortmund)"
                pp = New MODFACPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescMUPP")
            Case "Modified UNIFAC (NIST)"
                pp = New NISTMFACPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescNUPP")
            Case "Chao-Seader"
                pp = New ChaoSeaderPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescCSLKPP")
            Case "Grayson-Streed"
                pp = New GraysonStreedPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescGSLKPP")
            Case "Lee-Kesler-Plöcker"
                pp = New LKPPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescLKPPP")
            Case "Raoult's Law"
                pp = New RaoultPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescRPP")
            Case "IAPWS-IF97 Steam Tables"
                pp = New SteamTablesPropertyPackage(True)
                pp.ComponentDescription = DTL.App.GetLocalString("DescSteamTablesPP")
            Case Else
                Throw New ArgumentException("Property Package not found.")
        End Select
        If Not pp Is Nothing Then pp.ComponentName = PackageName
        Return pp
    End Function

    Public Function GetPropertyPackageList() As Object Implements ICapeThermoPropertyPackageManager.GetPropertyPackageList
        Return New String() {"Peng-Robinson (PR)", "Peng-Robinson-Stryjek-Vera 2 (PRSV2-M)", "Peng-Robinson-Stryjek-Vera 2 (PRSV2-VL)", "Soave-Redlich-Kwong (SRK)", "Peng-Robinson / Lee-Kesler (PR/LK)", _
                             "UNIFAC", "UNIFAC-LL", "Modified UNIFAC (Dortmund)", "Modified UNIFAC (NIST)", "NRTL", "UNIQUAC", _
                            "Chao-Seader", "Grayson-Streed", "Lee-Kesler-Plöcker", "Raoult's Law", "IAPWS-IF97 Steam Tables"}
    End Function

    Public Property ComponentDescription() As String Implements ICapeIdentification.ComponentDescription
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Public Property ComponentName() As String Implements ICapeIdentification.ComponentName
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Sub Edit() Implements ICapeUtilities.Edit
        Throw New Exception("Edit() not implemented.")
    End Sub

    Public Sub Initialize() Implements ICapeUtilities.Initialize
        '
    End Sub

    Public ReadOnly Property parameters() As Object Implements ICapeUtilities.parameters
        Get
            Throw New NotImplementedException
        End Get
    End Property

    Public WriteOnly Property simulationContext() As Object Implements ICapeUtilities.simulationContext
        Set(ByVal value As Object)
            'do nothing
        End Set
    End Property

    Public Sub Terminate() Implements ICapeUtilities.Terminate
        Me.Dispose()
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
