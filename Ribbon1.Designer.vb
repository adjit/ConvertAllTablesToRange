Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.convertRangeTab = Me.Factory.CreateRibbonTab
        Me.rangeRibbonGroup = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.convertRangeTab.SuspendLayout()
        Me.rangeRibbonGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'convertRangeTab
        '
        Me.convertRangeTab.Groups.Add(Me.rangeRibbonGroup)
        Me.convertRangeTab.Label = "Convert Tables To Range"
        Me.convertRangeTab.Name = "convertRangeTab"
        '
        'rangeRibbonGroup
        '
        Me.rangeRibbonGroup.Items.Add(Me.Button1)
        Me.rangeRibbonGroup.Label = "Convert to Range"
        Me.rangeRibbonGroup.Name = "rangeRibbonGroup"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageName = "rangeImage"
        Me.Button1.Label = "Convert All Tables To Range"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.convertRangeTab)
        Me.convertRangeTab.ResumeLayout(False)
        Me.convertRangeTab.PerformLayout()
        Me.rangeRibbonGroup.ResumeLayout(False)
        Me.rangeRibbonGroup.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents convertRangeTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents rangeRibbonGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
