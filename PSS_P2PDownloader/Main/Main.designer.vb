<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.MainMenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PSSVariantsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AmericasSTRToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AmericasSSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SubVariantsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AmericasDirectToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SubVariantsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.AmericasCustomizationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SubVarianToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AmericasLogisticsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.IndirectNAPDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TMSNALAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportsLAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportsNAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DirectToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SubVariantsToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem
        Me.MaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdateRegionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdatePlantsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdateSAPBoxesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdatePOrgsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdateUsersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdatePGrpsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.lb_TNumber = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lb_User = New System.Windows.Forms.Label
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.Main_ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.Main_ToolStripProgressBar = New System.Windows.Forms.ToolStripProgressBar
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.GB_Reports = New System.Windows.Forms.GroupBox
        Me.LK_ExportVariants = New System.Windows.Forms.LinkLabel
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.GB_Maintenance = New System.Windows.Forms.GroupBox
        Me.LK_PGrps = New System.Windows.Forms.LinkLabel
        Me.LK_Users = New System.Windows.Forms.LinkLabel
        Me.LK_Regions = New System.Windows.Forms.LinkLabel
        Me.LK_POrgs = New System.Windows.Forms.LinkLabel
        Me.LK_SAP_Boxes = New System.Windows.Forms.LinkLabel
        Me.LK_Plants = New System.Windows.Forms.LinkLabel
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GB_Variants = New System.Windows.Forms.GroupBox
        Me.LK_LogDirect = New System.Windows.Forms.LinkLabel
        Me.LK_LogImpNA = New System.Windows.Forms.LinkLabel
        Me.LK_LogImpLA = New System.Windows.Forms.LinkLabel
        Me.LK_LogTMS = New System.Windows.Forms.LinkLabel
        Me.LK_LogIndNAPD = New System.Windows.Forms.LinkLabel
        Me.LK_Americas_Custom = New System.Windows.Forms.LinkLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.LK_Americas_Direct = New System.Windows.Forms.LinkLabel
        Me.LK_Americas_SS = New System.Windows.Forms.LinkLabel
        Me.LK_Americas_STR = New System.Windows.Forms.LinkLabel
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.MenuStrip.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GB_Reports.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_Maintenance.SuspendLayout()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GB_Variants.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.BackColor = System.Drawing.Color.Lavender
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MainMenuToolStripMenuItem, Me.PSSVariantsToolStripMenuItem, Me.MaintenanceToolStripMenuItem, Me.AboutToolStripMenuItem, Me.CloseToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(984, 24)
        Me.MenuStrip.Stretch = False
        Me.MenuStrip.TabIndex = 0
        Me.MenuStrip.Text = "MenuStrip1"
        '
        'MainMenuToolStripMenuItem
        '
        Me.MainMenuToolStripMenuItem.Name = "MainMenuToolStripMenuItem"
        Me.MainMenuToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.MainMenuToolStripMenuItem.Text = "Menu"
        '
        'PSSVariantsToolStripMenuItem
        '
        Me.PSSVariantsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AmericasSTRToolStripMenuItem, Me.AmericasSSToolStripMenuItem, Me.AmericasDirectToolStripMenuItem, Me.AmericasCustomizationToolStripMenuItem, Me.AmericasLogisticsToolStripMenuItem})
        Me.PSSVariantsToolStripMenuItem.Name = "PSSVariantsToolStripMenuItem"
        Me.PSSVariantsToolStripMenuItem.Size = New System.Drawing.Size(83, 20)
        Me.PSSVariantsToolStripMenuItem.Text = "PSS Variants"
        '
        'AmericasSTRToolStripMenuItem
        '
        Me.AmericasSTRToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.AmericasSTRToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.AmericasSTRToolStripMenuItem.Name = "AmericasSTRToolStripMenuItem"
        Me.AmericasSTRToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.AmericasSTRToolStripMenuItem.Text = "Americas STR"
        '
        'AmericasSSToolStripMenuItem
        '
        Me.AmericasSSToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.AmericasSSToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SubVariantsToolStripMenuItem})
        Me.AmericasSSToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.AmericasSSToolStripMenuItem.Name = "AmericasSSToolStripMenuItem"
        Me.AmericasSSToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.AmericasSSToolStripMenuItem.Text = "Americas SS"
        '
        'SubVariantsToolStripMenuItem
        '
        Me.SubVariantsToolStripMenuItem.Name = "SubVariantsToolStripMenuItem"
        Me.SubVariantsToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.SubVariantsToolStripMenuItem.Text = "Sub-Variants"
        '
        'AmericasDirectToolStripMenuItem
        '
        Me.AmericasDirectToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.AmericasDirectToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SubVariantsToolStripMenuItem1})
        Me.AmericasDirectToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.AmericasDirectToolStripMenuItem.Name = "AmericasDirectToolStripMenuItem"
        Me.AmericasDirectToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.AmericasDirectToolStripMenuItem.Text = "Americas Direct"
        '
        'SubVariantsToolStripMenuItem1
        '
        Me.SubVariantsToolStripMenuItem1.Name = "SubVariantsToolStripMenuItem1"
        Me.SubVariantsToolStripMenuItem1.Size = New System.Drawing.Size(141, 22)
        Me.SubVariantsToolStripMenuItem1.Text = "Sub-Variants"
        '
        'AmericasCustomizationToolStripMenuItem
        '
        Me.AmericasCustomizationToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.AmericasCustomizationToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SubVarianToolStripMenuItem})
        Me.AmericasCustomizationToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.AmericasCustomizationToolStripMenuItem.Name = "AmericasCustomizationToolStripMenuItem"
        Me.AmericasCustomizationToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.AmericasCustomizationToolStripMenuItem.Text = "Americas Customization"
        '
        'SubVarianToolStripMenuItem
        '
        Me.SubVarianToolStripMenuItem.Name = "SubVarianToolStripMenuItem"
        Me.SubVarianToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.SubVarianToolStripMenuItem.Text = "Sub-Variants"
        '
        'AmericasLogisticsToolStripMenuItem
        '
        Me.AmericasLogisticsToolStripMenuItem.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.AmericasLogisticsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.IndirectNAPDToolStripMenuItem, Me.TMSNALAToolStripMenuItem, Me.ImportsLAToolStripMenuItem, Me.ImportsNAToolStripMenuItem, Me.DirectToolStripMenuItem})
        Me.AmericasLogisticsToolStripMenuItem.Name = "AmericasLogisticsToolStripMenuItem"
        Me.AmericasLogisticsToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.AmericasLogisticsToolStripMenuItem.Text = "Americas Logistics"
        '
        'IndirectNAPDToolStripMenuItem
        '
        Me.IndirectNAPDToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.IndirectNAPDToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.IndirectNAPDToolStripMenuItem.Name = "IndirectNAPDToolStripMenuItem"
        Me.IndirectNAPDToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.IndirectNAPDToolStripMenuItem.Text = "Indirect NAPD"
        '
        'TMSNALAToolStripMenuItem
        '
        Me.TMSNALAToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.TMSNALAToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.TMSNALAToolStripMenuItem.Name = "TMSNALAToolStripMenuItem"
        Me.TMSNALAToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.TMSNALAToolStripMenuItem.Text = "TMS NALA"
        '
        'ImportsLAToolStripMenuItem
        '
        Me.ImportsLAToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.ImportsLAToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.ImportsLAToolStripMenuItem.Name = "ImportsLAToolStripMenuItem"
        Me.ImportsLAToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.ImportsLAToolStripMenuItem.Text = "Imports LA"
        '
        'ImportsNAToolStripMenuItem
        '
        Me.ImportsNAToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.ImportsNAToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.ImportsNAToolStripMenuItem.Name = "ImportsNAToolStripMenuItem"
        Me.ImportsNAToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.ImportsNAToolStripMenuItem.Text = "Imports NA"
        '
        'DirectToolStripMenuItem
        '
        Me.DirectToolStripMenuItem.BackColor = System.Drawing.Color.MidnightBlue
        Me.DirectToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SubVariantsToolStripMenuItem3})
        Me.DirectToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.DirectToolStripMenuItem.Name = "DirectToolStripMenuItem"
        Me.DirectToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.DirectToolStripMenuItem.Text = "Direct"
        '
        'SubVariantsToolStripMenuItem3
        '
        Me.SubVariantsToolStripMenuItem3.Name = "SubVariantsToolStripMenuItem3"
        Me.SubVariantsToolStripMenuItem3.Size = New System.Drawing.Size(141, 22)
        Me.SubVariantsToolStripMenuItem3.Text = "Sub-Variants"
        '
        'MaintenanceToolStripMenuItem
        '
        Me.MaintenanceToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpdateRegionsToolStripMenuItem, Me.UpdatePlantsToolStripMenuItem, Me.UpdateSAPBoxesToolStripMenuItem, Me.UpdatePOrgsToolStripMenuItem, Me.UpdateUsersToolStripMenuItem, Me.UpdatePGrpsToolStripMenuItem})
        Me.MaintenanceToolStripMenuItem.Enabled = False
        Me.MaintenanceToolStripMenuItem.Name = "MaintenanceToolStripMenuItem"
        Me.MaintenanceToolStripMenuItem.Size = New System.Drawing.Size(88, 20)
        Me.MaintenanceToolStripMenuItem.Text = "Maintenance"
        '
        'UpdateRegionsToolStripMenuItem
        '
        Me.UpdateRegionsToolStripMenuItem.Name = "UpdateRegionsToolStripMenuItem"
        Me.UpdateRegionsToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdateRegionsToolStripMenuItem.Text = "Update Regions"
        '
        'UpdatePlantsToolStripMenuItem
        '
        Me.UpdatePlantsToolStripMenuItem.Name = "UpdatePlantsToolStripMenuItem"
        Me.UpdatePlantsToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdatePlantsToolStripMenuItem.Text = "Update Plants"
        '
        'UpdateSAPBoxesToolStripMenuItem
        '
        Me.UpdateSAPBoxesToolStripMenuItem.Name = "UpdateSAPBoxesToolStripMenuItem"
        Me.UpdateSAPBoxesToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdateSAPBoxesToolStripMenuItem.Text = "Update SAP Boxes"
        '
        'UpdatePOrgsToolStripMenuItem
        '
        Me.UpdatePOrgsToolStripMenuItem.Name = "UpdatePOrgsToolStripMenuItem"
        Me.UpdatePOrgsToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdatePOrgsToolStripMenuItem.Text = "Update POrgs"
        '
        'UpdateUsersToolStripMenuItem
        '
        Me.UpdateUsersToolStripMenuItem.Name = "UpdateUsersToolStripMenuItem"
        Me.UpdateUsersToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdateUsersToolStripMenuItem.Text = "Update Users"
        '
        'UpdatePGrpsToolStripMenuItem
        '
        Me.UpdatePGrpsToolStripMenuItem.Name = "UpdatePGrpsToolStripMenuItem"
        Me.UpdatePGrpsToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.UpdatePGrpsToolStripMenuItem.Text = "Update PGrps"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.CloseToolStripMenuItem.Text = "Close"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Controls.Add(Me.PictureBox8)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 24)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(984, 70)
        Me.Panel2.TabIndex = 4
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.lb_TNumber)
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Controls.Add(Me.lb_User)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel4.Location = New System.Drawing.Point(763, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(221, 70)
        Me.Panel4.TabIndex = 29
        '
        'lb_TNumber
        '
        Me.lb_TNumber.AutoSize = True
        Me.lb_TNumber.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lb_TNumber.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lb_TNumber.Location = New System.Drawing.Point(78, 35)
        Me.lb_TNumber.Name = "lb_TNumber"
        Me.lb_TNumber.Size = New System.Drawing.Size(69, 15)
        Me.lb_TNumber.TabIndex = 3
        Me.lb_TNumber.Text = "User TNumber"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label2.Location = New System.Drawing.Point(17, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "TNumber: "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(17, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Welcome: "
        '
        'lb_User
        '
        Me.lb_User.AutoSize = True
        Me.lb_User.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lb_User.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lb_User.Location = New System.Drawing.Point(78, 20)
        Me.lb_User.Name = "lb_User"
        Me.lb_User.Size = New System.Drawing.Size(55, 15)
        Me.lb_User.TabIndex = 0
        Me.lb_User.Text = "User Name"
        '
        'PictureBox8
        '
        Me.PictureBox8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(984, 70)
        Me.PictureBox8.TabIndex = 0
        Me.PictureBox8.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Main_ToolStripStatusLabel, Me.Main_ToolStripProgressBar})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 564)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(982, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Main_ToolStripStatusLabel
        '
        Me.Main_ToolStripStatusLabel.BackColor = System.Drawing.SystemColors.Control
        Me.Main_ToolStripStatusLabel.Name = "Main_ToolStripStatusLabel"
        Me.Main_ToolStripStatusLabel.Size = New System.Drawing.Size(865, 17)
        Me.Main_ToolStripStatusLabel.Spring = True
        Me.Main_ToolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Main_ToolStripProgressBar
        '
        Me.Main_ToolStripProgressBar.Name = "Main_ToolStripProgressBar"
        Me.Main_ToolStripProgressBar.Size = New System.Drawing.Size(100, 16)
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 564)
        Me.Splitter1.TabIndex = 27
        Me.Splitter1.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Transparent
        Me.Panel3.Controls.Add(Me.Panel5)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel3.Location = New System.Drawing.Point(688, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(294, 564)
        Me.Panel3.TabIndex = 28
        '
        'Panel5
        '
        Me.Panel5.BackgroundImage = CType(resources.GetObject("Panel5.BackgroundImage"), System.Drawing.Image)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(0, 496)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(294, 68)
        Me.Panel5.TabIndex = 0
        '
        'GB_Reports
        '
        Me.GB_Reports.Controls.Add(Me.LK_ExportVariants)
        Me.GB_Reports.Controls.Add(Me.PictureBox4)
        Me.GB_Reports.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Reports.ForeColor = System.Drawing.Color.MidnightBlue
        Me.GB_Reports.Location = New System.Drawing.Point(299, 85)
        Me.GB_Reports.Name = "GB_Reports"
        Me.GB_Reports.Size = New System.Drawing.Size(261, 98)
        Me.GB_Reports.TabIndex = 31
        Me.GB_Reports.TabStop = False
        Me.GB_Reports.Text = "Reports"
        '
        'LK_ExportVariants
        '
        Me.LK_ExportVariants.AutoSize = True
        Me.LK_ExportVariants.Location = New System.Drawing.Point(96, 30)
        Me.LK_ExportVariants.Name = "LK_ExportVariants"
        Me.LK_ExportVariants.Size = New System.Drawing.Size(119, 13)
        Me.LK_ExportVariants.TabIndex = 1
        Me.LK_ExportVariants.TabStop = True
        Me.LK_ExportVariants.Text = "Export Variants to Excel"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(21, 19)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(66, 50)
        Me.PictureBox4.TabIndex = 0
        Me.PictureBox4.TabStop = False
        '
        'GB_Maintenance
        '
        Me.GB_Maintenance.Controls.Add(Me.LK_PGrps)
        Me.GB_Maintenance.Controls.Add(Me.LK_Users)
        Me.GB_Maintenance.Controls.Add(Me.LK_Regions)
        Me.GB_Maintenance.Controls.Add(Me.LK_POrgs)
        Me.GB_Maintenance.Controls.Add(Me.LK_SAP_Boxes)
        Me.GB_Maintenance.Controls.Add(Me.LK_Plants)
        Me.GB_Maintenance.Controls.Add(Me.PictureBox5)
        Me.GB_Maintenance.Enabled = False
        Me.GB_Maintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Maintenance.ForeColor = System.Drawing.Color.MidnightBlue
        Me.GB_Maintenance.Location = New System.Drawing.Point(299, 189)
        Me.GB_Maintenance.Name = "GB_Maintenance"
        Me.GB_Maintenance.Size = New System.Drawing.Size(261, 235)
        Me.GB_Maintenance.TabIndex = 32
        Me.GB_Maintenance.TabStop = False
        Me.GB_Maintenance.Text = "Maintenance"
        '
        'LK_PGrps
        '
        Me.LK_PGrps.AutoSize = True
        Me.LK_PGrps.Location = New System.Drawing.Point(93, 155)
        Me.LK_PGrps.Name = "LK_PGrps"
        Me.LK_PGrps.Size = New System.Drawing.Size(74, 13)
        Me.LK_PGrps.TabIndex = 45
        Me.LK_PGrps.TabStop = True
        Me.LK_PGrps.Text = "Update PGrps"
        '
        'LK_Users
        '
        Me.LK_Users.AutoSize = True
        Me.LK_Users.Location = New System.Drawing.Point(93, 127)
        Me.LK_Users.Name = "LK_Users"
        Me.LK_Users.Size = New System.Drawing.Size(72, 13)
        Me.LK_Users.TabIndex = 44
        Me.LK_Users.TabStop = True
        Me.LK_Users.Text = "Update Users"
        '
        'LK_Regions
        '
        Me.LK_Regions.AutoSize = True
        Me.LK_Regions.Location = New System.Drawing.Point(93, 19)
        Me.LK_Regions.Name = "LK_Regions"
        Me.LK_Regions.Size = New System.Drawing.Size(84, 13)
        Me.LK_Regions.TabIndex = 43
        Me.LK_Regions.TabStop = True
        Me.LK_Regions.Text = "Update Regions"
        '
        'LK_POrgs
        '
        Me.LK_POrgs.AutoSize = True
        Me.LK_POrgs.Location = New System.Drawing.Point(93, 96)
        Me.LK_POrgs.Name = "LK_POrgs"
        Me.LK_POrgs.Size = New System.Drawing.Size(74, 13)
        Me.LK_POrgs.TabIndex = 42
        Me.LK_POrgs.TabStop = True
        Me.LK_POrgs.Text = "Update POrgs"
        '
        'LK_SAP_Boxes
        '
        Me.LK_SAP_Boxes.AutoSize = True
        Me.LK_SAP_Boxes.Location = New System.Drawing.Point(93, 71)
        Me.LK_SAP_Boxes.Name = "LK_SAP_Boxes"
        Me.LK_SAP_Boxes.Size = New System.Drawing.Size(98, 13)
        Me.LK_SAP_Boxes.TabIndex = 41
        Me.LK_SAP_Boxes.TabStop = True
        Me.LK_SAP_Boxes.Text = "Update SAP Boxes"
        '
        'LK_Plants
        '
        Me.LK_Plants.AutoSize = True
        Me.LK_Plants.Location = New System.Drawing.Point(93, 45)
        Me.LK_Plants.Name = "LK_Plants"
        Me.LK_Plants.Size = New System.Drawing.Size(74, 13)
        Me.LK_Plants.TabIndex = 5
        Me.LK_Plants.TabStop = True
        Me.LK_Plants.Text = "Update Plants"
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.InitialImage = CType(resources.GetObject("PictureBox5.InitialImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(21, 19)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(66, 52)
        Me.PictureBox5.TabIndex = 40
        Me.PictureBox5.TabStop = False
        '
        'Panel6
        '
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel6.Location = New System.Drawing.Point(3, 496)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(685, 68)
        Me.Panel6.TabIndex = 33
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GB_Variants)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.GB_Maintenance)
        Me.Panel1.Controls.Add(Me.GB_Reports)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Splitter1)
        Me.Panel1.Controls.Add(Me.StatusStrip1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(984, 588)
        Me.Panel1.TabIndex = 3
        '
        'GB_Variants
        '
        Me.GB_Variants.BackColor = System.Drawing.Color.White
        Me.GB_Variants.Controls.Add(Me.LK_LogDirect)
        Me.GB_Variants.Controls.Add(Me.LK_LogImpNA)
        Me.GB_Variants.Controls.Add(Me.LK_LogImpLA)
        Me.GB_Variants.Controls.Add(Me.LK_LogTMS)
        Me.GB_Variants.Controls.Add(Me.LK_LogIndNAPD)
        Me.GB_Variants.Controls.Add(Me.LK_Americas_Custom)
        Me.GB_Variants.Controls.Add(Me.Label3)
        Me.GB_Variants.Controls.Add(Me.LK_Americas_Direct)
        Me.GB_Variants.Controls.Add(Me.LK_Americas_SS)
        Me.GB_Variants.Controls.Add(Me.LK_Americas_STR)
        Me.GB_Variants.Controls.Add(Me.PictureBox2)
        Me.GB_Variants.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Variants.ForeColor = System.Drawing.Color.MidnightBlue
        Me.GB_Variants.Location = New System.Drawing.Point(21, 85)
        Me.GB_Variants.Name = "GB_Variants"
        Me.GB_Variants.Size = New System.Drawing.Size(261, 339)
        Me.GB_Variants.TabIndex = 34
        Me.GB_Variants.TabStop = False
        Me.GB_Variants.Text = "PSS Variants - Americas"
        '
        'LK_LogDirect
        '
        Me.LK_LogDirect.AutoSize = True
        Me.LK_LogDirect.Location = New System.Drawing.Point(100, 259)
        Me.LK_LogDirect.Name = "LK_LogDirect"
        Me.LK_LogDirect.Size = New System.Drawing.Size(79, 13)
        Me.LK_LogDirect.TabIndex = 13
        Me.LK_LogDirect.TabStop = True
        Me.LK_LogDirect.Text = "Logistics Direct"
        '
        'LK_LogImpNA
        '
        Me.LK_LogImpNA.AutoSize = True
        Me.LK_LogImpNA.Location = New System.Drawing.Point(100, 231)
        Me.LK_LogImpNA.Name = "LK_LogImpNA"
        Me.LK_LogImpNA.Size = New System.Drawing.Size(103, 13)
        Me.LK_LogImpNA.TabIndex = 12
        Me.LK_LogImpNA.TabStop = True
        Me.LK_LogImpNA.Text = "Logistics Imports NA"
        '
        'LK_LogImpLA
        '
        Me.LK_LogImpLA.AutoSize = True
        Me.LK_LogImpLA.Location = New System.Drawing.Point(100, 200)
        Me.LK_LogImpLA.Name = "LK_LogImpLA"
        Me.LK_LogImpLA.Size = New System.Drawing.Size(101, 13)
        Me.LK_LogImpLA.TabIndex = 11
        Me.LK_LogImpLA.TabStop = True
        Me.LK_LogImpLA.Text = "Logistics Imports LA"
        '
        'LK_LogTMS
        '
        Me.LK_LogTMS.AutoSize = True
        Me.LK_LogTMS.Location = New System.Drawing.Point(100, 175)
        Me.LK_LogTMS.Name = "LK_LogTMS"
        Me.LK_LogTMS.Size = New System.Drawing.Size(105, 13)
        Me.LK_LogTMS.TabIndex = 10
        Me.LK_LogTMS.TabStop = True
        Me.LK_LogTMS.Text = "Logistics TMS NALA"
        '
        'LK_LogIndNAPD
        '
        Me.LK_LogIndNAPD.AutoSize = True
        Me.LK_LogIndNAPD.Location = New System.Drawing.Point(100, 149)
        Me.LK_LogIndNAPD.Name = "LK_LogIndNAPD"
        Me.LK_LogIndNAPD.Size = New System.Drawing.Size(119, 13)
        Me.LK_LogIndNAPD.TabIndex = 9
        Me.LK_LogIndNAPD.TabStop = True
        Me.LK_LogIndNAPD.Text = "Logistics Indirect NAPD"
        '
        'LK_Americas_Custom
        '
        Me.LK_Americas_Custom.AutoSize = True
        Me.LK_Americas_Custom.Location = New System.Drawing.Point(100, 123)
        Me.LK_Americas_Custom.Name = "LK_Americas_Custom"
        Me.LK_Americas_Custom.Size = New System.Drawing.Size(72, 13)
        Me.LK_Americas_Custom.TabIndex = 8
        Me.LK_Americas_Custom.TabStop = True
        Me.LK_Americas_Custom.Text = "Customization"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(100, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Americas"
        '
        'LK_Americas_Direct
        '
        Me.LK_Americas_Direct.AutoSize = True
        Me.LK_Americas_Direct.Location = New System.Drawing.Point(100, 95)
        Me.LK_Americas_Direct.Name = "LK_Americas_Direct"
        Me.LK_Americas_Direct.Size = New System.Drawing.Size(35, 13)
        Me.LK_Americas_Direct.TabIndex = 6
        Me.LK_Americas_Direct.TabStop = True
        Me.LK_Americas_Direct.Text = "Direct"
        '
        'LK_Americas_SS
        '
        Me.LK_Americas_SS.AutoSize = True
        Me.LK_Americas_SS.Location = New System.Drawing.Point(100, 70)
        Me.LK_Americas_SS.Name = "LK_Americas_SS"
        Me.LK_Americas_SS.Size = New System.Drawing.Size(61, 13)
        Me.LK_Americas_SS.TabIndex = 5
        Me.LK_Americas_SS.TabStop = True
        Me.LK_Americas_SS.Text = "SelfService"
        '
        'LK_Americas_STR
        '
        Me.LK_Americas_STR.AutoSize = True
        Me.LK_Americas_STR.Location = New System.Drawing.Point(100, 44)
        Me.LK_Americas_STR.Name = "LK_Americas_STR"
        Me.LK_Americas_STR.Size = New System.Drawing.Size(55, 13)
        Me.LK_Americas_STR.TabIndex = 4
        Me.LK_Americas_STR.TabStop = True
        Me.LK_Americas_STR.Text = "Storeroom"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(21, 30)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(72, 62)
        Me.PictureBox2.TabIndex = 3
        Me.PictureBox2.TabStop = False
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(984, 612)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip
        Me.MinimumSize = New System.Drawing.Size(1000, 650)
        Me.Name = "Main"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "************"
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.GB_Reports.ResumeLayout(False)
        Me.GB_Reports.PerformLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_Maintenance.ResumeLayout(False)
        Me.GB_Maintenance.PerformLayout()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.GB_Variants.ResumeLayout(False)
        Me.GB_Variants.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MainMenuToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents lb_User As System.Windows.Forms.Label
    Friend WithEvents lb_TNumber As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Main_ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Main_ToolStripProgressBar As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents GB_Reports As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents GB_Maintenance As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GB_Variants As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents LK_Americas_STR As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_Americas_SS As System.Windows.Forms.LinkLabel
    Friend WithEvents PSSVariantsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AmericasSTRToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AmericasSSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubVariantsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_Plants As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_SAP_Boxes As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_Regions As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_POrgs As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_Users As System.Windows.Forms.LinkLabel
    Friend WithEvents MaintenanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateRegionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdatePlantsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateSAPBoxesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdatePOrgsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateUsersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdatePGrpsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_PGrps As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LK_Americas_Direct As System.Windows.Forms.LinkLabel
    Friend WithEvents AmericasDirectToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubVariantsToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_Americas_Custom As System.Windows.Forms.LinkLabel
    Friend WithEvents AmericasCustomizationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubVarianToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AmericasLogisticsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IndirectNAPDToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_LogIndNAPD As System.Windows.Forms.LinkLabel
    Friend WithEvents TMSNALAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_LogTMS As System.Windows.Forms.LinkLabel
    Friend WithEvents LK_LogImpLA As System.Windows.Forms.LinkLabel
    Friend WithEvents ImportsLAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_LogImpNA As System.Windows.Forms.LinkLabel
    Friend WithEvents ImportsNAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_LogDirect As System.Windows.Forms.LinkLabel
    Friend WithEvents DirectToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubVariantsToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LK_ExportVariants As System.Windows.Forms.LinkLabel



End Class
