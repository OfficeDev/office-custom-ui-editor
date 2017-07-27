using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using System.IO;

using System.Text;

using System.Xml;
using System.Xml.Schema;

using System.Diagnostics;
using UndoRedo.Control;

namespace CustomUIEditor
{
	/// <summary>
	/// Custom UI Editor
	/// </summary>
	public partial class MainForm : Form
	{
		private System.Windows.Forms.Timer colorTimer;
		private ContextMenuStrip editContextMenu;
		private ToolStripContainer toolStripContainer1;
		private StatusStrip statusBar;
		private MenuStrip menuMain;
		private ToolStripMenuItem fileToolStripMenuItem;
		private ToolStripMenuItem openToolStripMenuItem;
		private ToolStripMenuItem saveToolStripMenuItem;
		private ToolStripMenuItem saveAsToolStripMenuItem;
		private ToolStripMenuItem exitToolStripMenuItem;
		private ToolStripMenuItem editToolStripMenuItem;
		private ToolStripContainer toolStripContainer2;
		private ToolStrip commandBar;
		private ToolStripButton openToolStripButton;
		private ToolStripButton saveToolStripButton;
		private ToolStripSeparator cmbToolStripSeparator;
		private ToolStripContainer tsContainer;
		private ToolStripStatusLabel lbLine;
		private ToolStripStatusLabel lbColumn;
		private ToolStripStatusLabel lbDocumentName;
		private ToolStripMenuItem closeToolStripMenuItem;
		private ToolStripButton insertIconsToolStripButton;
		private ToolStripButton tsbValidateXml;
		private ToolStripButton tsbGenerateCallbacks;
		private ContextMenuStrip imageContextMenu;
		private ToolStripMenuItem changeToolStripMenuItem;
		private ToolStripMenuItem removeToolStripMenuItem;
		private ImageList tabImageList;
		private SplitContainer splMainContainer;
		private SplitContainer splcAuxContainer;
		private TreeView tvDocument;
		private RichTextBox rtbCustomUI;
		private ImageList tvImageList;
		private ContextMenuStrip partContextMenu;
		private ToolStripMenuItem insertIconsToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator1;
		private ToolStripMenuItem removePartToolStripMenuItem;
		private ToolStripMenuItem insertToolStripMenuItem;
		private ToolStripMenuItem office12PartToolStripMenuItem;
		private ToolStripMenuItem office14PartToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator2;
		private ToolStripMenuItem insertImagesToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator3;
		private ToolStripMenuItem sampleToolStripMenuItem;
		private ContextMenuStrip documentTreeContextMenuStrip;
		private ToolStripMenuItem o14ToolStripMenuItem;
		private ToolStripMenuItem o12ToolStripMenuItem;
		private System.ComponentModel.IContainer components;
		private ToolStripMenuItem helpToolStripMenuItem;
		private ToolStripMenuItem customizeTheRibbonToolStripMenuItem;
		private ToolStripMenuItem customizeTheOustpaceToolStripMenuItem;
		private ToolStripMenuItem commandIdentifiersToolStripMenuItem;
		private ToolStripMenuItem repurposingControlsToolStripMenuItem;
		private ToolStripMenuItem officeDevCenterToolStripMenuItem;
		private ToolStripMenuItem createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem;
		private ToolStripMenuItem aboutToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator4;
		private Commands _commands;

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

            //UndoRedo
            _commands = new Commands();

            //Initializing Edit menu options
            editToolStripMenuItem.DropDownItems.AddRange(this.CreateEditMenu());
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Windows.Forms.ToolStripSeparator fileSepartor1;
			System.Windows.Forms.ToolStripSeparator fileSepartor2;
			System.Windows.Forms.ToolStripSeparator fileSepartor3;
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
			this.colorTimer = new System.Windows.Forms.Timer(this.components);
			this.editContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
			this.statusBar = new System.Windows.Forms.StatusStrip();
			this.lbDocumentName = new System.Windows.Forms.ToolStripStatusLabel();
			this.lbLine = new System.Windows.Forms.ToolStripStatusLabel();
			this.lbColumn = new System.Windows.Forms.ToolStripStatusLabel();
			this.menuMain = new System.Windows.Forms.MenuStrip();
			this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.insertToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.office14PartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.office12PartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.insertImagesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.sampleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.customizeTheRibbonToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.customizeTheOustpaceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.repurposingControlsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.commandIdentifiersToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.officeDevCenterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripContainer2 = new System.Windows.Forms.ToolStripContainer();
			this.commandBar = new System.Windows.Forms.ToolStrip();
			this.openToolStripButton = new System.Windows.Forms.ToolStripButton();
			this.saveToolStripButton = new System.Windows.Forms.ToolStripButton();
			this.cmbToolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
			this.insertIconsToolStripButton = new System.Windows.Forms.ToolStripButton();
			this.tsbValidateXml = new System.Windows.Forms.ToolStripButton();
			this.tsbGenerateCallbacks = new System.Windows.Forms.ToolStripButton();
			this.tsContainer = new System.Windows.Forms.ToolStripContainer();
			this.splMainContainer = new System.Windows.Forms.SplitContainer();
			this.splcAuxContainer = new System.Windows.Forms.SplitContainer();
			this.tvDocument = new System.Windows.Forms.TreeView();
			this.documentTreeContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.o14ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.o12ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.rtbCustomUI = new System.Windows.Forms.RichTextBox();
			this.imageContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.changeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.removeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.tabImageList = new System.Windows.Forms.ImageList(this.components);
			this.tvImageList = new System.Windows.Forms.ImageList(this.components);
			this.partContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.insertIconsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.removePartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			fileSepartor1 = new System.Windows.Forms.ToolStripSeparator();
			fileSepartor2 = new System.Windows.Forms.ToolStripSeparator();
			fileSepartor3 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripContainer1.SuspendLayout();
			this.statusBar.SuspendLayout();
			this.menuMain.SuspendLayout();
			this.toolStripContainer2.SuspendLayout();
			this.commandBar.SuspendLayout();
			this.tsContainer.BottomToolStripPanel.SuspendLayout();
			this.tsContainer.ContentPanel.SuspendLayout();
			this.tsContainer.TopToolStripPanel.SuspendLayout();
			this.tsContainer.SuspendLayout();
			this.splMainContainer.Panel1.SuspendLayout();
			this.splMainContainer.Panel2.SuspendLayout();
			this.splMainContainer.SuspendLayout();
			this.splcAuxContainer.Panel1.SuspendLayout();
			this.splcAuxContainer.SuspendLayout();
			this.documentTreeContextMenuStrip.SuspendLayout();
			this.imageContextMenu.SuspendLayout();
			this.partContextMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// fileSepartor1
			// 
			fileSepartor1.Name = "fileSepartor1";
			fileSepartor1.Size = new System.Drawing.Size(145, 6);
			// 
			// fileSepartor2
			// 
			fileSepartor2.Name = "fileSepartor2";
			fileSepartor2.Size = new System.Drawing.Size(145, 6);
			// 
			// fileSepartor3
			// 
			fileSepartor3.Name = "fileSepartor3";
			fileSepartor3.Size = new System.Drawing.Size(145, 6);
			// 
			// colorTimer
			// 
			this.colorTimer.Interval = 300;
			this.colorTimer.Tick += new System.EventHandler(this.colorTimer_Tick);
			// 
			// editContextMenu
			// 
			this.editContextMenu.Name = "ctxMenu";
			this.editContextMenu.Size = new System.Drawing.Size(61, 4);
			this.editContextMenu.Opening += new System.ComponentModel.CancelEventHandler(this.editContextMenu_Opening);
			this.editContextMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.edit_ItemClicked);
			// 
			// toolStripContainer1
			// 
			// 
			// toolStripContainer1.ContentPanel
			// 
			this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(792, 520);
			this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
			this.toolStripContainer1.Name = "toolStripContainer1";
			this.toolStripContainer1.Size = new System.Drawing.Size(792, 545);
			this.toolStripContainer1.TabIndex = 8;
			this.toolStripContainer1.Text = "toolStripContainer1";
			// 
			// statusBar
			// 
			this.statusBar.Dock = System.Windows.Forms.DockStyle.None;
			this.statusBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lbDocumentName,
            this.lbLine,
            this.lbColumn});
			this.statusBar.Location = new System.Drawing.Point(0, 0);
			this.statusBar.Name = "statusBar";
			this.statusBar.Size = new System.Drawing.Size(792, 22);
			this.statusBar.TabIndex = 0;
			this.statusBar.Text = "Status Bar";
			// 
			// lbDocumentName
			// 
			this.lbDocumentName.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.lbDocumentName.Name = "lbDocumentName";
			this.lbDocumentName.Size = new System.Drawing.Size(777, 17);
			this.lbDocumentName.Spring = true;
			this.lbDocumentName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lbLine
			// 
			this.lbLine.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.lbLine.Name = "lbLine";
			this.lbLine.Size = new System.Drawing.Size(0, 17);
			this.lbLine.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lbColumn
			// 
			this.lbColumn.Name = "lbColumn";
			this.lbColumn.Size = new System.Drawing.Size(0, 17);
			this.lbColumn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// menuMain
			// 
			this.menuMain.Dock = System.Windows.Forms.DockStyle.None;
			this.menuMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.editToolStripMenuItem,
            this.insertToolStripMenuItem,
            this.helpToolStripMenuItem});
			this.menuMain.Location = new System.Drawing.Point(0, 0);
			this.menuMain.Name = "menuMain";
			this.menuMain.Size = new System.Drawing.Size(792, 24);
			this.menuMain.TabIndex = 0;
			this.menuMain.Text = "Main Menu";
			// 
			// fileToolStripMenuItem
			// 
			this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            fileSepartor1,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            fileSepartor2,
            this.closeToolStripMenuItem,
            fileSepartor3,
            this.exitToolStripMenuItem});
			this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
			this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
			this.fileToolStripMenuItem.Text = "&File";
			this.fileToolStripMenuItem.DropDownOpening += new System.EventHandler(this.fileToolStripMenuItem_DropDownOpening);
			// 
			// openToolStripMenuItem
			// 
			this.openToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.open;
			this.openToolStripMenuItem.Name = "openToolStripMenuItem";
			this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
			this.openToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
			this.openToolStripMenuItem.Text = "&Open";
			this.openToolStripMenuItem.Click += new System.EventHandler(this.open_Click);
			// 
			// saveToolStripMenuItem
			// 
			this.saveToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.save;
			this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
			this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
			this.saveToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
			this.saveToolStripMenuItem.Text = "&Save";
			this.saveToolStripMenuItem.Click += new System.EventHandler(this.save_Click);
			// 
			// saveAsToolStripMenuItem
			// 
			this.saveAsToolStripMenuItem.Enabled = false;
			this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
			this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
			this.saveAsToolStripMenuItem.Text = "Save &As";
			// 
			// closeToolStripMenuItem
			// 
			this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
			this.closeToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.W)));
			this.closeToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
			this.closeToolStripMenuItem.Text = "&Close";
			this.closeToolStripMenuItem.Click += new System.EventHandler(this.close_Click);
			// 
			// exitToolStripMenuItem
			// 
			this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
			this.exitToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
			this.exitToolStripMenuItem.Text = "E&xit";
			this.exitToolStripMenuItem.Click += new System.EventHandler(this.exit_Click);
			// 
			// editToolStripMenuItem
			// 
			this.editToolStripMenuItem.Name = "editToolStripMenuItem";
			this.editToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
			this.editToolStripMenuItem.Text = "&Edit";
			this.editToolStripMenuItem.DropDownOpening += new System.EventHandler(this.editToolStripMenuItem_DropDownOpening);
			this.editToolStripMenuItem.DropDownItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.edit_ItemClicked);
			// 
			// insertToolStripMenuItem
			// 
			this.insertToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.office14PartToolStripMenuItem,
            this.office12PartToolStripMenuItem,
            this.toolStripSeparator2,
            this.insertImagesToolStripMenuItem,
            this.toolStripSeparator3,
            this.sampleToolStripMenuItem});
			this.insertToolStripMenuItem.Name = "insertToolStripMenuItem";
			this.insertToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
			this.insertToolStripMenuItem.Text = "&Insert";
			this.insertToolStripMenuItem.DropDownOpened += new System.EventHandler(this.insertContextMenu_Opened);
			// 
			// office14PartToolStripMenuItem
			// 
			this.office14PartToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.xml;
			this.office14PartToolStripMenuItem.Name = "office14PartToolStripMenuItem";
			this.office14PartToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.office14PartToolStripMenuItem.Tag = CustomUIEditor.XMLParts.RibbonX14;
			this.office14PartToolStripMenuItem.Text = "Office 2010 Custom UI Part";
			this.office14PartToolStripMenuItem.Click += new System.EventHandler(this.addPartToolStripMenuItem_Click);
			// 
			// office12PartToolStripMenuItem
			// 
			this.office12PartToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.xml;
			this.office12PartToolStripMenuItem.Name = "office12PartToolStripMenuItem";
			this.office12PartToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.office12PartToolStripMenuItem.Tag = CustomUIEditor.XMLParts.RibbonX12;
			this.office12PartToolStripMenuItem.Text = "Office 2007 Custom UI Part";
			this.office12PartToolStripMenuItem.Click += new System.EventHandler(this.addPartToolStripMenuItem_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(213, 6);
			// 
			// insertImagesToolStripMenuItem
			// 
			this.insertImagesToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.insertPicture;
			this.insertImagesToolStripMenuItem.Name = "insertImagesToolStripMenuItem";
			this.insertImagesToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.insertImagesToolStripMenuItem.Text = "I&cons...";
			this.insertImagesToolStripMenuItem.Click += new System.EventHandler(this.insertIcons_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(213, 6);
			// 
			// sampleToolStripMenuItem
			// 
			this.sampleToolStripMenuItem.Name = "sampleToolStripMenuItem";
			this.sampleToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.sampleToolStripMenuItem.Text = "&Sample XML";
			// 
			// helpToolStripMenuItem
			// 
			this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.customizeTheRibbonToolStripMenuItem,
            this.customizeTheOustpaceToolStripMenuItem,
            this.repurposingControlsToolStripMenuItem,
            this.commandIdentifiersToolStripMenuItem,
            this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem,
            this.officeDevCenterToolStripMenuItem,
            this.toolStripSeparator4,
            this.aboutToolStripMenuItem});
			this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
			this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
			this.helpToolStripMenuItem.Text = "Help";
			// 
			// customizeTheRibbonToolStripMenuItem
			// 
			this.customizeTheRibbonToolStripMenuItem.Name = "customizeTheRibbonToolStripMenuItem";
			this.customizeTheRibbonToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.customizeTheRibbonToolStripMenuItem.Text = "Customizing the Ribbon";
			this.customizeTheRibbonToolStripMenuItem.Click += new System.EventHandler(this.customizeTheRibbonToolStripMenuItem_Click);
			// 
			// customizeTheOustpaceToolStripMenuItem
			// 
			this.customizeTheOustpaceToolStripMenuItem.Name = "customizeTheOustpaceToolStripMenuItem";
			this.customizeTheOustpaceToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.customizeTheOustpaceToolStripMenuItem.Text = "Customizing the Backstage";
			this.customizeTheOustpaceToolStripMenuItem.Click += new System.EventHandler(this.customizeTheOustpaceToolStripMenuItem_Click);
			// 
			// repurposingControlsToolStripMenuItem
			// 
			this.repurposingControlsToolStripMenuItem.Name = "repurposingControlsToolStripMenuItem";
			this.repurposingControlsToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.repurposingControlsToolStripMenuItem.Text = "Repurposing built-in commands";
			this.repurposingControlsToolStripMenuItem.Click += new System.EventHandler(this.repurposingControlsToolStripMenuItem_Click);
			// 
			// commandIdentifiersToolStripMenuItem
			// 
			this.commandIdentifiersToolStripMenuItem.Name = "commandIdentifiersToolStripMenuItem";
			this.commandIdentifiersToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.commandIdentifiersToolStripMenuItem.Text = "Office Fluent UI Command Identifiers";
			this.commandIdentifiersToolStripMenuItem.Click += new System.EventHandler(this.commandIdentifiersToolStripMenuItem_Click);
			// 
			// createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem
			// 
			this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem.Name = "createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem";
			this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem.Text = "Creating Office add-ins using Visual Studio (VSTO)";
			this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem.Click += new System.EventHandler(this.createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem_Click);
			// 
			// officeDevCenterToolStripMenuItem
			// 
			this.officeDevCenterToolStripMenuItem.Name = "officeDevCenterToolStripMenuItem";
			this.officeDevCenterToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.officeDevCenterToolStripMenuItem.Text = "Office Dev Center";
			this.officeDevCenterToolStripMenuItem.Click += new System.EventHandler(this.officeDevCenterToolStripMenuItem_Click);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(336, 6);
			// 
			// aboutToolStripMenuItem
			// 
			this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
			this.aboutToolStripMenuItem.Size = new System.Drawing.Size(339, 22);
			this.aboutToolStripMenuItem.Text = "About";
			this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
			// 
			// toolStripContainer2
			// 
			// 
			// toolStripContainer2.ContentPanel
			// 
			this.toolStripContainer2.ContentPanel.Size = new System.Drawing.Size(792, 520);
			this.toolStripContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.toolStripContainer2.Location = new System.Drawing.Point(0, 0);
			this.toolStripContainer2.Name = "toolStripContainer2";
			this.toolStripContainer2.Size = new System.Drawing.Size(792, 545);
			this.toolStripContainer2.TabIndex = 9;
			this.toolStripContainer2.Text = "toolStripContainer2";
			// 
			// commandBar
			// 
			this.commandBar.Dock = System.Windows.Forms.DockStyle.None;
			this.commandBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripButton,
            this.saveToolStripButton,
            this.cmbToolStripSeparator,
            this.insertIconsToolStripButton,
            this.tsbValidateXml,
            this.tsbGenerateCallbacks});
			this.commandBar.Location = new System.Drawing.Point(3, 24);
			this.commandBar.Name = "commandBar";
			this.commandBar.Size = new System.Drawing.Size(407, 25);
			this.commandBar.TabIndex = 1;
			this.commandBar.Text = "Command Bar";
			// 
			// openToolStripButton
			// 
			this.openToolStripButton.Image = global::CustomUIEditor.ImagesResource.open;
			this.openToolStripButton.Name = "openToolStripButton";
			this.openToolStripButton.Size = new System.Drawing.Size(56, 22);
			this.openToolStripButton.Text = "&Open";
			this.openToolStripButton.Click += new System.EventHandler(this.open_Click);
			// 
			// saveToolStripButton
			// 
			this.saveToolStripButton.Image = global::CustomUIEditor.ImagesResource.save;
			this.saveToolStripButton.Name = "saveToolStripButton";
			this.saveToolStripButton.Size = new System.Drawing.Size(51, 22);
			this.saveToolStripButton.Text = "&Save";
			this.saveToolStripButton.Click += new System.EventHandler(this.save_Click);
			// 
			// cmbToolStripSeparator
			// 
			this.cmbToolStripSeparator.Name = "cmbToolStripSeparator";
			this.cmbToolStripSeparator.Size = new System.Drawing.Size(6, 25);
			// 
			// insertIconsToolStripButton
			// 
			this.insertIconsToolStripButton.Image = global::CustomUIEditor.ImagesResource.insertPicture;
			this.insertIconsToolStripButton.Name = "insertIconsToolStripButton";
			this.insertIconsToolStripButton.Size = new System.Drawing.Size(87, 22);
			this.insertIconsToolStripButton.Text = "Insert Icons";
			this.insertIconsToolStripButton.Click += new System.EventHandler(this.insertIcons_Click);
			// 
			// tsbValidateXml
			// 
			this.tsbValidateXml.Image = global::CustomUIEditor.ImagesResource.check;
			this.tsbValidateXml.Name = "tsbValidateXml";
			this.tsbValidateXml.Size = new System.Drawing.Size(68, 22);
			this.tsbValidateXml.Text = "Validate";
			this.tsbValidateXml.Click += new System.EventHandler(this.tsbValidateXml_Click);
			// 
			// tsbGenerateCallbacks
			// 
			this.tsbGenerateCallbacks.Image = global::CustomUIEditor.ImagesResource.callbacks;
			this.tsbGenerateCallbacks.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.tsbGenerateCallbacks.Name = "tsbGenerateCallbacks";
			this.tsbGenerateCallbacks.Size = new System.Drawing.Size(127, 22);
			this.tsbGenerateCallbacks.Text = "Generate Callbacks";
			this.tsbGenerateCallbacks.Click += new System.EventHandler(this.tsbGenerateCallbacks_Click);
			// 
			// tsContainer
			// 
			// 
			// tsContainer.BottomToolStripPanel
			// 
			this.tsContainer.BottomToolStripPanel.Controls.Add(this.statusBar);
			// 
			// tsContainer.ContentPanel
			// 
			this.tsContainer.ContentPanel.Controls.Add(this.splMainContainer);
			this.tsContainer.ContentPanel.Size = new System.Drawing.Size(792, 474);
			this.tsContainer.Dock = System.Windows.Forms.DockStyle.Fill;
			// 
			// tsContainer.LeftToolStripPanel
			// 
			this.tsContainer.LeftToolStripPanel.Enabled = false;
			this.tsContainer.LeftToolStripPanelVisible = false;
			this.tsContainer.Location = new System.Drawing.Point(0, 0);
			this.tsContainer.Name = "tsContainer";
			this.tsContainer.RightToolStripPanelVisible = false;
			this.tsContainer.Size = new System.Drawing.Size(792, 545);
			this.tsContainer.TabIndex = 10;
			this.tsContainer.Text = "toolStripContainer3";
			// 
			// tsContainer.TopToolStripPanel
			// 
			this.tsContainer.TopToolStripPanel.Controls.Add(this.menuMain);
			this.tsContainer.TopToolStripPanel.Controls.Add(this.commandBar);
			// 
			// splMainContainer
			// 
			this.splMainContainer.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splMainContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
			this.splMainContainer.Location = new System.Drawing.Point(0, 0);
			this.splMainContainer.Name = "splMainContainer";
			// 
			// splMainContainer.Panel1
			// 
			this.splMainContainer.Panel1.Controls.Add(this.splcAuxContainer);
			// 
			// splMainContainer.Panel2
			// 
			this.splMainContainer.Panel2.Controls.Add(this.rtbCustomUI);
			this.splMainContainer.Size = new System.Drawing.Size(792, 474);
			this.splMainContainer.SplitterDistance = 160;
			this.splMainContainer.TabIndex = 0;
			// 
			// splcAuxContainer
			// 
			this.splcAuxContainer.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splcAuxContainer.Location = new System.Drawing.Point(0, 0);
			this.splcAuxContainer.Name = "splcAuxContainer";
			this.splcAuxContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
			// 
			// splcAuxContainer.Panel1
			// 
			this.splcAuxContainer.Panel1.Controls.Add(this.tvDocument);
			this.splcAuxContainer.Panel2Collapsed = true;
			this.splcAuxContainer.Size = new System.Drawing.Size(160, 474);
			this.splcAuxContainer.SplitterDistance = 253;
			this.splcAuxContainer.TabIndex = 0;
			// 
			// tvDocument
			// 
			this.tvDocument.ContextMenuStrip = this.documentTreeContextMenuStrip;
			this.tvDocument.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tvDocument.LabelEdit = true;
			this.tvDocument.Location = new System.Drawing.Point(0, 0);
			this.tvDocument.Name = "tvDocument";
			this.tvDocument.Size = new System.Drawing.Size(160, 474);
			this.tvDocument.TabIndex = 0;
			this.tvDocument.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.tvDocument_AfterLabelEdit);
			this.tvDocument.BeforeSelect += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvDocument_BeforeSelect);
			this.tvDocument.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvDocument_AfterSelect);
			this.tvDocument.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvDocument_NodeMouseClick);
			// 
			// documentTreeContextMenuStrip
			// 
			this.documentTreeContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.o14ToolStripMenuItem,
            this.o12ToolStripMenuItem});
			this.documentTreeContextMenuStrip.Name = "documentTreeContextMenuStrip";
			this.documentTreeContextMenuStrip.Size = new System.Drawing.Size(217, 48);
			this.documentTreeContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(this.insertContextMenu_Opened);
			// 
			// o14ToolStripMenuItem
			// 
			this.o14ToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.xml;
			this.o14ToolStripMenuItem.Name = "o14ToolStripMenuItem";
			this.o14ToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.o14ToolStripMenuItem.Tag = CustomUIEditor.XMLParts.RibbonX14;
			this.o14ToolStripMenuItem.Text = "Office 2010 Custom UI Part";
			this.o14ToolStripMenuItem.Click += new System.EventHandler(this.addPartToolStripMenuItem_Click);
			// 
			// o12ToolStripMenuItem
			// 
			this.o12ToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.xml;
			this.o12ToolStripMenuItem.Name = "o12ToolStripMenuItem";
			this.o12ToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
			this.o12ToolStripMenuItem.Tag = CustomUIEditor.XMLParts.RibbonX12;
			this.o12ToolStripMenuItem.Text = "Office 2007 Custom UI Part";
			this.o12ToolStripMenuItem.Click += new System.EventHandler(this.addPartToolStripMenuItem_Click);
			// 
			// rtbCustomUI
			// 
			this.rtbCustomUI.AcceptsTab = true;
			this.rtbCustomUI.ContextMenuStrip = this.editContextMenu;
			this.rtbCustomUI.Dock = System.Windows.Forms.DockStyle.Fill;
			this.rtbCustomUI.Location = new System.Drawing.Point(0, 0);
			this.rtbCustomUI.Name = "rtbCustomUI";
			this.rtbCustomUI.Size = new System.Drawing.Size(628, 474);
			this.rtbCustomUI.TabIndex = 0;
			this.rtbCustomUI.Text = "";
			this.rtbCustomUI.SelectionChanged += new System.EventHandler(this.customUI_SelectionChanged);
			this.rtbCustomUI.TextChanged += new System.EventHandler(this.customUI_TextChanged);
			// 
			// imageContextMenu
			// 
			this.imageContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.changeToolStripMenuItem,
            this.removeToolStripMenuItem});
			this.imageContextMenu.Name = "imageContextMenu";
			this.imageContextMenu.Size = new System.Drawing.Size(130, 48);
			this.imageContextMenu.Opening += new System.ComponentModel.CancelEventHandler(this.imageContextMenu_Opening);
			// 
			// changeToolStripMenuItem
			// 
			this.changeToolStripMenuItem.Name = "changeToolStripMenuItem";
			this.changeToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
			this.changeToolStripMenuItem.Text = "Change ID";
			this.changeToolStripMenuItem.Click += new System.EventHandler(this.editImageId_Click);
			// 
			// removeToolStripMenuItem
			// 
			this.removeToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.remove;
			this.removeToolStripMenuItem.Name = "removeToolStripMenuItem";
			this.removeToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
			this.removeToolStripMenuItem.Text = "Remove";
			this.removeToolStripMenuItem.Click += new System.EventHandler(this.removeImage_Click);
			// 
			// tabImageList
			// 
			this.tabImageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("tabImageList.ImageStream")));
			this.tabImageList.TransparentColor = System.Drawing.Color.Magenta;
			this.tabImageList.Images.SetKeyName(0, "xml.png");
			this.tabImageList.Images.SetKeyName(1, "callbacks.bmp");
			// 
			// tvImageList
			// 
			this.tvImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit;
			this.tvImageList.ImageSize = new System.Drawing.Size(16, 16);
			this.tvImageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// partContextMenu
			// 
			this.partContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.insertIconsToolStripMenuItem,
            this.toolStripSeparator1,
            this.removePartToolStripMenuItem});
			this.partContextMenu.Name = "partContextMenu";
			this.partContextMenu.Size = new System.Drawing.Size(144, 54);
			// 
			// insertIconsToolStripMenuItem
			// 
			this.insertIconsToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.insertPicture;
			this.insertIconsToolStripMenuItem.Name = "insertIconsToolStripMenuItem";
			this.insertIconsToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
			this.insertIconsToolStripMenuItem.Text = "Insert Icons...";
			this.insertIconsToolStripMenuItem.Click += new System.EventHandler(this.insertIconsToolStripMenuItem_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(140, 6);
			// 
			// removePartToolStripMenuItem
			// 
			this.removePartToolStripMenuItem.Image = global::CustomUIEditor.ImagesResource.remove;
			this.removePartToolStripMenuItem.Name = "removePartToolStripMenuItem";
			this.removePartToolStripMenuItem.Size = new System.Drawing.Size(143, 22);
			this.removePartToolStripMenuItem.Text = "Remove";
			this.removePartToolStripMenuItem.Click += new System.EventHandler(this.removePartToolStripMenuItem_Click);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
			this.AutoSize = true;
			this.ClientSize = new System.Drawing.Size(792, 545);
			this.Controls.Add(this.tsContainer);
			this.Controls.Add(this.toolStripContainer2);
			this.Controls.Add(this.toolStripContainer1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MainMenuStrip = this.menuMain;
			this.MinimumSize = new System.Drawing.Size(320, 240);
			this.Name = "MainForm";
			this.Text = "Microsoft Office Custom UI Editor";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.toolStripContainer1.ResumeLayout(false);
			this.toolStripContainer1.PerformLayout();
			this.statusBar.ResumeLayout(false);
			this.statusBar.PerformLayout();
			this.menuMain.ResumeLayout(false);
			this.menuMain.PerformLayout();
			this.toolStripContainer2.ResumeLayout(false);
			this.toolStripContainer2.PerformLayout();
			this.commandBar.ResumeLayout(false);
			this.commandBar.PerformLayout();
			this.tsContainer.BottomToolStripPanel.ResumeLayout(false);
			this.tsContainer.BottomToolStripPanel.PerformLayout();
			this.tsContainer.ContentPanel.ResumeLayout(false);
			this.tsContainer.TopToolStripPanel.ResumeLayout(false);
			this.tsContainer.TopToolStripPanel.PerformLayout();
			this.tsContainer.ResumeLayout(false);
			this.tsContainer.PerformLayout();
			this.splMainContainer.Panel1.ResumeLayout(false);
			this.splMainContainer.Panel2.ResumeLayout(false);
			this.splMainContainer.ResumeLayout(false);
			this.splcAuxContainer.Panel1.ResumeLayout(false);
			this.splcAuxContainer.ResumeLayout(false);
			this.documentTreeContextMenuStrip.ResumeLayout(false);
			this.imageContextMenu.ResumeLayout(false);
			this.partContextMenu.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

		private void MainForm_Load(object sender, System.EventArgs e)
		{
			string applicationFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
			this.LoadXmlSchemas(applicationFolder + @"\Schemas\");
			this.LoadSamples(applicationFolder + @"\Samples\");

			this.splMainContainer.Panel1Collapsed = true;

			this.rtbCustomUI.Rtf = XmlColorizer.rtfString;
			this.rtbCustomUI.Modified = false;
			this.colorTimer.Stop();

			this.LoadTreeViewImages();
			this.ChangePackageButtonState(false);
		}

		private void exit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		protected override void OnClosing(CancelEventArgs e)
		{
			base.OnClosing(e);
			e.Cancel = !OKToTrash();
		}

		#region Validates Xml

		private void tsbValidateXml_Click(object sender, EventArgs e)
		{
			this.ValidateXml(true /*showValidMessage*/);
		}

		private bool ValidateXml(bool showValidMessage)
		{
			if (this.rtbCustomUI.Text == null || this.rtbCustomUI.Text.Length == 0)
			{
				return false;
			}

			this.rtbCustomUI.SuspendLayout();

			// Test to see if text is XML first
			try
			{
				XmlDocument xmlDoc = new XmlDocument();
				XMLParts partType = (XMLParts)this.rtbCustomUI.Tag;

				XmlSchema targetSchema = this.customUISchemas[this.rtbCustomUI.Tag] as XmlSchema;

				if (targetSchema == null)
				{
					return false;
				}

				xmlDoc.Schemas.Add(targetSchema);

				xmlDoc.LoadXml(this.rtbCustomUI.Text);

				if (xmlDoc.DocumentElement.NamespaceURI.ToString() != targetSchema.TargetNamespace)
				{
					StringBuilder errorText = new StringBuilder();
					errorText.Append(StringsResource.idsUnknownNamespace.Replace("|1", xmlDoc.DocumentElement.NamespaceURI.ToString()));
					errorText.Append("\n" + StringsResource.idsCustomUINamespace.Replace("|1", targetSchema.TargetNamespace));

					ShowError(errorText.ToString());
					return false;
				}

				this.hasXmlError = false;
				xmlDoc.Validate(this.XmlValidationEventHandler);
			}
			catch (XmlException ex)
			{
				ShowError(StringsResource.idsInvalidXml + "\n" + ex.Message);
				return false;
			}

			this.rtbCustomUI.ResumeLayout();

			if (!this.hasXmlError)
			{
				if (showValidMessage)
				{
					MessageBox.Show(
						this,
						StringsResource.idsValidXml,
						this.Text,
						MessageBoxButtons.OK,
						MessageBoxIcon.Information);
				}
				return true;
			}
			return false;
		}

		private void XmlValidationEventHandler(object sender, ValidationEventArgs e)
		{
			lock (this)
			{
				this.hasXmlError = true;
			}
			MessageBox.Show(
				this,
				e.Message,
				e.Severity.ToString(),
				MessageBoxButtons.OK,
				(e.Severity == XmlSeverityType.Error ? MessageBoxIcon.Error : MessageBoxIcon.Warning));
		}

		private void LoadXmlSchemas(string folderName)
		{
			this.tsbValidateXml.Visible = false;
			this.tsbGenerateCallbacks.Visible = false;

			Debug.Assert(!(folderName == null || folderName.Length == 0));
			if (folderName == null || folderName.Length == 0) return;

			try
			{
				string[] schemas = Directory.GetFiles(folderName, "CustomUI*.xsd");
				if (schemas == null || schemas.Length == 0)
				{
					return;
				}
				this.customUISchemas = new Hashtable(schemas.Length);

				for (int i = 0; i < schemas.Length; i++)
				{
					XMLParts partType = schemas[i].Contains("14") ? XMLParts.RibbonX14 : XMLParts.RibbonX12;
					StreamReader reader = new StreamReader(schemas[i]);
					this.customUISchemas.Add(partType, XmlSchema.Read(reader, null));

					reader.Close();
				}

				this.tsbValidateXml.Visible = true;
				this.tsbGenerateCallbacks.Visible = true;
			}
			catch (System.Exception ex)
			{
				Debug.Fail(ex.Message);
			}
		}

		#endregion

		#region Generates Callbacks
		private void tsbGenerateCallbacks_Click(object sender, EventArgs e)
		{
			if (this.rtbCustomUI.Text == null || this.rtbCustomUI.Text.Length == 0)
			{
				return;
			}

			if (!this.ValidateXml(false /* showValidMessage */))
			{
				return;
			}

			try
			{
				System.Xml.XmlDocument customUI = new XmlDocument();
				customUI.LoadXml(this.rtbCustomUI.Text);

				System.Text.StringBuilder callbacks = CallbacksBuilder.GenerateCallback(customUI);
				if (callbacks == null || callbacks.Length == 0)
				{
					MessageBox.Show(StringsResource.idsNoCallback, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				this.package.IsDirty = this.rtbCustomUI.Modified;

				ShowCallbacks(callbacks.ToString());
			}
			catch (Exception ex)
			{
				Debug.Assert(false, ex.Message);
			}
		}

		private void ShowCallbacks(string callbacks)
		{
			this.splMainContainer.Panel2.SuspendLayout();
			if (this.rtbCallbacks == null)
			{
				this.rtbCallbacks = new RichTextBox();
				this.rtbCallbacks.Dock = DockStyle.Fill;
				this.rtbCallbacks.ContextMenuStrip = this.rtbCustomUI.ContextMenuStrip;
				this.rtbCallbacks.ReadOnly = true;
				this.splMainContainer.Panel2.Controls.Add(this.rtbCallbacks);
			}

			this.rtbCallbacks.Rtf = callbacks;

			this.rtbCustomUI.Visible = false;
			this.rtbCallbacks.Visible = true;
			this.splMainContainer.Panel2.ResumeLayout();

			this.commandBar.Enabled =
			this.menuMain.Enabled = false;
		}

		private void HideCallbacks()
		{
			if (this.rtbCallbacks == null || !this.rtbCallbacks.Visible)
			{
				return;
			}

			splMainContainer.Panel2.SuspendLayout();
			this.rtbCustomUI.Visible = true;
			this.rtbCallbacks.Visible = false;
			splMainContainer.Panel2.ResumeLayout();

			this.commandBar.Enabled =
			this.menuMain.Enabled = true;

			this.rtbCustomUI.Modified = this.package.IsDirty;
		}

		#endregion

		#region Private Variables
		private OfficeDocument package = null;

		private Hashtable customUISchemas;

		private bool hasXmlError;

		private RichTextBox rtbCallbacks = null;
		#endregion

		private void LoadSamples(string path)
		{
			Debug.Assert(!(path == null || path.Length == 0));
			if (path == null || path.Length == 0) return;

			this.sampleToolStripMenuItem.Visible = false;

			try
			{
				this.sampleFiles = System.IO.Directory.GetFiles(path, "*.xml");
				this.sampleToolStripMenuItem.Visible = !(this.sampleFiles == null || this.sampleFiles.Length == 0);
			}
			catch (System.IO.IOException ex)
			{
				Debug.Fail(ex.Message);
			}
		}

		private void ShowError(string errorText)
		{
			Debug.Assert(!(errorText == null || errorText.Length == 0));

			MessageBox.Show(
				this,
				errorText,
				this.Text,
				MessageBoxButtons.OK,
				MessageBoxIcon.Error);
		}

		private void LoadTreeViewImages()
		{
			tvImageList.Images.Add(OfficeApplications.XML.ToString(), ImagesResource.xml);
			tvImageList.Images.Add(OfficeApplications.Word.ToString(), ImagesResource.worddoc);
			tvImageList.Images.Add(OfficeApplications.Excel.ToString(), ImagesResource.excelwkb);
			tvImageList.Images.Add(OfficeApplications.PowerPoint.ToString(), ImagesResource.pptpre);
		}

		private void insertContextMenu_Opened(object sender, EventArgs e)
		{
			if (this.sampleToolStripMenuItem.DropDownItems.Count == 0)
			{
				PopulateSamplesSubMenu();
			}

			if (this.package == null || tvDocument == null) return;

			TreeNode rootNode = tvDocument.Nodes[0] as TreeNode;
			if (rootNode == null) return;

			this.office14PartToolStripMenuItem.Enabled = true;
			this.office14PartToolStripMenuItem.Enabled = true;
			this.insertImagesToolStripMenuItem.Enabled = false;

			foreach (TreeNode node in rootNode.Nodes)
			{
				XMLParts partType = (XMLParts)node.Tag;
				if (partType == XMLParts.RibbonX14)
				{
					this.office14PartToolStripMenuItem.Enabled = false;
				}
				else if (partType == XMLParts.RibbonX12)
				{
					this.office12PartToolStripMenuItem.Enabled = false;
				}
			}

			TreeNode selectedNode = tvDocument.SelectedNode as TreeNode;
			if (selectedNode == null || selectedNode.Tag == null) return;
			XMLParts selectedPartType = (XMLParts) selectedNode.Tag;

			this.insertImagesToolStripMenuItem.Enabled = ((selectedPartType == XMLParts.RibbonX12) || (selectedPartType == XMLParts.RibbonX14));
		}

		void addPartToolStripMenuItem_Click(object sender, EventArgs e)
		{
			XMLParts partType = (XMLParts)((ToolStripMenuItem)sender).Tag;
			addPart(partType);
		}

		private void addPart(XMLParts partType)
		{
			OfficePart newPart = this.package.CreateCustomPart(partType);

			TreeNode partNode = ConstructPartNode(newPart);

			TreeNode currentNode = tvDocument.Nodes[0] as TreeNode;
			Debug.Assert(currentNode != null);
			if (currentNode == null) return;

			tvDocument.SuspendLayout();
			currentNode.Nodes.Add(partNode);
			rtfContents[(int)partType] = string.Empty;
			tvDocument.SelectedNode = partNode;
			tvDocument.ResumeLayout();
		}

		private void removePartToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TreeNode currentNode = tvDocument.Tag as TreeNode;
			Debug.Assert(currentNode != null);

			XMLParts partType = (XMLParts)currentNode.Tag;
			if (partType == XMLParts.RibbonX12)
			{
				this.office12PartToolStripMenuItem.Enabled = true;
			}
			else if (partType == XMLParts.RibbonX14)
			{
				this.office14PartToolStripMenuItem.Enabled = true;
			}

			tvDocument.SuspendLayout();
			this.package.RemoveCustomPart(partType);
			rtfContents[(int)partType] = String.Empty;

			if (currentNode.Text == tvDocument.SelectedNode.Text)
			{
				tvDocument.SelectedNode.Tag = null;
				tvDocument.SelectedNode = tvDocument.Nodes[0];
				rtbCustomUI.Rtf = String.Empty;
				rtbCustomUI.Tag = XMLParts.RibbonX14;

				for (int i = 0; i < rtfContents.Length; i++)
				{
					if (rtfContents[i] == null || rtfContents[i] == string.Empty)
					{
						continue;
					}
					rtbCustomUI.Rtf = rtfContents[i];
					rtbCustomUI.Tag = (XMLParts)i;
					break;
				}
			}

			currentNode.Remove();
			tvDocument.ResumeLayout();
		}

		private void insertContextMenu_Opened(object sender, CancelEventArgs e)
		{
			if (this.package == null || tvDocument == null) return;

			TreeNode rootNode = tvDocument.Nodes[0] as TreeNode;
			if (rootNode == null) return;

			this.o12ToolStripMenuItem.Enabled = true;
			this.o14ToolStripMenuItem.Enabled = true;

			foreach (TreeNode node in rootNode.Nodes)
			{
				XMLParts partType = (XMLParts)node.Tag;
				if (partType == XMLParts.RibbonX14)
				{
					this.o14ToolStripMenuItem.Enabled = false;
				}
				else if (partType == XMLParts.RibbonX12)
				{
					this.o12ToolStripMenuItem.Enabled = false;
				}
			}
		}

		private void customizeTheRibbonToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://msdn.microsoft.com/en-us/library/aa338202(v=office.14).aspx");
		}

		private void customizeTheOustpaceToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://msdn.microsoft.com/en-us/library/ee691833(office.14).aspx");
		}

		private void commandIdentifiersToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://github.com/OfficeDev/office-fluent-ui-command-identifiers");
		}

		private void createCOMVSTOAddinsForOfficeByUsingVisualStudioToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://msdn.microsoft.com/en-us/library/jj620922.aspx");
		}

		private void officeDevCenterToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://dev.office.com/");
		}

		private void repurposingControlsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Process.Start("https://blogs.technet.microsoft.com/the_microsoft_excel_support_team_blog/2012/06/18/how-to-repurpose-a-button-in-excel-2007-or-2010/");
		}

		private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Form about = new AboutBox();
			about.Show();
		}
	}
}
