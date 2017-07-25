using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

using System.IO;
using System.IO.Packaging;

using System.Xml;
using System.Xml.Schema;

using System.Diagnostics;

namespace CustomUIEditor
{
	/// <summary>
	/// Custom UI Editor
	/// </summary>
	public partial class MainForm
	{
		#region Package
		private string[] rtfContents = new string[(int)XMLParts.LastEntry];

		private void PackageLoaded()
		{
			this.SuspendLayout();

			this.rtbCustomUI.ReadOnly = this.package.ReadOnly;

			string readOnly = (this.rtbCustomUI.ReadOnly ? " (" + StringsResource.idsReadOnly + ")" : String.Empty);
			this.Text = Path.GetFileName(this.package.Name) + readOnly + " - " + this.Text;
			this.lbDocumentName.Text = this.package.Name + readOnly;

			if (this.package.Parts != null && this.package.Parts.Count > 0)
			{
				if (this.rtbCustomUI.Tag != null)
				{
					this.rtbCustomUI.Rtf = string.Empty;
					rtfContents[(int)this.rtbCustomUI.Tag] = string.Empty;
				}

				foreach (OfficePart part in this.package.Parts)
				{
					int contentIndex = (int)part.PartType;
					rtbCustomUI.Rtf = XmlColorizer.Colorize(part.ReadContent());
					rtbCustomUI.Tag = part.PartType;
					rtfContents[contentIndex] = rtbCustomUI.Rtf;
				}
			}
			else
			{
				this.rtbCustomUI.Tag = XMLParts.RibbonX14;
			}

			this.sampleToolStripMenuItem.Enabled = !this.package.ReadOnly;

			this.ChangePackageButtonState(true);

			this.splMainContainer.Panel1Collapsed = false;

			this.rtbCustomUI.Modified = false;

			this.ResumeLayout();
		}

		private void PackageClosed()
		{
			this.SuspendLayout();

			Debug.Assert(this.package != null);
			if (this.package != null)
			{
				this.package.Dispose();
				this.package = null;
			}

			tvDocument.Nodes.Clear();
			tvDocument.ImageList = null;

			this.rtbCustomUI.Tag = null;
			this.rtbCustomUI.Clear();
			this.rtbCustomUI.ClearUndo();
			this.rtbCustomUI.Modified = false;
			this.colorTimer.Stop();

			this.ChangePackageButtonState(false);
			this.sampleToolStripMenuItem.Enabled = true;

			this.ResumeLayout();
		}

		private void ChangePackageButtonState(bool hasDocument)
		{
			if (!hasDocument)
			{
				this.lbDocumentName.Text = StringsResource.idsOpenDocumentPrompt;
				this.Text = StringsResource.idsApplicationTitle;
			}

			this.saveToolStripButton.Enabled =
			this.saveToolStripMenuItem.Enabled = hasDocument && !this.package.ReadOnly;

			this.tsbGenerateCallbacks.Enabled =
			this.closeToolStripMenuItem.Enabled = hasDocument;

			this.splMainContainer.Panel1Collapsed = !hasDocument;

			this.insertIconsToolStripButton.Enabled = hasDocument && !this.package.ReadOnly;
			this.insertImagesToolStripMenuItem.Enabled = this.insertIconsToolStripButton.Enabled;
			this.office12PartToolStripMenuItem.Enabled = this.office14PartToolStripMenuItem.Enabled = this.insertImagesToolStripMenuItem.Enabled;
		}

		private void fileToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
		{
			this.closeToolStripMenuItem.Enabled = (this.package != null);
			this.saveToolStripMenuItem.Enabled = this.saveToolStripButton.Enabled;
		}

		private void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
			try
			{
				Debug.Assert(this.package == null);
				if (this.package != null)
				{
					this.package.Dispose();
					this.PackageClosed();
				}

				string fileName = (sender as OpenFileDialog).FileName;
				if (fileName == null || fileName.Length == 0) return;

				Debug.WriteLine("Opening " + fileName + "...");
				this.package = new OfficeDocument(fileName);
				this.PackageLoaded();
				this.PackageTreeView();

                //UndoRedo
                _commands = new UndoRedo.Control.Commands(rtbCustomUI.Rtf);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, this.Text);
			}
		}

		private void close_Click(object sender, System.EventArgs e)
		{
			Debug.Assert(this.package != null);

			if (!this.OKToTrash()) return;

			this.package.Dispose();
			PackageClosed();
		}

		private void save_Click(object sender, System.EventArgs e)
		{
			this.Save();
		}

		private bool Save()
		{
			try
			{
				Debug.Assert(this.package != null);
				if (this.package == null) return false;

				if (this.rtbCustomUI.Modified)
				{
					this.colorTimer.Stop();
					Debug.WriteLine("Saving custom UI content...");

					if (this.rtbCustomUI.Text.Length > 0)
					{
						this.package.SaveCustomPart((XMLParts)this.rtbCustomUI.Tag, this.rtbCustomUI.Text, true /*isCreatingNewPart*/);
						if (!this.package.HasCustomUI)
						{
							this.office14PartToolStripMenuItem.Enabled = false;
							addPart(XMLParts.RibbonX14);
						}
					}
					else
					{
						this.package.SaveCustomPart((XMLParts)this.rtbCustomUI.Tag, this.rtbCustomUI.Text);
					}
					this.rtbCustomUI.Modified = false;
				}

				if (this.package.IsDirty)
				{
					Debug.WriteLine("Saving package...");
					this.package.Save();
				}

				return true;
			}
			catch (Exception ex)
			{
				ShowError(ex.Message);
			}
			return false;
		}

		private void open_Click(object sender, System.EventArgs e)
		{
			if (!this.OKToTrash()) return;

			OpenFileDialog openDocumentDialog = new OpenFileDialog();

			#region Initializing Open Document Dialog
			openDocumentDialog.Title = StringsResource.idsOpenDocumentDialogTitle;
			openDocumentDialog.Filter = StringsResource.idsFilterAllOfficeDocuments + "|" +
				StringsResource.idsFilterWordDocuments + "|" +
				StringsResource.idsFilterExcelDocuments + "|" +
				StringsResource.idsFilterPPTDocuments + "|" +
				StringsResource.idsFilterAllFiles;
			openDocumentDialog.FilterIndex = 0;
			openDocumentDialog.RestoreDirectory = true;
			#endregion

			openDocumentDialog.FileOk += new CancelEventHandler(this.openFileDialog_FileOk);
			openDocumentDialog.ShowDialog(this);
			openDocumentDialog.Dispose();
		}
		#endregion

		private bool OKToTrash()
		{
			if (this.package == null || !(this.rtbCustomUI.Modified || this.package.IsDirty))
			{
				return true;
			}

			DialogResult result = MessageBox.Show(
				this,
				StringsResource.idsAskForSaveChanges,
				this.Text,
				MessageBoxButtons.YesNoCancel,
				MessageBoxIcon.Warning);

			switch (result)
			{
				case DialogResult.No:
					return true;

				case DialogResult.Cancel:
					return false;

				case DialogResult.Yes:
					return this.Save();

				default:
					return false;
			}
		}

		private void PackageTreeView()
		{
			this.tvDocument.SuspendLayout();
			this.tvDocument.ImageList = tvImageList;

			TreeNode root = new TreeNode(System.IO.Path.GetFileName(this.package.Name));

			root.ImageKey = OfficeDocument.MapFileType(System.IO.Path.GetExtension(this.package.Name)).ToString();
			root.SelectedImageKey = root.ImageKey;

			foreach (OfficePart part in this.package.Parts)
			{
				TreeNode node = ConstructPartNode(part);
				root.Nodes.Add(node);
				root.Expand();
			}

			tvDocument.Nodes.Add(root);
			tvDocument.SelectedNode = root.Nodes.Count == 0 ? root : root.Nodes[0];
			this.tvDocument.ResumeLayout();
		}

		private TreeNode ConstructPartNode(OfficePart part)
		{
			TreeNode node = new TreeNode(part.Name);
			node.Tag = part.PartType;
			node.ImageKey = OfficeApplications.XML.ToString();
			node.SelectedImageKey = node.ImageKey;
			if (!this.package.ReadOnly)
			{
				node.ContextMenuStrip = partContextMenu;
			}

			List<TreeNode> images = part.GetImages(this.tvDocument.ImageList, this.package.ReadOnly ? null : imageContextMenu);
			if (images != null && images.Count > 0)
			{
				node.Nodes.AddRange(images.ToArray());
				node.Collapse();
			}

			return node;
		}

		private void tvDocument_AfterSelect(object sender, TreeViewEventArgs e)
		{
			if (e.Node.Tag == null || e.Node.Tag is TreeNode) return;

			XMLParts partType = (XMLParts)e.Node.Tag;

			if (rtbCustomUI.Tag != null && partType == (XMLParts)rtbCustomUI.Tag) return;

			rtbCustomUI.Tag = partType;
			rtbCustomUI.Rtf = rtfContents[(int)partType];

			this.rtbCustomUI.Modified = false;
			this.colorTimer.Stop();
		}

		private void tvDocument_BeforeSelect(object sender, TreeViewCancelEventArgs e)
		{
			if (this.tvDocument.SelectedNode == null || this.tvDocument.SelectedNode.Tag == null) return;

			if (this.rtbCustomUI.Modified)
			{
				this.rtbCustomUI.Modified = false;
				this.colorTimer.Stop();

				rtfContents[(int)((XMLParts)this.tvDocument.SelectedNode.Tag)] = this.rtbCustomUI.Rtf;
				this.package.SaveCustomPart((XMLParts)this.tvDocument.SelectedNode.Tag, this.rtbCustomUI.Text);
			}
		}

		private void tvDocument_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
		{
			HideCallbacks();
			if (e.Button == MouseButtons.Right)
			{
				if (e.Node != null)
				{
					this.tvDocument.Tag = e.Node;
				}
			}
		}

		private void tvDocument_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
		{
			if (e.CancelEdit || e.Label == null || e.Label.Length == 0)
			{
				return;
			}

			if (e.Node.Text == e.Label)
			{
				return;
			}

			OfficePart part = this.package.RetrieveCustomPart((XMLParts)e.Node.Tag);
			Debug.Assert(part != null);

			if (part == null) return;

			try
			{
				part.ChangeImageId(e.Node.Text, e.Label);
				this.package.IsDirty = true;
			}
			catch (Exception ex)
			{
				ShowError(ex.Message);
				e.CancelEdit = true;
			}
		}
	}
}
