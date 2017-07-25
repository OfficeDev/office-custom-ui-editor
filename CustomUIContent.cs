using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using System.IO;

using System.Xml;
using System.Xml.Schema;

using System.Diagnostics;
using UndoRedo.Control;

namespace CustomUIEditor
{
	/// <summary>
	/// Code related to custom UI Xml
	/// </summary>
	public partial class MainForm
	{
		#region Custom UI RichTextbox
		private void customUI_TextChanged(object sender, EventArgs e)
		{
			this.colorTimer.Stop();
			this.colorTimer.Start();
		}

		private void customUI_SelectionChanged(object sender, EventArgs e)
		{
			RichTextBox customUI = sender as RichTextBox;

			Debug.Assert(customUI != null);

			int end = customUI.SelectionStart + customUI.SelectionLength;

			int line = customUI.GetLineFromCharIndex(end);
			int col = end - customUI.GetFirstCharIndexFromLine(line);

			this.statusBar.SuspendLayout();
			this.lbLine.Text = StringsResource.idsLineStatus + (line + 1);
			this.lbColumn.Text = StringsResource.idsColumnStatus + (col + 1);
			this.statusBar.ResumeLayout();

            //UndoRedo
            _commands.NewCommand(rtbCustomUI.UndoActionName, rtbCustomUI.Rtf, rtbCustomUI.Text, rtbCustomUI.SelectionStart);
		}

		private void colorTimer_Tick(object sender, EventArgs e)
		{
			this.colorTimer.Stop();

			int selectionStart = this.rtbCustomUI.SelectionStart;
			int selectionLength = this.rtbCustomUI.SelectionLength;
			bool modified = this.rtbCustomUI.Modified;

			System.Drawing.Point pVisiblePoint = this.rtbCustomUI.GetPositionFromCharIndex(selectionStart);

			this.rtbCustomUI.SuspendLayout();
			this.rtbCustomUI.Rtf = XmlColorizer.Colorize(this.rtbCustomUI.Text);

			this.rtbCustomUI.SelectionStart = selectionStart;
			this.rtbCustomUI.SelectionLength = selectionLength;
			this.rtbCustomUI.Modified = modified;

			this.rtbCustomUI.ResumeLayout();

			this.colorTimer.Stop();
		}
		#endregion

		#region Custom Images

		private void insertIcons_Click(object sender, EventArgs e)
		{
			if (tvDocument.SelectedNode == null || tvDocument.SelectedNode.Tag == null) return;

			if (tvDocument.SelectedNode.Tag is XMLParts)
			{
				ShowInsertIconsDialog((XMLParts)tvDocument.SelectedNode.Tag);
			}
		}

		private void insertIconsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			XMLParts partType = (XMLParts)((TreeNode)tvDocument.Tag).Tag;
			ShowInsertIconsDialog(partType);
		}

		private void ShowInsertIconsDialog(XMLParts partType)
		{
			Debug.Assert(this.package != null);
			Debug.Assert(!this.package.ReadOnly);
			if (this.package == null || this.package.ReadOnly) return;

			OpenFileDialog insertIconsDialog = new OpenFileDialog();
			insertIconsDialog.Tag = partType;

			Debug.Assert(insertIconsDialog != null);

			#region Initializing Insert Icons Dialog
			insertIconsDialog.Title = StringsResource.idsInsertIconsDialogTitle;
			insertIconsDialog.Multiselect = true;
			insertIconsDialog.Filter = StringsResource.idsFilterAllSupportedImages + "|" + StringsResource.idsFilterAllFiles;
			insertIconsDialog.FilterIndex = 0;
			insertIconsDialog.RestoreDirectory = false;
			#endregion

			insertIconsDialog.FileOk += new CancelEventHandler(this.insertIcons_FileOk);
			insertIconsDialog.ShowDialog(this);

			insertIconsDialog.Dispose();
		}

		private void insertIcons_FileOk(object sender, CancelEventArgs e)
		{
			XMLParts partType = (XMLParts)(sender as OpenFileDialog).Tag;
			OfficePart part = this.package.RetrieveCustomPart(partType);
			Debug.Assert(part != null);

			TreeNode partNode = null;
			foreach (TreeNode node in tvDocument.Nodes[0].Nodes)
			{
				if (node.Text == part.Name)
				{
					partNode = node;
					break;
				}
			}
			Debug.Assert(partNode != null);

			tvDocument.SuspendLayout();
			foreach (string fileName in (sender as OpenFileDialog).FileNames)
			{
				try
				{
					string id = XmlConvert.EncodeName(Path.GetFileNameWithoutExtension(fileName));
					Stream imageStream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
					System.Drawing.Image image = System.Drawing.Image.FromStream(imageStream, true, true);

					// The file is a valid image at this point.
					id = part.AddImage(fileName, id);

					Debug.Assert(id != null, "Cannot create image part.");
					if (id == null) continue;

					imageStream.Close();

					TreeNode imageNode = new TreeNode(id);
					imageNode.ImageKey = "_" + id;
					imageNode.SelectedImageKey = imageNode.ImageKey;
					imageNode.ContextMenuStrip = imageContextMenu;
					imageNode.Tag = partType;

					tvDocument.ImageList.Images.Add(imageNode.ImageKey, image);
					partNode.Nodes.Add(imageNode);

					this.package.IsDirty = true;
				}
				catch (Exception ex)
				{
					ShowError(ex.Message);
					continue;
				}
			}

			tvDocument.ResumeLayout();
		}

		private void removeImage_Click(object sender, EventArgs e)
		{
			TreeNode currentNode = this.tvDocument.Tag as TreeNode;

			Debug.Assert(currentNode != null);
			if (currentNode == null) return;

			OfficePart part = this.package.RetrieveCustomPart((XMLParts)currentNode.Tag);
			Debug.Assert(part != null);

			part.RemoveImage(currentNode.Text);
			currentNode.Remove();

			this.package.IsDirty = true;
		}

		private void editImageId_Click(object sender, EventArgs e)
		{
			TreeNode currentNode = this.tvDocument.Tag as TreeNode;

			Debug.Assert(currentNode != null);
			if (currentNode == null) return;

			currentNode.BeginEdit();
		}

		private void imageContextMenu_Opening(object sender, CancelEventArgs e)
		{
			this.changeToolStripMenuItem.Enabled = !this.package.ReadOnly;
			this.removeToolStripMenuItem.Enabled = !this.package.ReadOnly;
			e.Cancel = false;
		}
		#endregion

		#region Edit Context Menu

		private void editContextMenu_Opening(object sender, CancelEventArgs e)
		{
			ContextMenuStrip contextMenu = sender as ContextMenuStrip;
			contextMenu.SourceControl.Focus();

			if (contextMenu.Items.Count > 0) return;

			contextMenu.Width = 150;
			contextMenu.Items.AddRange(this.CreateEditMenu());
			e.Cancel = false;
		}

		private void editToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
		{
			ToolStripMenuItem editMenu = sender as ToolStripMenuItem;

			if (editMenu.DropDownItems.Count > 0) return;

			editMenu.DropDownItems.AddRange(this.CreateEditMenu());
		}

		private ToolStripItem[] CreateEditMenu()
		{
			ToolStripItem[] items = new ToolStripItem[this.EditMenuItemText.Length];

			for (int i = 0; i < items.Length; i++)
			{
				if (this.EditMenuItemText[i] == null)
				{
					items[i] = new ToolStripSeparator();
					continue;
				}

				items[i] = new ToolStripMenuItem();
				items[i].Size = new Size(150, 22);
				items[i].Text = this.EditMenuItemText[i];
				items[i].Image = this.EditMenuItemImages[i];
				(items[i] as ToolStripMenuItem).ShortcutKeys = this.EditMenuItemKeys[i];
				items[i].Tag = this.EditMenuItemTags[i];
			}

			return items;
		}

		private void edit_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
		{
			System.Reflection.MethodInfo methodInfo = e.ClickedItem.Tag as System.Reflection.MethodInfo;
			if (methodInfo == null) return;

			Control control = null;

			if (sender is ContextMenuStrip)
			{
				control = (sender as ContextMenuStrip).SourceControl;
			}

			if (control == null)
			{
				if (this.rtbCustomUI.Focused)
				{
					control = this.rtbCustomUI;
				}
				else if (this.rtbCallbacks != null && this.rtbCallbacks.Focused)
				{
					control = this.rtbCallbacks;
				}
			}

			if (control != null && control is RichTextBox)
			{
                //RedoUndo
                if (methodInfo.Name.Equals("Undo"))
                {
                    RichTextBox rtb = control as RichTextBox;
                    int index;
                    string rtf = _commands.RemoveCommand(out index, true);
                    while (rtf.Equals(rtb.Rtf))
                    {
                        rtf = _commands.RemoveCommand(out index, false);
                        if (rtf == null) break;
                    }
                    if (rtf != null) rtb.Rtf = rtf;
                    rtb.SelectionStart = index;
                }
                else if (methodInfo.Name.Equals("Redo"))
                {
                    RichTextBox rtb = control as RichTextBox;
                    int index;
                    string rtf = _commands.RedoCommand(out index);
                    if (rtf != null)
                    {
                        while (rtf.Equals(rtb.Rtf))
                        {
                            rtf = _commands.RedoCommand(out index);

                            if (rtf == null) break;
                        }
                        rtb.Rtf = rtf;
                        rtb.SelectionStart = index;
                    }
                }
                else
                {
                    methodInfo.Invoke(control, null);
                }

				return;
			}

			Debug.Assert(false, "Fail to invoke " + methodInfo.Name);
		}

		#region Edit Menu Items Information
		private Keys[] EditMenuItemKeys = new Keys[] {
			Keys.Control | Keys.Z,
			Keys.Control | Keys.Y,
			Keys.Control,
			Keys.Control | Keys.X,
			Keys.Control | Keys.C,
			Keys.Control | Keys.V,
			Keys.Control,
			Keys.Control | Keys.A
		};

		private string[] EditMenuItemText = new String[] {
			StringsResource.idsUndo,
			StringsResource.idsRedo,
			null,
			StringsResource.idsCut,
			StringsResource.idsCopy,
			StringsResource.idsPaste,
			null,
			StringsResource.idsSellectAll
		};

		private System.Reflection.MethodInfo[] EditMenuItemTags = new System.Reflection.MethodInfo[] {
			typeof(RichTextBox).GetMethod("Undo"),
			typeof(RichTextBox).GetMethod("Redo"),
			null,
			typeof(RichTextBox).GetMethod("Cut"),
			typeof(RichTextBox).GetMethod("Copy"),
			typeof(RichTextBox).GetMethod("Paste", new Type[0]),
			null,
			typeof(RichTextBox).GetMethod("SelectAll")
		};

		private Image[] EditMenuItemImages = new Image[] {
			ImagesResource.undo,
			ImagesResource.redo,
			null,
			ImagesResource.cut,
			ImagesResource.copy,
			ImagesResource.paste,
			null,
			null
		};
		#endregion
		#endregion
	}
}
