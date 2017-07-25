using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using System.IO;

using System.Xml;
using System.Xml.Schema;

using System.Diagnostics;

namespace CustomUIEditor
{
	/// <summary>
	/// Code to insert sample Xml
	/// </summary>
	public partial class MainForm
	{
		private string[] sampleFiles;

		private void PopulateSamplesSubMenu()
		{
			Debug.Assert(this.sampleFiles != null);
			foreach (string file in this.sampleFiles)
			{
				ToolStripMenuItem item = new ToolStripMenuItem(Path.GetFileNameWithoutExtension(file));
				item.Image = ImagesResource.xml;
				item.Tag = file;
				item.Click += new EventHandler(sampleMenuItem_Click);
				sampleToolStripMenuItem.DropDownItems.Add(item);
			}
			sampleFiles = null;
		}

		private void sampleMenuItem_Click(object sender, EventArgs e)
		{
			ToolStripMenuItem sample = sender as ToolStripMenuItem;
			string sampleFileName = sample.Tag as string;

			Debug.Assert(sampleFileName.Length > 0);
			if (this.rtbCustomUI.Text != null && this.rtbCustomUI.Text.Length > 0)
			{
				if (!this.OKToTrash()) return;
			}

			if (this.package != null && !this.package.HasCustomUI)
			{
				addPart(XMLParts.RibbonX14);
				this.rtbCustomUI.Tag = XMLParts.RibbonX14;
				this.office14PartToolStripMenuItem.Enabled = false;
			}

			try
			{
				this.colorTimer.Stop();
				TextReader tr = new StreamReader(sampleFileName);
				this.rtbCustomUI.Rtf = XmlColorizer.Colorize(tr.ReadToEnd());

				if (this.package != null)
				{
					rtfContents[(int)((XMLParts)this.rtbCustomUI.Tag)] = this.rtbCustomUI.Rtf;
				}
				else
				{
					rtbCustomUI.Tag = XMLParts.RibbonX14;
				}

				tr.Close();

                //UndoRedo
                if (_commands != null)
                {
                    _commands.NewCommand("Paste", rtbCustomUI.Rtf, rtbCustomUI.Text, rtbCustomUI.SelectionStart);
                }
                else
                {
                    _commands = new UndoRedo.Control.Commands(rtbCustomUI.Rtf);
                }
			}
			catch (System.Exception ex)
			{
				Debug.Assert(false, ex.Message);
				Debug.Fail(ex.Message);
			}
		}
	}
}
