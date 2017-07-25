using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace CustomUIEditor
{
    public partial class MainForm : Form
    {
        private void richTextBoxMethod_Click(object sender, EventArgs e)
        {
            string methodName = sender.GetType().GetProperty("Tag").GetValue(sender, null) as string;
            if (methodName == null || methodName == string.Empty)
            {
                return;
            }
            try
            {
                this.richTextBox.GetType().InvokeMember(
                    methodName,
                    System.Reflection.BindingFlags.InvokeMethod,
                    null /*binder*/,
                    this.richTextBox,
                    null /*args*/);
            }
            catch { }
        }

        private void richTextBox_TextChanged(object sender, EventArgs e)
        {
            this.validateToolStripButton.Enabled = (this.richTextBox.TextLength > 0);
            this.colorTimer.Stop();
            this.colorTimer.Start();
        }

        private void richTextBox_ModifiedChanged(object sender, EventArgs e)
        {
            if (this.package != null)
            {
                this.package.IsDirty = this.richTextBox.Modified;
                this.saveToolStripButton.Enabled = this.richTextBox.Modified;
                this.saveToolStripMenuItem.Enabled = this.richTextBox.Modified;
            }
        }

        private void richTextBox_SelectionChanged(object sender, EventArgs e)
        {
            int start = this.richTextBox.SelectionStart;
            int length = this.richTextBox.SelectionLength;

            int line = this.richTextBox.GetLineFromCharIndex(start + length);
            int col = start + length - this.richTextBox.GetFirstCharIndexFromLine(line);
            this.line.Text = "Ln " + (line + 1);
            this.column.Text = "Col " + (col + 1);

            this.UpdateSelectionControlEnableState(length > 0);
        }

        private void UpdateSelectionControlEnableState(bool enable)
        {
            this.cutContextMenuItem.Enabled = enable;
            this.deleteContextMenuItem.Enabled = enable;
            this.copyContextMenuItem.Enabled = enable;
        }

        private void deleteContextMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectedText = "";
        }
    }
}