using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointWSRemote
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void txtPort_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.port = txtPort.Text;
            Properties.Settings.Default.Save();
            PowerPointWSAddIn.instance.Setup();
        }

        private void cbEnabled_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.enabled = cbEnabled.Checked;
            Properties.Settings.Default.Save();
            PowerPointWSAddIn.instance.Setup();
        }
    }
}
