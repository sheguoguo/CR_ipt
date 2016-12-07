using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CR_import
{
    public partial class mianform : Form
    {
        pkt_import_Form import_form;
        pkt_gsp_form pkt_gsp_fm;

        public mianform()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (import_form != null)
            {
                import_form.Dispose();
                import_form = new pkt_import_Form();
                import_form.Show();
                import_form.Focus();
            }
            else
            {
                import_form = new pkt_import_Form();
                import_form.Show();
                import_form.Focus();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (pkt_gsp_fm != null)
            {
                pkt_gsp_fm.Dispose();
                pkt_gsp_fm = new pkt_gsp_form();
                pkt_gsp_fm.Show();
                pkt_gsp_fm.Focus();
            }
            else
            {
                pkt_gsp_fm = new pkt_gsp_form();
                pkt_gsp_fm.Show();
                pkt_gsp_fm.Focus();
            }
        }
    }
}
