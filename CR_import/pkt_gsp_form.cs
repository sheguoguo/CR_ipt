using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace CR_import
{
    public partial class pkt_gsp_form : Form
    {
        public pkt_gsp_form()
        {
            InitializeComponent();
        }

        private void pkt_gsp_form_Shown(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string sql_text;
            int pack_qty = 0;
            int units_qty=0;
            int total_qty=0;
            object temp_result1,temp_result2;
            label7.Text = "0";
            label8.Text = "0";
            label9.Text = "0";
            label10.Text = "0";

            
            using (OracleConnection oraconn = oraclehelper.GetOracleConnectionAndOpen)
            {
                if (oraconn.State == ConnectionState.Open)
                {
                    //整件数量查询
                    sql_text = "select sum(ch.total_qty/im.std_pack_qty) total_pack from carton_hdr ch left join pkt_hdr ph on ph.pkt_ctrl_nbr=ch.pkt_ctrl_nbr left join item_master im on im.sku_id=ch.sku_id where  ph.pkt_nbr='" + label12.Text + "' and ch.carton_creation_code<>5";
                    temp_result1=oraclehelper.ExecuteScalar(sql_text);
                    //零头数量查询
                    sql_text = "select sum(ch.total_qty) total_pack from carton_hdr ch left join pkt_hdr ph on ph.pkt_ctrl_nbr=ch.pkt_ctrl_nbr  where  ph.pkt_nbr='" + label12.Text + "' and ch.carton_creation_code=5";
                    temp_result2 = oraclehelper.ExecuteScalar(sql_text);
                    if (temp_result1.ToString() != "" || temp_result2.ToString() != "")
                    {
                        if (temp_result1.ToString() != "")
                        {
                            pack_qty = Convert.ToInt16(temp_result1);
                            total_qty = pack_qty;
                            label7.Text = pack_qty.ToString();
                            label9.Text = total_qty.ToString();
                        }
                        if (temp_result2.ToString() != "")
                        {
                            units_qty = Convert.ToInt16(temp_result2);
                            total_qty = pack_qty+units_qty;
                            label8.Text = units_qty.ToString();
                            label9.Text = total_qty.ToString();
                        }
                    }
                    else
                    { 
                        MessageBox.Show("没有查询到该出库单，请核对日期或单号！");
                        return;
                    }
                    sql_text = "select c.gsp_nbr 药监码,c.batch_nbr 批号,im.sku_desc 品规 from c_gsp_nbr_trkg c left join pkt_hdr ph on ph.pkt_ctrl_nbr=c.pkt_ctrl_nbr left join item_master im on im.sku_id=c.sku_id  where  ph.pkt_nbr='" + dateTimePicker1.Text.Substring(3, 5) + textBox1.Text + "' and c.stat_code=0";
                    //DataSet gsp_ds = oraclehelper.ExecuteDataSet(sql_text);
                    DataTable gsp_tb=oraclehelper.ExecuteDataTable(sql_text);
                    dataGridView1.DataSource = gsp_tb;
                    dataGridView1.Refresh();
                    if (dataGridView1.RowCount != 0)
                    {
                        label10.Text = dataGridView1.RowCount.ToString();
                    }
                }
                oraconn.Close();
                button2.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount != 0 && label12.Text.Length == 11)
            {
                string improt_file = @"templates.xlsx";
                string output_file = @"output\"+"130201"+label12.Text+".xlsx";

                XSSFWorkbook wb = null;
                FileStream infs = File.Open(improt_file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                wb = new XSSFWorkbook(infs);
                infs.Close();
                infs.Dispose();
                ISheet sheet = wb.GetSheet("sheet1");
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    sheet.CreateRow(i + 1).CreateCell(0).SetCellValue(dataGridView1.Rows[i].Cells[0].Value.ToString());
                }
                FileStream outfs = File.Open(output_file, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                wb.Write(outfs);
                outfs.Close();
                outfs.Dispose();
                MessageBox.Show("文件: 130201" + label12.Text + ".xlsx"+" 已经成功生成。");
            }
            else
            {
                MessageBox.Show("请确保查询出药监码及单号正确！");
                return;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            label12.Text = dateTimePicker1.Text.Substring(3, 5) + textBox1.Text;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            label12.Text = dateTimePicker1.Text.Substring(3, 5) + textBox1.Text;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            { button1.PerformClick(); }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string ExcelName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelName = openFileDialog1.FileName;
                List<DataTable> ld = new List<DataTable>();
                ld = ExcelHepler.GetDataTablesFrom(ExcelName);
                dataGridView2.DataSource = ld[0];
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
           
                
        }

        private void dataGridView2_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView2.Columns[0].Name == "合并单号")
            {

                label12.Text = dataGridView2.SelectedRows[0].Cells["合并单号"].Value.ToString().Substring(6);
                textBox1.Text = dataGridView2.SelectedRows[0].Cells["合并单号"].Value.ToString().Substring(11, 6);
                label14.Text = "整件应扫：" + (Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["应拣数"].Value) / Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["件比"].Value)).ToString() + "  零头应扫：" + (Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["应拣数"].Value) % Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["件比"].Value)).ToString();
                label16.Text = dataGridView2.SelectedRows[0].Cells["批号"].Value.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            string sql_text;
            int pack_qty = 0;
            int units_qty = 0;
            int total_qty = 0;
            object temp_result1, temp_result2;
            label7.Text = "0";
            label8.Text = "0";
            label9.Text = "0";
            label10.Text = "0";


            using (OracleConnection oraconn = oraclehelper.GetOracleConnectionAndOpen)
            {
                if (oraconn.State == ConnectionState.Open)
                {
                    //整件数量查询
                    sql_text = "select sum(cd.units_pakd/im.std_pack_qty) total_pack from carton_dtl cd left join carton_hdr ch on ch.carton_nbr=cd.carton_nbr left join pkt_hdr ph on ph.pkt_ctrl_nbr=ch.pkt_ctrl_nbr left join item_master im on im.sku_id=ch.sku_id where  ph.pkt_nbr='" + label12.Text + "' and ch.carton_creation_code<>5 and cd.batch_nbr='"+label16.Text+"'";
                    temp_result1 = oraclehelper.ExecuteScalar(sql_text);
                    //零头数量查询
                    sql_text = "select sum(cd.units_pakd) total_pack from carton_dtl cd left join carton_hdr ch on ch.carton_nbr=cd.carton_nbr left join pkt_hdr ph on ph.pkt_ctrl_nbr=ch.pkt_ctrl_nbr  where  ph.pkt_nbr='" + label12.Text + "' and ch.carton_creation_code=5 and cd.batch_nbr='" + label16.Text + "'";
                    temp_result2 = oraclehelper.ExecuteScalar(sql_text);
                    if (temp_result1.ToString() != "" || temp_result2.ToString() != "")
                    {
                        if (temp_result1.ToString() != "")
                        {
                            pack_qty = Convert.ToInt16(temp_result1);
                            total_qty = pack_qty;
                            label7.Text = pack_qty.ToString();
                            label9.Text = total_qty.ToString();
                        }
                        if (temp_result2.ToString() != "")
                        {
                            units_qty = Convert.ToInt16(temp_result2);
                            total_qty = pack_qty + units_qty;
                            label8.Text = units_qty.ToString();
                            label9.Text = total_qty.ToString();
                        }
                    }
                    else
                    {
                        MessageBox.Show("没有查询到该出库单，请核对日期或单号！");
                        return;
                    }
                    sql_text = "select c.gsp_nbr 药监码,c.batch_nbr 批号,im.sku_desc 品规 from c_gsp_nbr_trkg c left join pkt_hdr ph on ph.pkt_ctrl_nbr=c.pkt_ctrl_nbr left join item_master im on im.sku_id=c.sku_id  where  ph.pkt_nbr='" + dateTimePicker1.Text.Substring(3, 5) + textBox1.Text + "' and c.batch_nbr='" + label16.Text + "' and c.stat_code=0";
                    //DataSet gsp_ds = oraclehelper.ExecuteDataSet(sql_text);
                    DataTable gsp_tb = oraclehelper.ExecuteDataTable(sql_text);
                    dataGridView1.DataSource = gsp_tb;
                    dataGridView1.Refresh();
                    if (dataGridView1.RowCount != 0)
                    {
                        label10.Text = dataGridView1.RowCount.ToString();
                    }
                }
                oraconn.Close();
            }
        }

        private void pkt_gsp_form_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void pkt_gsp_form_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
