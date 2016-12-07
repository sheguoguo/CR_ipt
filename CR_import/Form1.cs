using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.DataAccess.Client;

namespace CR_import
{
    public partial class pkt_import_Form : Form
    {
        string ExcelName = "";

        public pkt_import_Form()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelName = openFileDialog1.FileName;
                List<DataTable> ld=new List<DataTable>();
                ld=ExcelHepler.GetDataTablesFrom(ExcelName);
                dataGridView1.DataSource = ld[0];
                dataGridView2.Rows.Clear();
                button3.Enabled = true;
                button2.Enabled = false;
                

            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount != 0)
            {
                try
                {
                    using (OracleConnection oraconn = oraclehelper.GetOracleConnectionAndOpen)
                    {
                        if (oraconn.State == ConnectionState.Open)
                        {
                            for (int rc = 0; rc <= dataGridView2.RowCount - 1; rc++)
                            {
                                int pkt_dtl_inst_result=0;
                                int pkt_hdr_inst_result = 0;

                                if (dataGridView2.Rows[rc].Cells["细单序号"].Value.ToString() == "1")
                                {
                                    string pkt_hdr_inst = "insert into inpt_pkt_hdr(pkt_ctrl_nbr,whse,co,div,pkt_nbr,pkt_sfx,shipto,shipto_name,shipto_addr_1,shipto_city,shipto_cntry,soldto_state,acct_rcvbl_acct_nbr,nbr_of_times_appt_chgd,ship_via,pkt_genrtn_date_time,terms_pcnt,swc_nbr_seq,ord_nbr_for_swc,nbr_ords_for_swc,carton_label_type,nbr_of_label,contnt_label_type,nbr_of_contnt_label,nbr_of_pakng_slips,auto_invc_stat_code,cust_rte,est_vol_bridged,total_nbr_of_units,total_dlrs_undisc,total_dlrs_disc,nbr_of_hngr,bol_type,prod_value,carton_asn_reqd,max_carton_wt,max_carton_units,nbr_of_zones,wave_stat_code,stat_code,wave_seq_nbr,est_carton_bridged,partl_carton_min_wt,est_wt_bridged,carton_cubng_indic,partl_carton_optn,trlr_stop_seq_nbr,major_pkt_grp_attr,whse_xfer_flag,error_seq_nbr,proc_stat_code,rte_stop_seq,ftsr_nbr,duty_and_tax,rte_load_seq,plt_cubng_indic,est_wt,est_carton,est_vol,host_inpt_id,est_carton_per_pallet_bridged,est_pallet_bridged,pre_pack_flag,parent_ord_id,monetary_value,incoterm_id,tms_ord_type,create_date_time,mod_date_time,duty_tax_payment_type,hub_id) values('";
                                    pkt_hdr_inst += dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "','S00','SPL','SPL','" + dataGridView2.Rows[rc].Cells["合并单号"].Value.ToString().Substring(6, 11) + "','S','B100077434','" + dataGridView2.Rows[rc].Cells["客户名称"].Value.ToString() + "','" + dataGridView2.Rows[rc].Cells["合并单号"].Value.ToString().Substring(11, 6) + "','0','CN',0,'7911391',0,'YS06',to_date('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','yyyy-mm-dd hh24:mi:ss'),0,'0','0','0','001',1,'001','0','0','0','N',0,0,0,0,0,'R',0,'Y',0,0,0,0,10,0,0,0,0,'51','P',0,'CR',0,0,0,0,'" + dataGridView2.Rows[rc].Cells["合并单号"].Value.ToString() + "',0,0,0,0,0,0,'" + dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "',0,0,0,0,0,'CR出库订单','自提',to_date('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','yyyy-mm-dd hh24:mi:ss'),to_date('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','yyyy-mm-dd hh24:mi:ss'),0,-1)";
                                    
                                    pkt_hdr_inst_result = oraclehelper.ExecuteNonQuery(pkt_hdr_inst);
                                    
                                }
                                string pkt_dtl_inst="insert into inpt_pkt_dtl(pkt_ctrl_nbr,pkt_seq_nbr,season,style,color,SIZE_DESC,orig_ord_qty,orig_pkt_qty,back_ord_qty,cancel_qty,cube_mult_qty,units_pikd,units_pakd,std_bundl_qty,std_pack_qty,unit_wt,unit_vol,ppack_grp_code,carton_break_attr,critcl_dim_1,critcl_dim_2,critcl_dim_3,std_case_wt,std_case_vol,batch_nbr,price,retail_price,stat_code,pick_rate,pack_rate,ovrsz_len,orig_ord_line_nbr,orig_pkt_line_nbr,srl_nbr_reqd_flag,spl_instr_code_1,spl_instr_code_2,to_be_verf_as_pakd,ppack_qty,std_sub_pack_qty,std_case_qty,verf_as_pakd,wave_proc_type,to_be_pikd,std_plt_qty,shelf_days,partl_fill,proc_stat_code,error_seq_nbr,host_inpt_id,po_size_value,actual_cost,budget_cost,carton_epc_type,plt_epc_type,pkg_type_instance,monetary_value,unit_monetary_value,prod_sched_ref_nbr,cust_po_line_nbr,create_date_time,mod_date_time) ";
                                pkt_dtl_inst += "values('" + dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "'," + dataGridView2.Rows[rc].Cells["细单序号"].Value.ToString() + ",'CR','" + dataGridView2.Rows[rc].Cells["商品编码"].Value.ToString() + "','" + dataGridView2.Rows[rc].Cells["件比"].Value.ToString() + "','" + dataGridView2.Rows[rc].Cells["商品编码"].Value.ToString() + "',0," + dataGridView2.Rows[rc].Cells["应拣数"].Value.ToString() + ",0,0,0,0,0,0,0,0,0,'S','1',0,0,0,0,0,'" + dataGridView2.Rows[rc].Cells["批号"].Value.ToString() + "',0,0,0,0,0,0,0,0,'0','00','0',0,0,0,0,0,0,0,0,0,'N',0,0,'" + dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "',0,0,0,0,0,'国药准字',0,0,'001',0,to_date('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','yyyy-mm-dd hh24:mi:ss'),to_date('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','yyyy-mm-dd hh24:mi:ss'))";
                                //string pkt_dtl_inst = "insert into inpt_pkt_dtl(pkt_ctrl_nbr,pkt_seq_nbr,season,style,SIZE_DESC,orig_ord_qty,orig_pkt_qty,back_ord_qty,cancel_qty,cube_mult_qty,units_pikd,units_pakd,std_bundl_qty,std_pack_qty,unit_wt,unit_vol,ppack_grp_code,carton_break_attr,critcl_dim_1,critcl_dim_2,critcl_dim_3,std_case_wt,std_case_vol,batch_nbr,price,retail_price,stat_code,pick_rate,pack_rate,ovrsz_len,orig_ord_line_nbr,orig_pkt_line_nbr,srl_nbr_reqd_flag,spl_instr_code_1,spl_instr_code_2,to_be_verf_as_pakd,ppack_qty,std_sub_pack_qty,std_case_qty,verf_as_pakd,wave_proc_type,to_be_pikd,std_plt_qty,shelf_days,partl_fill,proc_stat_code,error_seq_nbr,host_inpt_id,po_size_value,actual_cost,budget_cost,carton_epc_type,plt_epc_type,pkg_type_instance,monetary_value,unit_monetary_value,prod_sched_ref_nbr,cust_po_line_nbr,create_date_time,mod_date_time) values('" + dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "',1,'CR','11921','11921',0,400,0,0,0,0,0,0,0,0,0,'S','1',0,0,0,0,0,'N99532',0,0,0,0,0,0,0,0,'0','00','0',0,0,0,0,0,0,0,0,0,'N',0,0,'" + dataGridView2.Rows[rc].Cells["拣货单号"].Value.ToString() + "',0,0,0,0,0,'国药准字',0,0,'001',0,to_date('2016-08-29 15:23:16','yyyy-mm-dd hh24:mi:ss'),to_date('2016-08-29 15:23:16','yyyy-mm-dd hh24:mi:ss'))";
                                pkt_dtl_inst_result = oraclehelper.ExecuteNonQuery(pkt_dtl_inst);
                                if (pkt_dtl_inst_result==1)
                                {
                                    dataGridView2.Rows[rc].DefaultCellStyle.BackColor=Color.LightGreen;
                                }
                            }

                        }
                        oraconn.Close();
                        button2.Enabled = false;
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("出现异常, 异常信息: " + ex.Message);
                }


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount != 0)
            {
                dataGridView2.Rows.Clear();
                string pkt_seq_sql = "select inpt_pkt_seq.nextval from dual";
                try
                {
                    using (OracleConnection oraconn = oraclehelper.GetOracleConnectionAndOpen)
                    {
                        if (oraconn.State == ConnectionState.Open)
                        {
                            for (int rc = 0; rc <= dataGridView1.RowCount - 1; rc++)
                            {
                                //处理合并单号细单序号及拣货单总数量
                                int RowIndex = dataGridView2.Rows.Add();
                                string pkt_nbr = "";
                                int pkt_dtl_seq = 1;
                                //int total_qty = 0;
                                if (RowIndex!=0)
                                {
                                    
                                    for (int rc2 = 0; rc2 < rc; rc2++)
                                    {
                                        if (dataGridView1.Rows[rc].Cells["合并单号"].Value.ToString() == dataGridView2.Rows[rc2].Cells["合并单号"].Value.ToString())
                                        {
                                            pkt_dtl_seq += 1;
                                            pkt_nbr = dataGridView2.Rows[rc2].Cells["拣货单号"].Value.ToString();
                                            //total_qty += Convert.ToInt32(dataGridView2.Rows[rc2].Cells["应拣数"].Value);
                                            
                                        }
                                    
                                    }
                                }

                                //判断是否重复导入

                                string pkt_nbr_sql = "select count(*) from pkt_hdr ph left join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr=ph.pkt_ctrl_nbr where phi.stat_code<>99 and ph.pkt_nbr='" + dataGridView1.Rows[rc].Cells["合并单号"].Value.ToString().Substring(6, 11) + "'";

                              //  string pkt_nbr_sql = "select count(*) from pkt_hdr ph where ph.pkt_nbr='" + dataGridView1.Rows[rc].Cells["合并单号"].Value.ToString().Substring(6, 11) +"'";

                                if (oraclehelper.ExecuteScalar(pkt_nbr_sql).ToString() != "0")
                                {
                                    MessageBox.Show("系统已有相同的合并单号" + dataGridView1.Rows[rc].Cells["合并单号"].Value.ToString() + "，请核对后再重新导入!");
                                    return;
                                }

                                dataGridView2.Rows[RowIndex].Cells["合并单号"].Value = dataGridView1.Rows[rc].Cells["合并单号"].Value;
                                dataGridView2.Rows[RowIndex].Cells["商品编码"].Value = dataGridView1.Rows[rc].Cells["商品编码"].Value;
                                dataGridView2.Rows[RowIndex].Cells["商品名称"].Value = dataGridView1.Rows[rc].Cells["商品名称"].Value;
                                dataGridView2.Rows[RowIndex].Cells["批号"].Value = dataGridView1.Rows[rc].Cells["批号"].Value;
                                dataGridView2.Rows[RowIndex].Cells["应拣数"].Value = dataGridView1.Rows[rc].Cells["应拣数"].Value;
                                dataGridView2.Rows[RowIndex].Cells["客户名称"].Value = dataGridView1.Rows[rc].Cells["客户名称"].Value;
                                dataGridView2.Rows[RowIndex].Cells["件比"].Value = dataGridView1.Rows[rc].Cells["件比"].Value;
                                dataGridView2.Rows[RowIndex].Cells["细单序号"].Value = pkt_dtl_seq;
                                if (pkt_dtl_seq == 1)
                                {
                                    dataGridView2.Rows[RowIndex].Cells["拣货单号"].Value = "CRS0" + oraclehelper.ExecuteScalar(pkt_seq_sql);

                                }
                                else
                                {
                                    dataGridView2.Rows[RowIndex].Cells["拣货单号"].Value = pkt_nbr;
                                }
                            }
                        
                        }
                        oraconn.Close();
                        for (int rc = dataGridView2.RowCount-1; rc >=1 ; rc--)
                        {
                            for (int rc_twice = rc - 1; rc_twice >= 0; rc_twice--)
                            {
                                if (dataGridView2.Rows[rc].Cells["合并单号"].Value.ToString() == dataGridView2.Rows[rc_twice].Cells["合并单号"].Value.ToString() && dataGridView2.Rows[rc].Cells["商品编码"].Value.ToString() == dataGridView2.Rows[rc_twice].Cells["商品编码"].Value.ToString() && dataGridView2.Rows[rc].Cells["批号"].Value.ToString() == dataGridView2.Rows[rc_twice].Cells["批号"].Value.ToString())
                                {
                                    dataGridView2.Rows[rc_twice].Cells["应拣数"].Value = Convert.ToInt32(dataGridView2.Rows[rc_twice].Cells["应拣数"].Value) + Convert.ToInt32(dataGridView2.Rows[rc].Cells["应拣数"].Value);
                                    dataGridView2.Rows.RemoveAt(rc);
                                    break;
                                }
                            }
                        }
                        button2.Enabled = true;
                        button3.Enabled = false;
                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("出现异常, 异常信息: " + ex.Message);
                } 

                
            }
        }
    }
}
