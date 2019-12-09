using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;
using Excel;
using System.Collections;
using AttendanceRecord.Helper;
using AttendanceRecord.Entities;
namespace AttendanceRecord
{
    public partial class FrmLearning : Form
    {
        public FrmLearning() {
            InitializeComponent();
        }
        public static string _action = "Learning";
        /// <summary>
        /// 依据姓名查询。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbName_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (13 != e.KeyChar) return;
            btnSubmit_Click(sender, e);
        }

          private void timerClsResult_Tick(object sender, EventArgs e)
        {
            lblResult.Text = "";
            lblResult.BackColor = Color.SkyBlue;
            timerClsResult.Stop();
        }
      
        private void FrmLearning_Load(object sender, EventArgs e)
        {
            this.dgv.DataSource = Learning.getAllLearners();
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void delTheLearnerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow dgvR = dgv.CurrentRow;
            if (dgvR == null) return;
            string name = dgvR.Cells["姓名"].Value.ToString();
            Learning learner = new Learning(name);
            learner.del();
            ShowResult.show(lblResult, "已删除！", true);
            timerClsResult.Start();
            this.dgv.DataSource = Learning.getAllLearners();
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            string name = tbName.Text.Trim();
            if (name.Length == 0) return;
            Learning learner = new Learning(name);
            if (learner.ifExists()) {
                ShowResult.show(lblResult, "已存在！", false);
                timerClsResult.Start();
                tbName.Clear();
                return;
            }
            learner.add();
            ShowResult.show(lblResult, "已添加！",true);
            timerClsResult.Start();
            refreshDgv();
            tbName.Clear();
        }
        private void refreshDgv() {
            this.dgv.DataSource = Learning.getAllLearners();
            DGVHelper.AutoSizeForDGV(dgv);
        }
    }
}
