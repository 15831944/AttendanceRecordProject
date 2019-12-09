using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AttendanceRecord.Entities;
namespace AttendanceRecord
{
    public partial class FormLogin : Form
    {
        public FormLogin()
        {
            InitializeComponent();
        }
        private void tbUserName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar !=13) {
                return;
            }
            if (String.IsNullOrEmpty(tbUserName.Text.Trim())) return;
            //检查用户名： 
            if (!CheckUsers.ifExistsUserName(tbUserName.Text.Trim())) {
                MessageBox.Show("用户名无效","提示：",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                tbUserName.Clear();
                return;
            };
            tbPwd.Focus();
        }
        /// <summary>
        /// 判断密码是否有效。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbPwd_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 13)
            {
                return;
            }
            if (String.IsNullOrEmpty(tbUserName.Text.Trim())) return;
            if (String.IsNullOrEmpty(tbPwd.Text.Trim())) return;
            //检查密码 
            if (!CheckUsers.isPasswordRight(tbUserName.Text.Trim(), tbPwd.Text.Trim()))
            {
                MessageBox.Show("密码无效", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tbPwd.Clear();
                return;
            };
            Program._userInfo = CheckUsers.getUserInfo(tbUserName.Text.Trim());
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
