using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CommonService;
namespace DataBaseToDocument
{
    public partial class Form1 : Form
    {
        BaseService service = new BaseService();
        NpoiToDoc docservice = new NpoiToDoc();
        public static string Form1Value; // 注意，必须申明为static变量
        public Form1()
        {
            InitializeComponent();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            var servername = txtServer.Text.Trim();
            var uid = txtUser.Text.Trim();
            var pwd = txtPwd.Text.Trim();
            var constr =service.GetConnectioning(servername,uid,pwd);
            if (service.ConnectionTest(constr))
            {
                MessageBox.Show("连接数据库成功！");
                comboBox1.DataSource = service.GetDBNameList(constr);         
            }
            else
            {
                MessageBox.Show("连接数据库失败！");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnToDoc_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue==null)
            {
                MessageBox.Show("请选择数据库");
            }
            else
            {
                var db = comboBox1.SelectedValue.ToString();
                var servername = txtServer.Text.Trim();
                var uid = txtUser.Text.Trim();
                var pwd = txtPwd.Text.Trim();
                var constr = service.GetConnectioning(servername, uid, pwd,db);
                var listnew = service.GetTableDetail("UserInfo", constr);
                var list = service.GetDBTableList(constr);
                
                docservice.CreateToWord(list,constr, db);
                MessageBox.Show("生成成功");
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue == null)
            {
                MessageBox.Show("请先保证服务器连接成功");
            }
            else
            {
                var db = comboBox1.SelectedValue.ToString();
                var servername = txtServer.Text.Trim();
                var uid = txtUser.Text.Trim();
                var pwd = txtPwd.Text.Trim();
                var constr = service.GetConnectioning(servername, uid, pwd, db);
                Form1Value = constr;            
                //this.Hide();
                var fr = new FormToBak();
                fr.ShowDialog();
                this.Close();
            }

        }
    }
}
