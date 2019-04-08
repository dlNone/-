using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace calculation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox2.Text == "")
            {
                textBox1.Text = "";
                MessageBox.Show("未输入生成计算题数量！");
            }                
            else
            {
                textBox1.Text = "";
                int num, j;
                num = Convert.ToInt32(textBox2.Text);
                bool flag = false;
                generate gen = new generate();
                if (checkBox1.CheckState == CheckState.Checked)
                    flag = true;
                for (int i = 0; i < num; i++)
                {
                    j = i + 1;
                    textBox1.Text = textBox1.Text + j.ToString() + ": " + gen.fun(flag) + Environment.NewLine;
                    System.Threading.Thread.Sleep(15);
                }
            }
                
            
                
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {          
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)          
                e.Handled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                MessageBox.Show("无计算题，无法导出！");
            else
            {
                FileStream fs = new FileStream(".\\text.txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(textBox1.Text);
                sw.Flush();
                sw.Close();
                fs.Close();
                MessageBox.Show("计算题导出成功！");
            }
        }
    }
}
