using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using Spire.Doc;
using Spire.Doc.Documents;

namespace calculation
{
    
    public partial class Form1 : Form
    {
        String[] shizi = new String[10000];
        bool ischanged = true;

        public object Application { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }
        //生成按钮点击事件
        private void button1_Click(object sender, EventArgs e)
        {
            ischanged = false;
            if (textBox2.Text == "")
            {
                textBox1.Text = "";
                MessageBox.Show("未输入生成计算题数量！");
            }                
            else
            {
                textBox1.Text = "";
                int num;
                generate gen = new generate();
                num = Convert.ToInt32(textBox2.Text);
                shizi = gen.fun(num);
                //显示答案
                if (checkBox1.CheckState == CheckState.Checked)
                {
                    dataGridView1.Rows.Clear();                    
                    for (int i = 0; i < num; i++)
                    {
                        String[] str = shizi[i].Split('=');
                        int index = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index].Cells[0].Value = i + 1;
                        this.dataGridView1.Rows[index].Cells[1].Value = str[0] + '=';
                        this.dataGridView1.Rows[index].Cells[2].Value = str[1];

                    }
                }
                    
                //不显示答案        
                else
                {
                    dataGridView1.Rows.Clear();
                    for (int i = 0; i < num; i++)
                    {
                        String[] str = shizi[i].Split('=');                       
                        int index = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index].Cells[0].Value = i+1;
                        this.dataGridView1.Rows[index].Cells[1].Value = str[0] + '=';

                    }                       
                }                
                
            }
            ischanged = true;

        }
        //限制只能输入数字
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {          
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)          
                e.Handled = true;
        }
        //导出按钮点击事件
        private void button2_Click(object sender, EventArgs e)
        {
            int num = 0;
            if(textBox2.Text != "")
                num = Convert.ToInt32(textBox2.Text);
            
            //创建Word文档
            Document doc = new Document();
            //添加section
            Section section = doc.AddSection();

            //添加表格
            Table table = section.AddTable(true);
                
            //添加第1行
            TableRow row1 = table.AddRow();
            TableCell cell1 = row1.AddCell();
            cell1.AddParagraph().AppendText("题 号");
            cell1.CellFormat.BackColor = Color.Gold;
            TableCell cell2 = row1.AddCell();
            cell2.AddParagraph().AppendText("题 目");
            
            TableCell cell3 = row1.AddCell();
            cell3.AddParagraph().AppendText("题 号");
            cell3.CellFormat.BackColor = Color.Gold;
            TableCell cell4 = row1.AddCell();
            cell4.AddParagraph().AppendText("题 目");
            
            TableCell cell5 = row1.AddCell();
            cell5.AddParagraph().AppendText("题 号");
            cell5.CellFormat.BackColor = Color.Gold;
            TableCell cell6 = row1.AddCell();
            cell6.AddParagraph().AppendText("题 目");
            
            TableCell cell7 = row1.AddCell();
            cell7.AddParagraph().AppendText("题 号");
            cell7.CellFormat.BackColor = Color.Gold;
            TableCell cell8 = row1.AddCell();
            cell8.AddParagraph().AppendText("题 目");
            
            int j = 0;
            for (int i = 0; i < num / 4 + 1; i++)
            {                    
                TableRow row = table.AddRow();
                j++;
                if(j < num + 1)
                {
                    table.Rows[i + 1].Cells[0].AddParagraph().AppendText(j.ToString());
                    table.Rows[i + 1].Cells[0].CellFormat.BackColor = Color.Gold;
                    table.Rows[i + 1].Cells[1].AddParagraph().AppendText(shizi[j - 1]);                  
                }                    
                j++;
                if (j < num + 1)
                {
                    table.Rows[i + 1].Cells[2].AddParagraph().AppendText(j.ToString());
                    table.Rows[i + 1].Cells[2].CellFormat.BackColor = Color.Gold;
                    table.Rows[i + 1].Cells[3].AddParagraph().AppendText(shizi[j - 1]);                 
                }
                    
                j++;
                if (j < num + 1)
                {
                    table.Rows[i + 1].Cells[4].AddParagraph().AppendText(j.ToString());
                    table.Rows[i + 1].Cells[4].CellFormat.BackColor = Color.Gold;
                    table.Rows[i + 1].Cells[5].AddParagraph().AppendText(shizi[j - 1]);
                }
                j++;
                if (j < num + 1)
                {
                    table.Rows[i + 1].Cells[6].AddParagraph().AppendText(j.ToString());
                    table.Rows[i + 1].Cells[6].CellFormat.BackColor = Color.Gold;
                    table.Rows[i + 1].Cells[7].AddParagraph().AppendText(shizi[j - 1]);
                }
            }
            //保存文档
            table.AutoFit(AutoFitBehaviorType.FixedColumnWidths);
            doc.SaveToFile("Table.docx");
            MessageBox.Show("导出成功！");                     
        }
        //是否显示答案
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            ischanged = false;
            textBox1.Text = "";
            int num = 0;
            try
            {                
                num = Convert.ToInt32(textBox2.Text);
            }
            catch { }            
            if (checkBox1.CheckState == CheckState.Checked)
            {
                dataGridView1.Rows.Clear();
                for (int i = 0; i < num; i++)
                {
                    String[] str = shizi[i].Split('=');
                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = i + 1;
                    this.dataGridView1.Rows[index].Cells[1].Value = str[0] + '=';
                    this.dataGridView1.Rows[index].Cells[2].Value = str[1];
                }
            }                                   
            else
            {
                dataGridView1.Rows.Clear();
                for (int i = 0; i < num; i++)
                {
                    String[] str = shizi[i].Split('=');
                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = i + 1;
                    this.dataGridView1.Rows[index].Cells[1].Value = str[0] + '=';
                }
            }
            ischanged = true;
        }
        //表格中填写答案
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(ischanged)
                if(e.ColumnIndex == 2 && e.RowIndex > -1)
                {

                    string msg = String.Format("Cell at row {0}, column {1} value changed", e.RowIndex, e.ColumnIndex);
                    if (this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() != shizi[e.RowIndex].Split('=')[1])
                    {
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                        MessageBox.Show("答案错误！");
                    }
                    else
                        dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                }
        }
    }
}
