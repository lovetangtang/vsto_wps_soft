using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WindowsFormsApp1.Models;
using WindowsFormsApp1.Service;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Word;
using WordAddIn1;
using CCWin;

namespace WindowsFormsApp1
{
    public partial class Form1 : Skin_Mac
    {
        public Word.Application m_app;
        List<PeopleModel> people = new List<PeopleModel>();
        public Form1()
        {
            m_app = Globals.ThisAddIn.Application;
            InitializeComponent();
            LoadPeopleList();
        }
        private void WireUpPeopleList()
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = people;
            this.dataGridView1.Columns[0].Visible = false;
        }
        private void LoadPeopleList()
        {
            people = PeopleService.LoadPeople();
            WireUpPeopleList();
            InitDataGridColumnHeader();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void InitDataGridColumnHeader()
        {
            this.dataGridView1.Columns[1].Width = 80;
            this.dataGridView1.Columns[2].Width = 80;
            this.dataGridView1.Columns[3].Width = 80;
            this.dataGridView1.Columns[4].Width = 80;
            this.dataGridView1.Columns[5].Width = 80;

            this.dataGridView1.Columns[1].HeaderText = "角色名";
            this.dataGridView1.Columns[2].HeaderText = "年龄";
            this.dataGridView1.Columns[3].HeaderText = "性别";
            this.dataGridView1.Columns[4].HeaderText = "CV";
            this.dataGridView1.Columns[5].HeaderText = "颜色";
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            txtID.Text = row.Cells[0].Value.ToString();
            txtName.Text = row.Cells[1].Value.ToString();
            txtAge.Text = row.Cells[2].Value.ToString();
            txtCV.Text = row.Cells[4].Value.ToString();
            txtColor.Text = row.Cells[5].Value.ToString();
            if (row.Cells[3].Value.ToString().Equals("男"))
            {
                btnBoy.Checked = true;
            }
            else
            {
                btnGril.Checked = true;
            }
            btn_Add.Text = "修改";
        }

        private void Save(bool flag = true)
        {
            if (txtAge.Text == "")
            {
                MessageBox.Show("请输入年龄");
                return;
            }
            PeopleModel p = new PeopleModel();

            p.MName = txtName.Text;
            p.Age = Convert.ToInt32(txtAge.Text);
            p.Sex = btnBoy.Checked ? "男" : "女";
            p.CV = txtCV.Text;
            p.Color = txtColor.Text;
            var list = PeopleService.GetByID(p.MName);
            if (txtID.Text.Equals("添加时无编号") && list.Count == 0)
            {
                PeopleService.SavePerson(p);
            }
            else
            {
                if (txtID.Text == "添加时无编号")
                    txtID.Text = list[0].ToString();
                PeopleService.UpdatePerson(p, Convert.ToInt32(txtID.Text));
            }
            if (flag)
            {
                txtName.Text = "";
                txtAge.Text = "";
                txtCV.Text = "";
                txtID.Text = "添加时无编号";
                txtColor.Text = "";
            }
            LoadPeopleList();
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void btn_Delete_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.dataGridView1.SelectedRows.Count > 0)
                {
                    DialogResult result = MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel);
                    if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                    int x = int.Parse(dataGridView1.SelectedRows[0].Cells[0].Value.ToString());
                    PeopleService.Delete(x);
                    LoadPeopleList();
                }
                else
                {
                    MessageBox.Show("请选择需要删除的行", "提示", MessageBoxButtons.OK);
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(this, x.Message, "Error", MessageBoxButtons.OK);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LoadPeopleList();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            txtID.Text = row.Cells[0].Value.ToString();
            txtName.Text = row.Cells[1].Value.ToString();
            txtAge.Text = row.Cells[2].Value.ToString();
            txtCV.Text = row.Cells[4].Value.ToString();
            txtColor.Text = row.Cells[5].Value.ToString();
            if (row.Cells[3].Value.ToString().Equals("男"))
            {
                btnBoy.Checked = true;
            }
            else
            {
                btnGril.Checked = true;
            }
            btn_Add.Text = "修改";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (txtName.Text == "")
            {
                MessageBox.Show("请选择角色名");
                return;
            }
            Save(false);
            setCheck(txtName.Text);
        }
        private void setCheck(string text, string type = "刷色")
        {
            Word.Paragraphs paras = m_app.ActiveDocument.Paragraphs;
            int count = paras.Count;
            Word.Paragraph para;
            try
            {
                var dg = m_app.Dialogs[WdWordDialog.wdDialogFormatFont];
                dg.Show();
                var color = dg.Application.Selection.Range.Font.Color;
                var colorIndex = dg.Application.Selection.Range.Font.ColorIndex;
                //m_app.ActiveDocument.Research.Query("tets", "犯得");
                //return;

                var selList = m_app.ActiveDocument.Range().FindAll(text);
                //m_app.Selection.WholeStory Find.Execute(FindText: findText, MatchCase: true);
                //var r = m_app.Selection.Range.FindAll("犯得").Select(s => new { s.Start, s.End }).ToList();
                //colorDialog1.ShowDialog();
                System.Threading.Thread.Sleep(1000);
                foreach (var item in selList)
                {
                    item.Select();

                    //break;
                    // m_app.Selection.SetRange(item.Start, item.End);
                    Word.Range rg = m_app.Selection.Range;

                    if (type == "刷色")
                        rg.Font.Color = color;
                    else
                        rg.HighlightColorIndex = colorIndex;
                }

                //Form1 form1 = new Form1();
                //form1.Show();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtName.Text = "";
            txtAge.Text = "20";
            txtCV.Text = "";
            txtID.Text = "添加时无编号";
            txtColor.Text = "";
            btn_Add.Text = "保存";
            LoadPeopleList();
        }

        private void btn_tu_Click(object sender, EventArgs e)
        {
            if (txtName.Text == "")
            {
                MessageBox.Show("请选择角色名");
                return;
            }
            Save(false);
            setCheck(txtName.Text, "标亮");
        }
    }
}
