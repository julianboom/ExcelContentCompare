using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
/// <summary>
/// 功能说明，在副表中找出主表缺少的数据，并在副表中标记为红色
/// </summary>
namespace ContentCompare
{
    public partial class BU : Form
    {
        public BU()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = VV();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox2.Text = VV();
        }
        public string VV()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";//过滤一下，只要表格格式的
            ofd.InitialDirectory = "c:\\";
            ofd.RestoreDirectory = true;
            ofd.FilterIndex = 1;
            ofd.AddExtension = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.ShowHelp = true;//是否显示帮助按钮
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
            this.button2.Enabled = false;
            this.button3.Enabled = false;
            this.label3.Visible = true;
            new Thread(() => {
                try
                {
                    List<string> tableNames1 = getNames(this.textBox1.Text);
                    List<string> list = new List<string>();//定义主表中姓名的容器
                    Workbook workbook = new Workbook();
                    workbook.Open(this.textBox2.Text);
                    int col = 0;
                    int row = 0;
                    Cells cells = null;
                    geCells(workbook, ref cells, ref col, ref row);
                    Style style = workbook.Styles[workbook.Styles.Add()];//新增样式
                    style.HorizontalAlignment = TextAlignmentType.Center;//文字居中
                    style.Font.Name = "宋体";//文字字体
                                           //style.Font.Size = 14;//文字大小
                                           //style.Font.IsBold = true;//粗体
                    style.IsTextWrapped = true;//单元格内容自动换行
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = Color.Red;//设置背景颜色
                    style.Font.Color = Color.Black;//文本颜色

                    for (; row < cells.MaxDataRow + 1; row++)//遍历每行数据
                    {
                        MatchCollection matchCollection = Regex.Matches(cells[row, col].StringValue, "[\u4e00-\u9fa5]+");//通过正则匹配出字符串中的姓名部分
                        if (matchCollection.Count > 0)//如果含有姓名
                            foreach (var item in matchCollection)
                            {
                                string result = tableNames1.Find(f => f.Equals(item.ToString()));
                                if (result == null)
                                {
                                    cells[row, col].SetStyle(style);
                                }
                            }
                    }
                    workbook.Save(this.textBox2.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("数据统计错误，请验证表格数据");
                    return;
                }
                finally
                {
                    this.button1.Enabled = true;
                    this.button2.Enabled = true;
                    this.button3.Enabled = true;
                    this.label3.Visible = false;
                }
                MessageBox.Show("数据统计结束，请查看表格数据");
            })
            {
                IsBackground = true,
                Priority = ThreadPriority.Highest
            }.Start();
        }

        private List<string> getNames(string path)
        {
            Workbook workbook = new Workbook();
            workbook.Open(path);
            Cells cells = workbook.Worksheets[0].Cells;

            int col = 0;//定义列起点
            int row = 0;//定义行起点
            for (; col < cells.MaxDataColumn + 1; col++)//遍历列
            {

                string s = cells[row, col].StringValue.Trim();//获取每一列
                if (s.Equals(this.textBox3.Text))
                {
                    break;//定位至目标列
                }
            }
            row++;//从第二行开始遍历
            List<string> list = new List<string>();//定义主表中姓名的容器
            for (; row < cells.MaxDataRow + 1; row++)//遍历每行数据
            {
                MatchCollection matchCollection = Regex.Matches(cells[row, col].StringValue, "[\u4e00-\u9fa5]+");//通过正则匹配出字符串中的姓名部分
                if (matchCollection.Count > 0)//如果含有姓名
                    foreach (var item in matchCollection)
                    {
                        list.Add(item.ToString());//则将该姓名添加至集合中
                    }
            }
            return list;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="cells"></param>
        /// <param name="col">//定义列，第一列开始遍历</param>
        /// <param name="row">//定义行，第二行开始遍历</param>
        private void geCells(Workbook workbook, ref Cells cells,ref int col,ref int row)
        {
            
            cells = workbook.Worksheets[0].Cells;
            for (; col < cells.MaxDataColumn + 1; col++)//遍历列
            {

                string s = cells[row, col].StringValue.Trim();//获取每一列
                if (s.Equals(this.textBox4.Text))
                {
                    break;//定位至目标列
                }
            }
            row++;
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
