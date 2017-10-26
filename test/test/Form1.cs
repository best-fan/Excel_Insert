using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
using Office;
using System.IO;

namespace test
{
    public partial class Form1 : Form
    {
        Excel.Application etApp;
        Excel.Workbook etbook;
        Excel.Worksheet etsheet;
        Excel.Range etrange;
        //SelectName 表格名称   SelectPicPath  图片存放路径  PicRows图片起始行 PicColumns 添加图片的列  Picname 图片名称 
        //NameColumns 姓名所在列  AllRows 所有行
        public String SelectName, SelectPicPath, Picname;
        public int AllRows, PicRows, PicColumns, NameColumns;
        public int Nowrows;
        public  Dictionary<int , String> PicNa = new Dictionary<int, String>();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// excel 操作；
        /// </summary>
        public void Eto()
        {

            //获取工作表表格
            etApp = new Excel.Application();
            etbook = (Excel.Workbook)etApp.Workbooks.Open(SelectName);

            //获取数据区域
            etsheet = (Excel.Worksheet)etbook.Worksheets.get_Item(1);
            //获取单元格行数
            AllRows = etsheet.UsedRange.Rows.Count;
            //获取单元格列数
            // columns = etsheet.UsedRange.Columns.Count;
            //获取数据区域 
            etrange = (Excel.Range)etsheet.UsedRange;

            double m = etsheet.UsedRange.Height;
            //MessageBox.Show(m.ToString());
            // 4. 读取某单元格的数据内容：
            for (int i = PicRows; i < AllRows; i++)
            {
                Picname = ((Excel.Range)etrange.get_Item(i, NameColumns)).Text;
                PicNa.Add(i,Picname); 
            }

            //5. 写入某单元格的数据内容：

            // ((Excel.Range)etrange.get_Item(i, j)).Value = strData;

            //6. 关闭文件及相关资源：

            // Get_Close();

        }
        /// <summary>
        /// 判段图片是否存在，如果存在进行添加 
        /// </summary>
        private void Result()
        {
            foreach (KeyValuePair<int, string> kvp in PicNa)
            {
                
           
            Picname =kvp.Value;
            string PicturePath = SelectPicPath + "\\" + Picname.Trim() + ".jpg";
            if (File.Exists(PicturePath))
            {
                etsheet = (Excel.Worksheet)etbook.Worksheets.get_Item(1);

                etrange = ((Excel.Range)etrange.get_Item(kvp.Key, PicColumns));
                Nowrows = kvp.Key;//将写入图片的行数赋值
                InsertPicture(etrange, etsheet, PicturePath);
            }
            else
            {
               //和图片插入错误一样 ((Excel.Range)etrange.get_Item(kvp.Key, PicColumns)).Value ="图片不存在";
                textBox1.AppendText("图片：" + Picname + "不存在 \n");
                //文件不存在
            }
            }

        }

        private void Get_Close()
        {
            etbook.Close();
            etApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(etrange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(etsheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(etbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(etApp);
        }



        /// <summary>
        /// 开始执行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            PicNa.Clear();//清空集合
            Boolean bl = false;
            Fist(ref bl);
            if (bl == false)
            {
                return;
            }
            Eto();
            Result();
            Get_Close();
            MessageBox.Show("添加成功", "提示信息");
        }

        public int test;
        public float left,top;
        /// </summary>
        /// <param name="rng">Excel单元格选中的区域</param>
        /// <param name="PicturePath">要插入图片的绝对路径。</param>
        public void InsertPicture(Excel.Range rng, Excel._Worksheet sheet, string PicturePath)
        {
            rng.Select();
            float PicLeft, PicTop, PicWidth, PicHeight;
           
          
            try
            {
                //设置行高
               // etsheet.UsedRange.RowHeight = 200;
                //设置列宽
               // etsheet.UsedRange.ColumnWidth = 55;
                if (test == 0)
                {
                    PicLeft = Convert.ToSingle(rng.Left);
                    left = PicLeft;
                    
                   
                }
                if (test==1)
                {
                    PicTop = Convert.ToSingle(rng.Top);
                    top = PicTop;
                }

                PicTop = top * (Nowrows-1);
                PicWidth = Convert.ToSingle(rng.Width);
                PicHeight = Convert.ToSingle(rng.Height);
               // PicLeft = 55;
              //  PicWidth = 20;
               // PicHeight = 10;

                //参数含义：
                //图片路径
                //是否链接到文件
                //图片插入时是否随文档一起保存
                //图片在文档中的坐标位置 坐标
                //图片显示的宽度和高度

                sheet.Shapes.AddPicture(PicturePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, PicTop, PicWidth, PicHeight);
                 test++;

                
               

            }
            catch (Exception ex)
            {
                MessageBox.Show("错误：" + ex.Message);
            }
        }
        /// <summary>
        /// 选择xls文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                SelectName = dialog.FileName;

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = "选择图片存放目录";
            if (folder.ShowDialog() == DialogResult.OK)
            {

                SelectPicPath = folder.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.AppendText("使用说明：\n");
            textBox1.AppendText("1、选择Excel表格 图片所在目录\n");
            textBox1.AppendText("2、选择姓名所在列、添加图片列、添加起始行\n");
            textBox1.AppendText("3、点击开始   等待提示 保存信息\n\n\n\n");
            textBox1.AppendText("作者：小樊 QQ:393719509");
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }
        /// <summary>
        /// 判断数值是否为空
        /// </summary>
        /// <param name="rs"></param>
        public void Fist(ref  Boolean rs)
        {
            Boolean res = false;
            if (SelectName == string.Empty || SelectName == null)
            {
                MessageBox.Show("请选择一个表格");
            }
            else if (SelectPicPath == String.Empty || SelectPicPath == null)
            {
                MessageBox.Show("请选择图片所在路径");
            }
            else if (comboBox1.SelectedIndex == 0)
            {
                MessageBox.Show("请选择添加图片的起始行");
            }
            else if (comboBox2.SelectedIndex == 0)
            {
                MessageBox.Show("请选择图片所在的列数");
            }
            else if (comboBox3.SelectedIndex == 0)
            {
                MessageBox.Show("请选择姓名所在列数");
            }
            else
            {
                res = true;
            }
            rs = res;

        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != 0)
            {
                PicRows = Convert.ToInt32(comboBox1.SelectedItem);
            }
        
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex!=0)
            {
                
           
            PicColumns = Convert.ToInt32(comboBox2.SelectedItem.ToString());
            }
        }

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex!=0)
            {
                
            
            NameColumns = Convert.ToInt32(comboBox3.SelectedItem.ToString());
            }
        }


    }
}
