using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DgvSavePicDemo
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!ckbCl1.Checked && !ckbCl10.Checked && !ckbCl2.Checked && ckbCl3.Checked && !ckbCl4.Checked && !ckbCl5.Checked && !ckbCl6.Checked && !ckbCl7.Checked && !ckbCl8.Checked && !ckbCl9.Checked)
            {
                MessageBox.Show("至少选择一列");
                return;
            }
            int startId = Convert.ToInt32(numStart.Value);
            int endId = Convert.ToInt32(numEnd.Value);
            if (startId > endId)
            {
                MessageBox.Show("开始行数不能大于结束行数");
                return;
            }
            DataTable dt = GetDgvToTable(dgvTest);  //获取到所选数据
            
            //MessageBox.Show(string.Format("rows: {0},cols: {1}", ResultRows, dt.Columns.Count));
            Image result = GetImageByDt(dt);
            SaveFileDialog objSaveDialog = new SaveFileDialog();
            objSaveDialog.FileName = "outPic" + DateTime.Now.ToString("yyMMddHHmmss");
            objSaveDialog.Filter = "JPG(*.jpg)|*.jpg";
            if (objSaveDialog.ShowDialog() == DialogResult.OK)

                result.Save(objSaveDialog.FileName);

        }
        private Image GetImageByDt(DataTable dtPar)
        {
            const int rowHeight = 25;  //行高
            const int titleHeight = 30;//标题高
            const int colWidth = 100;  //列宽
            const int widWape = 5; //边宽
            int startId = Convert.ToInt32(numStart.Value);
            int endId = Convert.ToInt32(numEnd.Value);
            int ResultRows = endId - startId + 1;



            int picW = widWape * 2 + colWidth * dtPar.Columns.Count;
            int picH =titleHeight + widWape * 2 + rowHeight * ResultRows;
            Image ResultImg = new Bitmap(picW, picH);
            Graphics g = Graphics.FromImage(ResultImg);
            g.Clear(Color.White); 
            //Application.DoEvents();
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.SmoothingMode = SmoothingMode.HighQuality; //高绘图质量
            g.InterpolationMode = InterpolationMode.HighQualityBilinear;
            Brush myBrush = new SolidBrush(Color.LightGray);
            g.FillRectangle (myBrush, widWape, widWape, picW - widWape, titleHeight); //填充标题背景
            myBrush = new SolidBrush(Color.Black);
            Pen myPen = new Pen(myBrush, 2);
            g.DrawRectangle(myPen, widWape, widWape, picW - widWape*2, picH -widWape*2 );  //绘制表格外框
            myPen = new Pen(myBrush, 1);
            //绘制内部表格线
            int temY = widWape + titleHeight;
            for (int i = 0; i < dtPar.Rows.Count; i++)  //横向
            {
                g.DrawLine(myPen, new Point(widWape, temY), new Point(picW - widWape, temY));
                temY += rowHeight;
            }
            int temX = widWape + colWidth;
            if (dtPar.Columns.Count > 0)  //纵向
            {
                for (int i = 0; i < dtPar.Columns.Count -1; i++)
                {
                    g.DrawLine(myPen, new Point(temX , widWape ), new Point(temX , picH -widWape ));
                    temX += colWidth ;
                }
            }
            //字体及字号
           Font  titleFont = new Font("黑体", 12, FontStyle.Regular);
            Font conteFont = new Font("宋体", 9, FontStyle.Regular);
            //绘制标题文字
            temX = widWape + 1;
            temY = widWape + 6;
            for (int i = 0; i < dtPar.Columns.Count ; i++)
            {
                g.DrawString(dtPar.Columns[i].Caption, titleFont, myBrush, new Point(temX, temY));
                temX += colWidth;
            }

            //绘制正文内容
            temX = widWape + 1;
            temY = widWape + titleHeight + 5;
            for (int j = startId-1 ; j <= endId-1  ; j++)
            {
                for (int i = 0; i < dtPar.Columns.Count; i++)
                {
                    if (dtPar.Rows[j][i]  is Bitmap  )
                    {
                        g.DrawImage((Bitmap)dtPar.Rows[j][i], new Rectangle(temX, temY-4,colWidth -2,rowHeight -1));  //Rectangle为图片绘制的坐标和宽高，根据你DGV调节

                    }
                    else
                    {
                        g.DrawString(dtPar.Rows[j][i].ToString(), conteFont, myBrush, new Point(temX, temY));
                    }



                    temX += colWidth;
                }
                temX = widWape + 1;
                temY += rowHeight;
            }


            g.Save();
            return ResultImg;


        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            //添加测试数据
            List<tableData> dataList = new List<DgvSavePicDemo.tableData>();
            for (int i = 0; i < 200; i++)
            {
                //替换为本地测试图片地址
                Image temImage = Image.FromFile(@"D:\My Documents\My Pictures\92b3a2fad8bcbcabc0d6b1a6bcf8ece10e18.jpg");
                int col = 0;
                dataList.Add(new tableData()
                {
                    cl1 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl2 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl3 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl4 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl5 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl6 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl7 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl8 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl9 = string.Format("第{0}行 第{1}列", i + 1, ++col),
                    cl10 = temImage,
                });
            }
            dgvTest.DataSource = dataList;
            cl1.Visible = false;
            cl2.Visible = false;
            cl3.Visible = false;
        }
        private DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count-1; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].HeaderText.ToString());
                dt.Columns.Add(dc);
            }
            dt.Columns.Add("列10",typeof(Image) );
            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count-1; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dr[9] = (Bitmap)(dgv.Rows[count].Cells[9].Value);

                //这里增加图片获取


                dt.Rows.Add(dr);
            }
            //去掉未选择的列
            if (!ckbCl10.Checked) dt.Columns.Remove("列10");
            if (!ckbCl9.Checked) dt.Columns.Remove("列9");
            if (!ckbCl8.Checked) dt.Columns.Remove("列8");
            if (!ckbCl7.Checked) dt.Columns.Remove("列7");
            if (!ckbCl6.Checked) dt.Columns.Remove("列6");
            if (!ckbCl5.Checked) dt.Columns.Remove("列5");
            if (!ckbCl4.Checked) dt.Columns.Remove("列4");
            if (!ckbCl3.Checked) dt.Columns.Remove("列3");
            if (!ckbCl2.Checked) dt.Columns.Remove("列2");
            if (!ckbCl1.Checked) dt.Columns.Remove("列1");
            return dt;
        }
        /// <summary>
          /// DataGridView跨越滚动条截图
          /// </summary>
          /// <param name="dgv">DataGridView</param>
          /// <returns>图形</returns>
          private  Image GetDataGridView(DataGridView dgv)
          {
            if (!ckbCl10.Checked) cl10.Visible = false;
            if (!ckbCl9.Checked) cl9.Visible = false;
            if (!ckbCl8.Checked) cl8.Visible = false;
            if (!ckbCl7.Checked) cl7.Visible = false;
            if (!ckbCl6.Checked) cl6.Visible = false;
            if (!ckbCl5.Checked) cl5.Visible = false;
            if (!ckbCl4.Checked) cl4.Visible = false;
            if (!ckbCl3.Checked) cl3.Visible = false;
            if (!ckbCl2.Checked) cl2.Visible = false;
            if (!ckbCl1.Checked) cl1.Visible = false;
            Application.DoEvents();
            PictureBox pic = new PictureBox();
              pic.Size = dgv.Size;
             pic.Location = dgv.Location;
             Bitmap bmpPre = new Bitmap(pic.Width, pic.Height);
             dgv.DrawToBitmap(bmpPre, new Rectangle(0, 0, pic.Width, pic.Height));
             pic.Image = bmpPre;          
             dgv.Parent.Controls.Add(pic);
 
             dgv.Visible = false;
             dgv.AutoSize = true;
 
             Bitmap bmpNew = new Bitmap(dgv.Width, dgv.Height);
 
             dgv.DrawToBitmap(bmpNew, new Rectangle(0, 0, dgv.Width, dgv.Height));
 
             dgv.AutoSize = false;
             dgv.Visible = true;
 
             dgv.Parent.Controls.Remove(pic);
             bmpPre.Dispose();
             pic.Dispose();
             return bmpNew;
         }

        private void btnSaveScrPic_Click(object sender, EventArgs e)
        {

            Image result = GetDataGridView(dgvTest);
            SaveFileDialog objSaveDialog = new SaveFileDialog();
            objSaveDialog.FileName = "outPic" + DateTime.Now.ToString("yyMMddHHmmss");
            objSaveDialog.Filter = "JPG(*.jpg)|*.jpg";
            if (objSaveDialog.ShowDialog() == DialogResult.OK)

                result.Save(objSaveDialog.FileName);
        }

        private void btnOtherTest_Click(object sender, EventArgs e)
        {
            //dgvTest.Visible = false;
            cl1.Visible = false;
            cl2.Visible = true;
            cl3.Visible = false;
            Application.DoEvents();
        }
    }
    public class tableData
    {
        public string  cl1 { get; set; }
        public string cl2 { get; set; }
        public string cl3 { get; set; }
        public string cl4 { get; set; }
        public string cl5 { get; set; }
        public string cl6 { get; set; }
        public string cl7 { get; set; }
        public string cl8 { get; set; }
        public string cl9 { get; set; }
        public Image  cl10 { get; set; }
    }
}
