using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dzxt
{
    public partial class Form1 : Form
    {
        private static List<DzModel> ImportDzList = null;
        private static List<DzModel> ExportDzList = null;

        public Form1()
        {
            InitializeComponent();
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //导入


            OpenFileDialog openFile = new OpenFileDialog();
            DialogResult res = openFile.ShowDialog();
            if (res == DialogResult.OK)
            {
                string path = openFile.FileName;
                //导入得到的集合
                try
                {
                    ImportDzList = NPOIService.ImportToList<DzModel>(path);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("请关闭文件后重试");

                }

            }

            if (ImportDzList != null)
            {
                //把导入的集合直接给导出的集合
                //ExportDzList = ImportDzList;
                //处理
                ExportDzList = Cl(ImportDzList);

            }

            //导出
            Dictionary<string, string> columnNames = new Dictionary<string, string>();
            columnNames.Add("standard", "标准列");
            columnNames.Add("contrast", "对比列");
            columnNames.Add("result", "结果");

            //调用导出方法
            bool result = NPOIService.ExportExecl<DzModel>("result.xlsx", ExportDzList, columnNames, 1);
            if (result)
            {
                MessageBox.Show("处理成功");

                Application.Exit();

                //DialogResult dialog =  MessageBox.Show("处理成功是否打开文件","处理成功",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                //if (dialog == DialogResult.Yes)
                //{
                //    Process.Start("result.xls");
                //}
            }
        }

        private List<DzModel> Cl(List<DzModel> dzModels)
        {

            List<string> one = new List<string>();
            List<string> two = new List<string>();

            List<DzModel> model = new List<DzModel>();
            //创建两列对比
            foreach (var item in dzModels)
            {
                if (string.IsNullOrWhiteSpace(item.standard))
                {
                    one.Add("0");
                }
                else
                {
                    one.Add(item.standard);
                }

                if (string.IsNullOrWhiteSpace(item.contrast))
                {
                    two.Add("0");
                }
                else
                {
                    two.Add(item.contrast);
                }
            }
            for (int i = 0; i < two.Count; i++)
            {
                int count = 0;
                string msg = "";
                bool flag = false;
                DzModel dz = new DzModel();
                if (i <= two.Count)
                {
                    dz.contrast = two[i];
                }
                else
                {
                    dz.contrast = "无数据";
                }

                if (i <= one.Count)
                {
                    dz.standard = one[i];
                }
                else
                {
                    dz.standard = "无数据";
                }
                for (int j = 0; j < one.Count; j++)
                {
                    if (two[i].Equals("0"))
                    {
                        flag = true;
                        msg = "对比列为0";
                        break;
                    }
                    if (two[i].Equals(one[j]))
                    {
                        count++;
                        msg += (j + 2) + "/";
                    }
                }
                if (flag)
                {
                    dz.result = msg;
                }
                else
                {
                    if (count == 0)
                    {
                        dz.result = "无";
                    }
                    else
                    {
                        dz.result = "找到" + count + "条数据" + "分别在" + msg + "行";
                    }
                }


                model.Add(dz);
            }

            return model;

        }

    }
}
