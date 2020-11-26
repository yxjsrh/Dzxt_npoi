using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Dzxt
{

    public class NPOIService
    {


        public static List<T> ImportToList<T>(string fileName) where T : new()
        {

            try
            {
                //创建文件流读取文件
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    //1.创建工作簿
                    string extName = Path.GetExtension(fileName);
                    dynamic workbook = null;
                    if (extName == ".xls")
                    {
                        HSSFWorkbook objHSSF = new HSSFWorkbook(fs);
                        workbook = objHSSF;
                    }
                    else
                    {

                        XSSFWorkbook objXSSF = new XSSFWorkbook(fs);
                        workbook = objXSSF;
                    }
                    //2.获取第一个工作表
                    ISheet sheet = workbook.GetSheetAt(0);
                    Type type = typeof(T);
                    PropertyInfo[] propertyInfos = type.GetProperties();
                    //3.循环将execl中的数据读到集合中
                    List<T> modelList = new List<T>();

                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i); //获取行，排除标题
                        if (row != null)
                        {
                            T model = new T();
                            for (int j = 0; j < propertyInfos.Length; j++)
                            {
                                var cell = row.GetCell(j);//根据行和列号获取cell对象
                                object value;//因为不确定泛型累成员的具体数据类型
                                if (cell != null)
                                {
                                    string str = propertyInfos[j].PropertyType.Name;
                                    switch (str)
                                    {
                                        case "String":
                                            value = cell.ToString();
                                            break;
                                        case "Decimal":
                                            value = Convert.ToDecimal(cell.ToString());
                                            break;
                                        case "Double":
                                            value = Convert.ToDouble(cell.ToString());
                                            break;
                                        case "Int16":
                                        case "Int32":
                                        case "Int64":
                                            value = Convert.ToInt32(cell.ToString());
                                            break;
                                        case "DataTime":
                                            value = Convert.ToDateTime(cell.ToString());
                                            break;
                                        case "Boolean":
                                            value = Convert.ToBoolean(cell.ToString());
                                            break;
                                        default:
                                            value = cell.ToString();
                                            break;
                                    }
                                    //给实体对象赋值
                                    propertyInfos[j].SetValue(model, value, null);
                                }

                            }
                            modelList.Add(model);
                        }
                    }
                    return modelList;


                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        //将泛型集合中得实体导出到指定得execl
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">泛型方法类型</typeparam>
        /// <param name="fileName">execl路径和文件名</param>
        /// <param name="dataList">包含若干对象得泛型集合</param>
        /// <param name="columnNames">实体列得中文标题词典</param>
        /// <param name="version">execl版本号，规定0:2007以下,1:2007及以上</param>
        /// <returns>成功返回true</returns>
        public static bool ExportExecl<T>(string fileName, List<T> dataList, Dictionary<string, string> columnNames, int version = 0) where T : class
        {
            //1.基于NPOI创建工作薄
            HSSFWorkbook hssf = new HSSFWorkbook();//2007以下
            XSSFWorkbook xssf = new XSSFWorkbook();//2007以上
            //根据不同版本创建不同得对象
            IWorkbook workbook = null;
            if (version == 0)
            {
                workbook = hssf;
            }
            else
            {
                workbook = xssf;
            }
            //2.创建工作表
            ISheet sheet = workbook.CreateSheet("sheet1");
            //3.生成标题和设置样式
            IRow rowTitle = sheet.CreateRow(0);

            Type type = typeof(T);
            PropertyInfo[] propertyInfos = type.GetProperties();//获取类型得属性
            for (int i = 0; i < propertyInfos.Length; i++)
            {
                ICell cell = rowTitle.CreateCell(i);//创建单元格对象
                cell.SetCellValue(columnNames[propertyInfos[i].Name]);
                SetCellStyle(workbook, cell);
                SetColumnwIDTH(sheet, i, 20);
            }
            //4.循环实体集合 生成数据
            for (int i = 0; i < dataList.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < propertyInfos.Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    T model = dataList[i];  //根据泛型找到具体化得实体对象
                    string value = propertyInfos[j].GetValue(model, null).ToString();//基于反射获取实体属性
                    cell.SetCellValue(value);//赋值
                    SetCellStyle(workbook, cell);

                }
            }
            //5.保存为execl文件
            using (FileStream fs = File.OpenWrite(fileName))
            {
                workbook.Write(fs);
                return true;
            }

        }


        /// <summary>
        /// 设置cell单元格边框
        /// </summary>
        /// <param name="workbook">接口类型工作簿</param>
        /// <param name="cell">cell单元格对象</param>
        private static void SetCellStyle(IWorkbook workbook, ICell cell)
        {

            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            IFont font = workbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeight = 15 * 15;
            font.Color = IndexedColors.Black.Index;
            style.SetFont(font);
            cell.CellStyle = style;



        }
        /// <summary>
        /// 设置cell单元格列宽
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">第几列</param>
        /// <param name="width">具体宽度</param>
        private static void SetColumnwIDTH(ISheet sheet, int index, int width)
        {
            sheet.SetColumnWidth(index, width * 256);
        }
    }
}
