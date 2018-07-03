using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI;
//对应的是excel低版本的
using NPOI.HSSF.UserModel;
//对应的是excel高版本的
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.IO;

namespace excel操作
{
    public class ExcelHelper
    {
        public static DataTable ExcelToDataTable(string path)
        {
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;

            DataTable table = new DataTable();
            DataRow dataRow = null;


            try
            {
                using (FileStream fs = File.OpenRead(path))
                {
                    // Path.GetExtension(path) 获取文件的后缀名
                    if (Path.GetExtension(path) == ".xlsx")
                        //从流中获取excel 工作簿
                        workbook = new XSSFWorkbook(fs);
                    else
                        workbook = new HSSFWorkbook(fs);

                    //从工作簿里获取指定的表
                    sheet = workbook.GetSheetAt(0);

                    for (int i = 0; i < sheet.GetRow(0).LastCellNum; i++)
                    {
                        //根据excel表里的列数为datatable创建列
                        table.Columns.Add(new DataColumn());
                    }

                    //根据表格的行数进行循环创建dataRow
                    for (int i = 0; i < sheet.LastRowNum; i++)
                    {
                        //为datatable创建行
                        dataRow = table.NewRow();
                        //从excel表中获取指定的行
                        row = sheet.GetRow(i);
                     
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //从行里获取单元格
                            cell = row.GetCell(j);
                            switch (cell.CellType)
                            {
                                //如果excel单元格内的值的数据类型为 数值
                                case CellType.Numeric:
                                    //将excel表中单元格的值赋给datatable行中的某列
                                    dataRow[j] = cell.NumericCellValue;
                                    break;
                                case CellType.String:
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                                case CellType.Blank:
                                    dataRow[j] = "";
                                    break;
                                default:
                                    throw new Exception("格式错误");                                 
                            }
                        }
                        //将创建的行添加到datatable的rows（行集合）
                        table.Rows.Add(dataRow);
                    }
                    //返回结果
                    return table;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //无论文件上传成功还是失败 都要讲文件删除释放内存
                File.Delete(path);
            }          
        }

        public static MemoryStream DataTableToExcel(DataTable table)
        {

            //创建工作簿
            IWorkbook workbook = new XSSFWorkbook();
            //创建工作簿里的表
            ISheet sheet = workbook.CreateSheet(); ;
            //创建表里的行
            IRow row = null;
            //创建行里的列
            ICell cell = null;


            try
            {
                //创建内存流  保存excel
                MemoryStream ms = new MemoryStream();

                //创建一个excel表的一行 放的是标题列
                row = sheet.CreateRow(0);
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    //创建单元格
                    cell = row.CreateCell(i);
                    //给单元格进行赋值
                    cell.SetCellValue(table.Columns[i].ColumnName);
                    row.Cells.Add(cell);
                }
                //从excel表格的第二行开始
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    //在excel中创建一行
                    row = sheet.CreateRow(i + 1);

                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        //创建单元格
                        cell = row.CreateCell(j);
                        //给单元格进行赋值
                        cell.SetCellValue(table.Rows[i][j].ToString());
                        row.Cells.Add(cell);
                    }
                }
                //将工作簿放入内存流中
                workbook.Write(ms);
                return ms;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}