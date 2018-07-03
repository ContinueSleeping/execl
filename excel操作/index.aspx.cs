using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.IO;


namespace excel操作
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ///实现最基本 简单的  excel  表格导入导出
            /// 一、自定义ExcelHelper类的使用
            /// 1、使用自定义ExcelHelper类 实现数据库中数据导出
            ///      a、从数据库中查询出数据 以 DataTable类型存储
            ///      b、使用ExcelHelper中GetDataTableToExcelStream方法得到 内存流
            ///      c、将内存流响应给客户端
            ///      Response.ContentType = "application/vnd.ms-excel";
            ///             //将响应类型 改为 excel 表格（告诉浏览器，响应给他的数据是什么格式的）
            ///      Response.AddHeader("Content-Disposition", "attachment;Filename=表格名称.xlsx");
            ///             //指定文件名称
            ///      Response.BinaryWrite(ms.ToArray());
            ///      Response.End();
            /// 2、使用自定义ExcelHelper类 实现Excel数据导入数据库中
            ///      a、得到上传文件，并保存在服务器磁盘
            ///      b、调用ExcelHelper中GetExcelToDataTable方法得到 DataTable 数据
            ///      c、遍历DataTable 数据 将 数据存入 数据库中
            /// 二、实现自定义ExcelHelper类
            /// 1、实现GetExcelToDataTable //将 Excel 表转换成 DataTable 
            ///      a、创建 dataTable 
            ///      b、读取Excel表格 使用文件流保存 
            ///      c、根据Excel后缀名判断 版本号 根据不同的版本创建不同的 类
            ///      d、为
            ///      e、遍历 Excel 表 的行
            ///         从 dataTable 中创建出 行
            ///      g、遍历 每一行的的所有单元格 根据不同的数据类型 从单元格中取出数据添加到 dataTable
            /// 
            /// 2、实现GetDataTableToExcelStream //将 DataTable 表转换成 Excel
            /// 
            /// 
            /// 3、参考资料：
            /// NPOI 是一个开源的C#读写Excel、WORD等微软OLE2组件文档的项目。
            /// 
            /// IWorkbook workbook //工作簿 Excel 文件 
            /// ISheet sheet //表格
            /// IRow row //行
            /// ICell cell//列
            /// row.LastCellNum 最后有数据单元格的下标
            /// cell.CellType 单元格的数据类型 
            /// 
           
        }
        static DataTable table=null;
        protected void BtnUp_Click(object sender, EventArgs e)
        {

            string filePath = "Down/" + FileUpload1.FileName;

            //将虚拟路径转换为真的服务器磁盘路径
            FileUpload1.SaveAs(Server.MapPath(filePath));

            DataTable table = ExcelHelper.ExcelToDataTable(Server.MapPath(filePath));

            foreach (DataRow item in table.Rows)
            {
                string n = item[0].ToString();
            }


            GridView1.DataSource = table;

            GridView1.DataBind();

            //int i = table.Rows.Count;


            //工作簿
        }

        protected void BtnDown_Click(object sender, EventArgs e)
        {


            MemoryStream ms = ExcelHelper.DataTableToExcel(table);

            //以什么格式响应
            Response.ContentType = "application/vnd.ms-excel";
            
            //告诉浏览器这不是预览是下载 下载的文件名为
            Response.AddHeader("Content-Disposition", "attachment;Filename=书籍.xlsx");

            //将内存流以2进制字符写入http输出流
            Response.BinaryWrite(ms.ToArray());
            //响应结束
            Response.End();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string sql = "select bookinfo.Id,BookName,BookTypeName,author,press from Bookinfo,bookTypeInfo where bookinfo.typeid = booktypeinfo.id";
            table = SqlHelper.GetDataTable(sql, false);
            GridView2.DataSource = table;
            GridView2.DataBind();
        }
    }
}