using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;


namespace excel操作
{
    public static class SqlHelper
    {
        //从配置文件中读取连接字符串
        private static string connStr =
            ConfigurationManager.ConnectionStrings["connStr"].ConnectionString;
   
        /// <summary>
        /// 通用的获取SqlCommand对象
        /// </summary>
        /// <param name="sql">SQL语句 或 存储过程名字</param>
        /// <param name="isProc">是否为存储过程</param>
        /// <param name="paras">sql语句的参数或存储过程的参数</param>
        /// <returns></returns>
        private static SqlCommand GetSqlCommand(string sql,bool isProc = false,params SqlParameter[] paras)
        {
            SqlConnection conn = new SqlConnection(connStr);
            SqlCommand cmd = new SqlCommand(sql, conn);
            if (isProc)
            {
                //如果是存储过程，将cmd.CommandType属性改为System.Data.CommandType.StoredProcedure;
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
            }
            if (paras.Length > 0)
            {
                //如果参数长度大于0 向cmd的Parameters 中添加
                cmd.Parameters.AddRange(paras);
            }
            cmd.Connection.Open();
            return cmd;
        }

       /// <summary>
       /// 返回增、删、改是否执行成功
       /// </summary>
       /// <param name="sql"></param>
       /// <param name="isProc"></param>
       /// <param name="paras"></param>
       /// <returns></returns>
        public static bool ExectueNonQuery(string sql, bool isProc = false, params SqlParameter[] paras)
        {
            using (SqlCommand cmd = GetSqlCommand(sql,isProc,paras))
            {
                try
                {
                    return  cmd.ExecuteNonQuery() > 0 ? true : false; 
                }
                catch (Exception)
                {
                    throw;
                }              
            }
        }

        /// <summary>
        /// 返回多行多列的结果集
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="isProc"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public static SqlDataReader ExectueReader(string sql, bool isProc = false, params SqlParameter[] paras)
        {
            SqlCommand cmd = GetSqlCommand(sql, isProc, paras);
            return cmd.ExecuteReader();            
        }

        /// <summary>
        /// 返回第一行第一列的结果
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="isProc"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public static object ExectueScalar(string sql, bool isProc = false, params SqlParameter[] paras)
        {
            using (SqlCommand cmd = GetSqlCommand(sql, isProc, paras))
            {
                return cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// 返回一张datatable
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="isProc"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public static DataTable GetDataTable(string sql, bool isProc,params SqlParameter[] paras)
        {
            using (SqlCommand cmd = GetSqlCommand(sql,isProc,paras))
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = cmd;
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                return dataSet.Tables[0];
            }
        }

        /// <summary>
        /// 返回一个泛型集合
        /// </summary>
        /// <typeparam name="T">要查询的model</typeparam>
        /// <returns></returns>
        //public static List<T> FindAll<T>()
        //{
        //    //使用typeof方法获取T 的类型 
        //    Type type = typeof(T);
        //    //创建泛型集合
        //    List<T> list = new List<T>();
        //    string sql = $"select * from {type.Name}";
        //    using (SqlCommand cmd = GetSqlCommand(sql,false))
        //    {     
        //        SqlDataReader reader = cmd.ExecuteReader();
        //        while (reader.Read())
        //        {
        //            //使用反射创建相应类型的对象
        //            T t = (T)Activator.CreateInstance(type);
        //            //type.GetProperties()
        //            //获取类中所有的公共属性
        //            foreach (var prop in type.GetProperties())
        //            {
        //               // string name =  AttributeHelper.GetColumnName(prop);
        //                prop.SetValue(t, reader[name] is DBNull ? null : reader[name]);
        //            }
        //            list.Add(t);
        //        }
        //        return list;
        //    }
        //}

    }
}
