using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;

namespace EtcJob.HelpClass
{
    public static class ExcelTo
    {
        /// <summary>
        /// 将Excel文件转换成List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public static List<T> ImportExcelToList<T>(string excelFilePath)
        {
            var table = ReadExcelToDataTable(excelFilePath);
            return ConvertDataTableToList<T>(table);
        }

        #region List => DateTable
        /// <summary>    
        /// 将List转换成DataTable    
        /// </summary>    
        /// <param name="tList">集合</param>    
        /// <returns></returns>    
        //public static DataTable ConvertListToDataTable<T>(List<T> tList)
        //{
        //    var dt = new DataTable();
        //    Type tp = typeof(T);
        //    PropertyInfo[] propertyInfos = tp.GetProperties();

        //    foreach (var prop in propertyInfos)
        //    {
        //        var description = GetDescriptionByField<T>(prop.Name);
        //        if (string.IsNullOrEmpty(description)) continue;
        //        dt.Columns.Add(prop.Name, prop.PropertyType);
        //        dt.Columns[prop.Name].ColumnName = description;
        //    }

        //    foreach (var item in tList)
        //    {
        //        DataRow dr = dt.NewRow();
        //        foreach (PropertyInfo proInfo in propertyInfos)
        //        {
        //            object obj = proInfo.GetValue(item);
        //            try
        //            {
        //                dr[GetDescriptionByField<T>(proInfo.Name)] = obj;
        //            }
        //            catch (Exception)
        //            {

        //            }
        //        }
        //        dt.Rows.Add(dr);
        //    }
        //    return dt;
        //}

        /// <summary>
        /// 将集合类转换成DataTable   现用
        /// </summary>
        /// <param name="tList"></param>
        /// <returns></returns>
        public static DataTable ConvertListToDataTable<T>(List<T> tList)
        {
            var dt = new DataTable();
            Type type = typeof(T);
            PropertyInfo[] pArray = type.GetProperties();
            var dictKeys = new Dictionary<string, string>();

            foreach (var p in pArray)
            {
                var description = GetDescriptionByField<T>(p.Name);
                if (string.IsNullOrEmpty(description)) continue;
                dictKeys.Add(p.Name, description);
                dt.Columns.Add(p.Name, p.PropertyType);
                dt.Columns[p.Name].ColumnName = description;
            }

            foreach (var item in tList)
            {
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo p in pArray)
                {
                    if (dictKeys.Count == 0) break;
                    if (!dictKeys.ContainsKey(p.Name)) continue;
                    object obj = p.GetValue(item);
                    dr[dictKeys[p.Name]] = obj;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 将集合类转换成DataTable   现用
        /// </summary>
        /// <param name="tList"></param>
        /// <returns></returns>
        public static DataTable ConvertListToDataTable<T>(List<T> tList, Dictionary<string, string> dictKeys)
        {
            var dt = new DataTable();
            Type type = typeof(T);
            PropertyInfo[] pArray = type.GetProperties();

            foreach (var p in pArray)
            {
                if (!dictKeys.ContainsKey(p.Name)) continue;
                var description = dictKeys[p.Name];
                if (string.IsNullOrEmpty(description)) continue;
                dt.Columns.Add(p.Name, p.PropertyType);
                dt.Columns[p.Name].ColumnName = description;
            }

            foreach (var item in tList)
            {
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo p in pArray)
                {
                    if (dictKeys.Count == 0) break;
                    if (!dictKeys.ContainsKey(p.Name)) continue;
                    object obj = p.GetValue(item);
                    dr[dictKeys[p.Name]] = obj;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 将集合类转换成DataTable,根据dictKeys的顺序排序   现用
        /// </summary>
        /// <param name="tList"></param>
        /// <returns></returns>
        public static DataTable ConvertListToDataTableOrderBy<T>(List<T> tList, Dictionary<string, string> dictKeys)
        {
            var dt = new DataTable();
            Type type = typeof(T);
            PropertyInfo[] pArray = type.GetProperties();
            var dictionary = new Dictionary<string, Type>();//用于存储列名和描述

            foreach (var p in pArray)
            {
                if (!dictKeys.ContainsKey(p.Name)) continue;
                var description = dictKeys[p.Name];
                if (string.IsNullOrEmpty(description)) continue;
                dictionary.Add(p.Name, p.PropertyType);
            }

            foreach (var item in dictKeys)
            {
                if (!dictionary.ContainsKey(item.Key)) continue;
                dt.Columns.Add(item.Key, dictionary[item.Key]);
                dt.Columns[item.Key].ColumnName = item.Value;
            }

            foreach (var item in tList)
            {
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo p in pArray)
                {
                    if (dictKeys.Count == 0) break;
                    if (!dictKeys.ContainsKey(p.Name)) continue;
                    object obj = p.GetValue(item);
                    dr[dictKeys[p.Name]] = obj;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        #endregion

        #region DateTable => List
        /// <summary>
        /// 将DataTable转化为List
        /// </summary>
        /// <typeparam name="T">实体对象</typeparam>
        /// <param name="dt">datatable表</param>
        /// <param name="isStoreDB">是否存入数据库datetime字段，date字段没事，取出不用判断</param>
        /// <returns>返回list集合</returns>
        public static List<T> ConvertDataTableToList<T>(DataTable dt)
        {
            if (dt == null || dt.Rows == null || dt.Columns == null || dt.Columns.Count == 0) return null;
            Type type = typeof(T);
            List<T> list = new List<T>();
            var columns = new List<object>();
            PropertyInfo[] pArray = type.GetProperties(); //集合属性数组
            var dictionary = new Dictionary<string, string>();

            foreach (var item in dt.Columns)
            {
                columns.Add(item.ToString());
                dictionary.Add(item.ToString(), item.ToString());
            }

            if (dt.Columns.Contains("F1"))
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (!dictionary.ContainsKey(dt.Columns[i].ToString())) continue;
                    dictionary[dt.Columns[i].ToString()] = dt.Rows[0].ItemArray[i].ToString();
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                if (columns.Contains("F1"))
                {
                    columns = dt.Rows[0].ItemArray.ToList();
                    continue;
                }

                T entity = Activator.CreateInstance<T>(); //新建对象实例 
                foreach (PropertyInfo p in pArray)
                {
                    var description = GetDescriptionByField<T>(p.Name);
                    if (!dictionary.ContainsValue(description)) continue;
                    description = dictionary.FirstOrDefault(m => m.Value == description).Key;
                    if (string.IsNullOrEmpty(description)) continue;
                    if (row[description] == null || row[description] == DBNull.Value) continue;  //DataTable列中不存在集合属性或者字段内容为空则，跳出循环，进行下个循环   
                    var dateTime = new DateTime();
                    if (p.PropertyType == typeof(DateTime)
                        && DateTime.TryParse(Convert.ToString(row[description]), out dateTime)
                        && Convert.ToDateTime(row[description]) < Convert.ToDateTime("1900-01-01"))
                        continue;
                    try
                    {
                        var obj = Convert.ChangeType(row[description], p.PropertyType);//类型强转，将table字段类型转为集合字段类型  
                        p.SetValue(entity, obj, null);
                    }
                    catch (Exception)
                    {

                    }
                }
                list.Add(entity);
            }
            return list;
        }
        #endregion

        /// <summary>
        /// 获取Description标签的值
        /// </summary>
        /// <param name="field"></param>
        /// <returns></returns>
        private static string GetDescriptionByField<T>(string field)
        {
            var type = typeof(DescriptionAttribute);
            AttributeCollection attributes = TypeDescriptor.GetProperties(typeof(T))[field].Attributes;
            DescriptionAttribute myAttribute = (DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
            return myAttribute.Description;
        }

        #region Excel => DateTable
        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public static DataTable ReadExcelToDataTable(string excelFilePath)
        {
            string strExtension = Path.GetExtension(excelFilePath);
            //Excel的连接
            OleDbConnection objConn = null;
            switch (strExtension)
            {
                case ".xls":
                    objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";" + "Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"");
                    break;
                case ".xlsx":
                    objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";" + "Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1;\"");
                    break;
                case ".csv":
                    objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";" + "Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"");
                    break;
                default:
                    objConn = null;
                    break;
            }
            objConn.Open();

            DataTable sheetNames = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            var sheet1 = sheetNames.Rows[0][2];
            string strSql = "select * from [" + sheet1 + "]";
            OleDbDataAdapter oleDb = new OleDbDataAdapter(strSql, objConn);
            DataSet ds = new DataSet();
            oleDb.Fill(ds);
            objConn.Close();
            if (ds == null) return null;
            return ds.Tables[0];
        }
        #endregion
    }
}
