using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Reflection;

namespace TalkUnderSun
{
    public static class ExportTo
    {
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

        ///// <summary>
        ///// 导出Excel文件V2
        ///// </summary>
        ///// <param name="tList"></param>
        ///// <returns></returns>
        //public static void ExportExcel<T>(List<T> tList, Dictionary<string, string> dictKeys)
        //{
        //    if (tList == null || tList.Count == 0) return null;
        //    if (dictKeys == null || dictKeys.Count == 0) return null;

        //    Hashtable hashTable = new Hashtable();
        //    Type tp = typeof(T);

        //    string ExportFileName = tp.Name.Remove(0, 5);//删除Model tp.Name.Remove(0,5)

        //    var dt = ConvertListToDataTable(tList, dictKeys);

        //    Export(dt, ExportFileName);

        //    return;
        //}

        /// <summary>
        /// 将文件上传到文件服务器
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static void Export(DataTable dt, string path, string filename)
        {
            byte[] byteStream = null;

            string file = path + "\\" + filename + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            FileStream fs = null;
            Stream outputStream = new MemoryStream();
            try
            {
                if (!Directory.Exists(path))//判断是否存在
                {
                    Directory.CreateDirectory(path);//创建新路径
                }
                if (File.Exists(file)) //判断文件的存在
                {
                    //存在文件
                    file = path + "\\" + filename + DateTime.Now.AddMinutes(1).ToString("yyyyMMddHHmmss") + ".xls";
                }

                fs = new FileStream(file, FileMode.Create, FileAccess.Write);
                outputStream = new ExcelHelper().Export(dt, EnumExcelType.XLS);

                // 文件写入到流中,从流中写入到文件
                byte[] bytes = new byte[outputStream.Length];
                //必须指定流的位置为开始位置，否则输出内容为空
                byteStream = bytes;
                outputStream.Position = 0;
                outputStream.Read(bytes, 0, (int)outputStream.Length);
                fs.Write(bytes, 0, bytes.Length);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                fs.Close();
                outputStream.Close();
            }
        }
    }
}
