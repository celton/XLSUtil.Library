using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace XLSUtil.Library
{
    public class XLS2Model<T> where T : class
    {
        private readonly string _filePath;
        private readonly Expression<Func<T, Object>>[] _properties;

        public XLS2Model(string filePath, params Expression<Func<T, Object>>[] properties)
        {
            _filePath = filePath;
            _properties = properties;
        }

        public List<T> Get()
        {
            List<T> objList = (List<T>)Activator.CreateInstance(typeof(List<T>));

            var dataList = GetRawData();

            foreach (var line in dataList)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                for (int i = 0; i < _properties.Count(); i++)
                {
                    var column = line[i];
                    var property = _properties[i];

                    PropertyInfo propertyInfo = obj.GetType().GetProperty(GetCorrectPropertyName(property));

                    propertyInfo.SetValue(obj, CastPropertyValue(propertyInfo, column.ToString()), null);
                }
                objList.Add(obj);
            }

            return objList;
        }        

        private string GetCorrectPropertyName(Expression<Func<T, Object>> expression)
        {
            if (expression.Body is MemberExpression)
            {
                return ((MemberExpression)expression.Body).Member.Name;
            }
            else
            {
                var op = ((UnaryExpression)expression.Body).Operand;
                return ((MemberExpression)op).Member.Name;
            }
        }

        private object CastPropertyValue(PropertyInfo property, string value)
        {
            try
            {
                if (property == null || String.IsNullOrEmpty(value))
                    return null;
                if (property.PropertyType.IsEnum)
                {
                    Type enumType = property.PropertyType;
                    if (Enum.IsDefined(enumType, value))
                        return Enum.Parse(enumType, value);
                }
                if (property.PropertyType == typeof(bool))
                    return value == "1" || value == "true" || value == "on" || value == "checked";
                else if (property.PropertyType == typeof(Uri))
                    return new Uri(Convert.ToString(value));
                else
                    return Convert.ChangeType(value, property.PropertyType);
            }
            catch (Exception ex)
            {
                throw new Exception("Convertion type error.\nSheet has a incompatible field type", ex);
            }
        }

        private List<List<object>> GetRawData(bool skipFirstLine = true)
        {
            try
            {
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _filePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
                OleDbDataReader reader = null;
                using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
                {
                    oleDbConnection.Open();

                    OleDbCommand oleDbCommand = new OleDbCommand();
                    oleDbCommand.Connection = oleDbConnection;
                    var schemaDataTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (schemaDataTable == null || schemaDataTable.Rows.Count < 1)
                        throw new Exception("Cannot resolve first line.");

                    string firstSheetName = schemaDataTable.Rows[0]["TABLE_NAME"].ToString();

                    oleDbCommand.CommandText = "Select * from [" + firstSheetName + "]";

                    reader = oleDbCommand.ExecuteReader();

                    return reader.Convert2List(skipFirstLine);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Cannot get content.", ex);
            }
        }
    }
}
