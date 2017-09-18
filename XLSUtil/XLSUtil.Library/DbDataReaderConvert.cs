using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSUtil.Library
{
    public static class DbDataReaderConvert
    {
        public static List<List<object>> Convert2List(this DbDataReader dbDataReader, bool skipFirstLine = true)
        {
            var rawData = new List<List<object>>();
            if (dbDataReader.HasRows)
            {
                if (skipFirstLine)
                    dbDataReader.Read();

                while (dbDataReader.Read())
                {
                    var line = new List<object>();
                    for (int i = 0; i < dbDataReader.FieldCount; i++)
                        line.Add(dbDataReader[i]);

                    rawData.Add(line);
                }
            }
            else
            {
                return null;
            }

            return rawData;
        }

    }
}
