using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FFF.Core.Data.Attributes;
using FFF.Core.Data;

namespace ToolsLibrary.Entity
{
    public class ColumnMap
    {
        public const string Query="Query";

        [DataContextMap(MapKey = Query, SQLContextType = EnumSQLContext.SQLString)]
        public ColumnMap()
        {
            
        }

        /// <summary>
        /// Column Name
        /// </summary>
       [DataContextColumn(MapKey =Query, ColumnName = "COLUMN_NAME")]
        public string ColumnName{get;set;}

        /// <summary>
        /// Column Seq ID
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "COLUMN_ID")]
        public int ColumnID{get;set;}
        
        /// <summary>
        /// Default Value
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "DATA_DEFAULT")]
        public object DefaultData{get;set;}

        /// <summary>
        /// Nullable
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "NULLABLE")]
        public string NullAble{get;set;}

        /// <summary>
        /// Data Type
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "DATA_TYPE")]
        public string DataType{get;set;}

        /// <summary>
        /// Data Length
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "DATA_LENGTH")]
        public int DataLength{get;set;}

        /// <summary>
        /// Data Precision
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "DATA_PRECISION")]
        public decimal DataPresision{get;set;}
        
        /// <summary>
        /// Comment
        /// </summary>
        [DataContextColumn(MapKey =Query, ColumnName = "COMMENTS")]
        public string Comments{get;set;}
  
    }
}
