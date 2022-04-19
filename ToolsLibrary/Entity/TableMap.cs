using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FFF.Core.Data.Attributes;
using FFF.Core.Data;

namespace ToolsLibrary.Entity
{
    public class TableMap
    {
        

        public const string Query = "Query";

        [DataContextMap(MapKey = Query, SQLContextType = EnumSQLContext.SQLString)]
        public TableMap()
        { 
            
        }

        /// <summary>
        /// Table Name
        /// </summary>
        [DataContextColumn(MapKey = Query, ColumnName = "TABLE_NAME")]
        public string TableName { get; set; }

        ///// <summary>
        ///// Primary Key 
        ///// </summary>
        //public List<string> PrimaryKey { get; set; }
        
        ///// <summary>
        ///// Index Key
        ///// </summary>
        //public List<string> IndexKey { get; set; }
        
        /// <summary>
        /// Column List
        /// </summary>
        public List<ColumnMap> ColumnList { get; set; }
    }
}
