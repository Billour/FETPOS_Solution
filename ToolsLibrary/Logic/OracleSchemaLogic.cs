using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ToolsLibrary.Entity;
using FFF.Core.Data;

namespace ToolsLibrary.Logic
{
    public class OracleSchemaLogic
    {
        public SchmaMap GetOracleSchema()
        {
            SchmaMap map = new SchmaMap();

            map.Tables = GetTables();

            map.Tables.ForEach(p=>{

                p.ColumnList = GetColumns(p.TableName);
            });

            return map;

            
        }

        /// <summary>
        /// Get Tables
        /// </summary>
        /// <returns></returns>
        public List<TableMap> GetTables()
        {
            string sql = "select *  from all_all_tables where owner = 'WEBPOS' order by table_name";
            DataContext context = new DataContext("Portal");

            return context.QuerySelect<TableMap>(sql, TableMap.Query); 
        }

        /// <summary>
        /// 取回Column List
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public List<ColumnMap> GetColumns(string tableName)
        {
            string sql = @"select 
                      ALL_TAB_COLUMNS.COLUMN_NAME,  
                      COLUMN_ID,
                      DATA_DEFAULT, 
                      NULLABLE,  
                      DATA_TYPE,  
                      DATA_PRECISION, 
                      COMMENTS, DATA_LENGTH 
                      from ALL_TAB_COLUMNS, All_Col_Comments
                      where  ALL_TAB_COLUMNS.OWNER = ALL_COL_COMMENTS.Owner AND 
                      ALL_TAB_COLUMNS.TABLE_NAME = ALL_COL_COMMENTS.TABLE_NAME AND 
                      ALL_TAB_COLUMNS.COLUMN_NAME = all_col_comments.column_name and 
                      ALL_TAB_COLUMNS.OWNER  = 'WEBPOS'  and ALL_TAB_COLUMNS.TABLE_NAME = '{0}'";

            DataContext context = new DataContext("Portal");

            return context.QuerySelect<ColumnMap>(sql, ColumnMap.Query, tableName);
        }
    }
}
