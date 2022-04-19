using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word= Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using ToolsLibrary.Entity;
using log4net;

namespace ToolsLibrary.Helper
{
    public class WordCore
    {
        private ILog log = log4net.LogManager.GetLogger(typeof(WordCore));

        #region Field
        public static Object Missing = System.Reflection.Missing.Value;
        public static Object missing = System.Reflection.Missing.Value;

        Word.Application wApp = new Word.Application() {Visible=false};
        Word.Document wDoc = new Word.Document();
        object oEndOfDoc = "\\endofdoc";

        #endregion

        #region Open&Close&Save
        public Boolean Open(String FileName)
        {
            try
            {
                wApp.Visible = true;
                Boolean visible = true;

                Object x, y;
                x = FileName;
                y = visible;
                wApp.Visible = false;

                wDoc = wApp.Documents.Open(ref x, Visible:y);

                //wDoc = wApp.Documents.Open(ref x, ref Missing, ref  Missing, ref  Missing, ref  Missing,
                //    ref Missing, ref  Missing, ref  Missing, ref  Missing, ref  Missing, ref  Missing,
                //    ref y, ref Missing, ref Missing, ref Missing, ref Missing);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public Boolean CloseWord()
        {
            try
            {
                Boolean visible;
                Object y;
                visible = true;
                y = visible;
                
                wDoc.Close(ref y);
                //wDoc.Close(ref y, ref Missing, ref Missing);

                visible = false;
                y = visible;
                
                wApp.Quit(ref y);
                //wApp.Quit(ref y, ref Missing, ref Missing);

                Marshal.ReleaseComObject(wDoc);

                Marshal.ReleaseComObject(wApp);

                GC.Collect();

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Save Data
        /// </summary>
        /// <param name="SavePath"></param>
        /// <returns></returns>
        public Boolean Save(String SavePath)
        {
            try
            {
                Object FileName;
                FileName = SavePath;
                wDoc.SaveAs(FileName);
                //wDoc.SaveAs(ref FileName, ref  Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing);
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Action
        //public Boolean InsertRow(Int32 TableIndex, Int32 RowIndex)
        //{
        //    try
        //    {
        //        object temp = wDoc.Tables[TableIndex].Rows[RowIndex];
        //        wDoc.Tables[TableIndex].Rows.Add(ref temp);
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}
        //public Boolean DeletRow(Int32 TableIndex, Int32 RowIndex)
        //{
        //    try
        //    {
        //        wDoc.Tables[TableIndex].Rows[RowIndex].Delete();
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}
        #endregion

        #region DataTrans
        //public void ExportData(Int32 TableIndex, Int32 RowIndex, Int32 ColIndex, String Value)
        //{
        //    Word.Range r;
        //    Table wTable = wDoc.Tables[TableIndex];

        //    r = wTable.Cell(RowIndex, ColIndex).Range;

        //    r.Paragraphs[1].Range.Text = Value;

        //}
        //public void SetHeadData(Int32 TableIndex, Int32 RowIndex, Int32 ColIndex, String Value)
        //{
        //    HeaderFooter hr = wDoc.Sections.First.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
        //    hr.Range.Tables[TableIndex].Cell(RowIndex, ColIndex).Range.Text = Value;
        //}

        //public bool SetBookMarkValue(string bookMarkTagName, string bookMarkValue)
        //{
        //    Object oBookMarkName = bookMarkTagName;
        //    try
        //    {
        //        wDoc.Bookmarks.get_Item(ref oBookMarkName).Range.Select();
        //        wDoc.Bookmarks.get_Item(ref oBookMarkName).Range.Text = bookMarkValue;

        //    }
        //    catch (Exception err)
        //    {

        //        throw;
        //    }

        //    return true;

        //}

        public bool InsertTableToWord(SchmaMap map)
        {
            try
            {
                log.Info("Enter Word");
                Object oBookMarkName = "TableParagraph";
                object obreak =Word.WdBreakType.wdSectionBreakNextPage;
                object style = Word.WdBuiltinStyle.wdStyleHeading1;
                int count = map.Tables.Count;
                int index = 1;
                int tableIndex = 3;

                log.Info("Table Numbers="+map.Tables.Count);

                if (index == 1)
                {
                    foreach (TableMap t in map.Tables)
                    {
                        log.Info("Table =" + index);

                        object oRng = wDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                        Word.Paragraph oPara2 = wDoc.Content.Paragraphs.Add(ref oRng);
                        oPara2.Range.Text = index + "." + t.TableName;
                        oPara2.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                        oPara2.Format.SpaceAfter = 1;
                        oPara2.set_Style(ref style);
                        oPara2.Range.InsertBreak(ref obreak);//加br
                        oPara2.Range.InsertParagraphAfter();//insert  在斷落的後面
                        oPara2.Range.Select();//所有選擇的範圍

                        wDoc.Tables[2].Range.Copy();//取得第二個table  從一開始算

                        oPara2.Range.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);

                        //開始來實作資料欄位
                        //設定Table Title
                        Word.Table myTable = wDoc.Tables[tableIndex];

                        myTable.Cell(1, 2).Range.Text = t.TableName;
                        myTable.Cell(1, 4).Range.Text = "";
                        myTable.Cell(1, 6).Range.Text = "";

                        //加資料
                        object temp = myTable.Rows[3];
                        int rowIndex = 3;

                        //加Row
                        for (int i = 1; i < t.ColumnList.Count; i++)
                        {
                            myTable.Rows.Add(ref temp);

                        }//先把格式做出來，再放數值進去

                        //放入所有的值
                        foreach (ColumnMap cdoc in t.ColumnList)
                        {
                            //log.Info("Doc Column="+cdoc.Name);    
                            myTable.Rows[rowIndex].Cells[1].Range.Text = cdoc.ColumnID.ToString();
                            myTable.Rows[rowIndex].Cells[2].Range.Text = "";
                            myTable.Rows[rowIndex].Cells[3].Range.Text = "";
                            myTable.Rows[rowIndex].Cells[4].Range.Text = "";
                            myTable.Rows[rowIndex].Cells[5].Range.Text = cdoc.DataType;
                            myTable.Rows[rowIndex].Cells[6].Range.Text = cdoc.DataLength.ToString();
                            myTable.Rows[rowIndex].Cells[7].Range.Text = cdoc.NullAble;
                            myTable.Rows[rowIndex].Cells[8].Range.Text = cdoc.ColumnName;
                            myTable.Rows[rowIndex].Cells[9].Range.Text = cdoc.Comments;
                            myTable.Rows[rowIndex].Cells[10].Range.Text = "";

                            rowIndex++;

                        }



                        index++;
                        tableIndex++;

                    }
                }

                log.Info("Add Directory");

                Object oDirBookMarkName = "DirBookMark";
                Word.Range oRngDir = wDoc.Bookmarks.get_Item(ref oDirBookMarkName).Range;
                object HeadingLevel = 1;
                object useLink = true;
                wDoc.TablesOfContents.Add(oRngDir,
                ref missing, ref HeadingLevel, ref HeadingLevel,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref useLink, ref missing, ref missing);//產生目錄的方法

                //刪除Template Table

                log.Info("Delete Temp Word");
                wDoc.Tables[2].Delete();
                
            }
            catch (Exception err)
            {
                log.Info("Exception=" + err.Message + err.StackTrace);
                throw;
            }

            log.Info("Insert Word OK");

            return true;

        }

        #endregion
    }
}
