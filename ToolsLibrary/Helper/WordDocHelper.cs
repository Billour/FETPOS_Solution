using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ToolsLibrary.Entity;

namespace ToolsLibrary.Helper
{
    public class WordDocHelper
    {
        private string _TargetFileSavePath = "";

        WordCore word = new WordCore();

        public WordDocHelper(string templateFilePath, string targetFolder, string targetFileSavePath)
        {

            _TargetFileSavePath = targetFileSavePath;

            word.Open(templateFilePath);
        }

        public  bool Save()
        {
            return word.Save(_TargetFileSavePath);
        }

        public  void Close()
        {
            
            word.CloseWord();
        }

        public void InsertSchema(SchmaMap map)
        {
            word.InsertTableToWord(map);
        }
    }
}
