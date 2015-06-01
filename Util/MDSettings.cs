using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn4Pdf.Util
{
    class MDSettings
    {
        #region 属性
        //单实例
        private static MDSettings m_instance = null;
        //配置文件路径
        private readonly string FilePath = System.AppDomain.CurrentDomain.BaseDirectory + "\\Config\\Company";
        //配置文件后缀名
        private readonly string FileExt = "*.MD";
        /// <summary>配置文件信息</summary>
        private Dictionary<string, Dictionary<string, string>> dicconfig = new Dictionary<string, Dictionary<string, string>>();

        public Dictionary<string, Dictionary<string, string>> DicConfig
        {
            get { return dicconfig; }
            private set { dicconfig = value; }
        } 
        #endregion

        #region 私有方法
        /// <summary>读取配置文件夹</summary>
        private void Initialize()
        {
            dicconfig.Clear();
            string strPath = FilePath;
            if (!Directory.Exists(strPath)) return;
            DirectoryInfo folder = new DirectoryInfo(strPath);
            foreach (FileInfo file in folder.GetFiles(FileExt))
            {
                dicconfig.Add(file.Name.TrimEnd(FileExt.ToCharArray()), new Dictionary<string, string>());
                using (StreamReader sr = File.OpenText(file.FullName))
                {
                    while (!sr.EndOfStream)
                        SetDictionary(sr.ReadLine(), dicconfig[file.Name.TrimEnd(FileExt.ToCharArray())]);
                }
            }
        }

        /// <summary>读取数据写入字典</summary>
        /// <param name="strLine">数据行</param>
        /// <remarks>输入参数左右以"="分割,右值为多个时以"|"分割/remarks>
        /// <param name="dic">要设置的字典</param>
        private void SetDictionary(string strLine, Dictionary<string, string> dic)
        {
            if (!strLine.Contains("=")) return;
            int i = strLine.IndexOf('=');
            dic.Add(strLine.Substring(0, i).Trim().ToUpper(), strLine.Substring(i + 1).Trim());
        }

        #endregion

        #region 构造函数
        /// <summary>构造函数</summary>
        /// <remarks>读取配置文件初始化字典</remarks>
        private MDSettings()
        {
            Initialize();
        }

        public static MDSettings getInstance()
        {
            if (m_instance == null)
            {
                m_instance = new MDSettings();
            }
            return m_instance;
        } 
        #endregion

        #region 公有方法

        /// <summary>保存单个MD文件</summary>
        /// <param name="key">船公司</param>
        public void Save(string key)
        {
            if (!dicconfig.ContainsKey(key)) return;
            string path = FilePath + "\\" + key+".MD";
            using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                foreach (string deHB in dicconfig[key].Keys)
                    sw.WriteLine(string.Format("{0}={1}", deHB, dicconfig[key][deHB]));
                sw.Flush();
            }
        }

        /// <summary>保存所有MD文件</summary>
        public void Save()
        {
            foreach (string mdkey in dicconfig.Keys)
            {
                Save(mdkey);
            }
        }

        /// <summary>设置键值(不存在时新增)</summary>
        /// <param name="Mdkey">文件名</param>
        /// <param name="k">索引</param>
        /// <param name="v">值</param>
        public void SetKeyValue(string Mdkey, string k, string v)
        {
            if (!dicconfig.ContainsKey(Mdkey)) return;
            if (dicconfig[Mdkey].ContainsKey(k))
                dicconfig[Mdkey][k] = v;
            else
                dicconfig[Mdkey].Add(k, v);
            Save(Mdkey);
        } 
        #endregion
    }
}
