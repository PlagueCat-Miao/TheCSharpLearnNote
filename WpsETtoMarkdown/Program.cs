using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Xml;


namespace WpsETtoMarkdown
{
    class MyRowData 
    {
        public string KeyName;
        public string KeyNameZCH;
        public string necessary; // 必填
        public string type;//节点数据类型 
    }

    class Program
    {
        /// <summary>
        /// 处理一个et文件
        /// </summary>
        /// <param name="fileName"></param>
        static void parseOneEt(string fileName)
        {
            Excel.Application appli = new Excel.Application();
            try
            {
               
                string markdownFileName = Path.GetDirectoryName(fileName) + "\\ans.md";
                string[] sheetsName = new string[2] { "sheet1", "sheet2" };

                //读取Excel
                Excel._Workbook wk = appli.Workbooks.Open(fileName);
                var sheetsDic = getSheetNum(wk, sheetsName);
                var dataset = getDataSet(wk, sheetsDic);

                //Show(dataset, sheetsName);
                Console.WriteLine(fileName + ":" + ETToMarkdown(dataset, markdownFileName, sheetsName));
                appli.Workbooks.Close();
                appli.Quit();
            }
            catch (Exception ex)
            {
                appli.Workbooks.Close();
                appli.Quit();
                Console.WriteLine(fileName + "  EX:" + ex.ToString());
                return;
            }
             

        }
        /// <summary>
        /// 读取文件夹内所有et
        /// </summary>
        /// <param name="FilePath"></param>
        /// <returns></returns>
        static private string[] getAllFile(string FilePath)
        {
            return Directory.GetFiles(FilePath, "*.et");
        }
        /// <summary>
        /// XML指定要解析的et
        /// </summary>
        /// <param name="xmlFileName"></param>
        /// <returns></returns>
        static private List<string> getFileNameListFromXML(string xmlFileName)
        {
            List<string> fileNameList = new List<string>();
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName);
            XmlNode root = doc.SelectSingleNode("fileList");
            // 得到根节点的所有子节点
            XmlNodeList xnl = root.ChildNodes;
            foreach (XmlNode node in xnl) 
            {
                XmlElement xe = (XmlElement)node;
                fileNameList.Add(Path.GetDirectoryName(xmlFileName) +"\\" +xe.GetAttribute("name").ToString());
            }
            return fileNameList;
        }
        /// <summary>
        ///  获取指定名字的sheet号
        /// </summary>
        /// <param name="wk">Excel._Workbook</param>
        /// <param name="sheetsName">所需表名</param>
        /// <returns></returns>
        static private Dictionary<string, int> getSheetNum(Excel._Workbook wk, string[] sheetsName)
        {
            Dictionary<string, int> ans = new Dictionary<string, int>();
            foreach (string name in sheetsName)
            {
                ans.Add(name, -1);
            }

            //读取sheet
            int sheetsCount = wk.Worksheets.Count;//获取sheet数量
            //遍历每一个sheet页面
            for (int k = 1; k <= sheetsCount; k++)
            {
                Excel.Worksheet sheet = wk.Worksheets.get_Item(k);
                
                if (ans.ContainsKey(sheet.Name))
                {
                    ans[sheet.Name] = k;
                }
            }
            return ans;
        }
        /// <summary>
        /// 根据表号，扫表格。
        /// </summary>
        /// <param name="wk">Excel._Workbook</param>
        /// <param name="sheetsDic"></param>
        /// <returns></returns>
        static private DataSet getDataSet(Excel._Workbook wk, Dictionary<string, int> sheetsDic)
        {
            DataSet dataSet = new DataSet();
            foreach (var kv in sheetsDic) 
            {
                DataTable dt = new DataTable(kv.Key);
                if (kv.Value == -1) 
                {
                    Console.WriteLine("sheet:{0} not found", kv.Key);
                    continue;
                }
                Excel.Worksheet sheet = wk.Worksheets.get_Item(kv.Value);

                Excel.Range range = sheet.UsedRange;
                int rowCount = range.Rows.Count;//获取行数
                int columCount = range.Columns.Count;//获取列数

                //设置列头
                for (int j = 1; j <= columCount; j++)
                {
                    dt.Columns.Add(((Excel.Range)range.get_Item(1, j)).Text);
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow datarow = dt.NewRow();
                    for (int j = 1; j <= columCount; j++)
                    {
                        string title = ((Excel.Range)range.get_Item(j)).Text;
                        datarow[title] = range.get_Item(i, j).Text;
                    }
                    dt.Rows.Add(datarow);
                }
                dataSet.Tables.Add(dt);
            }

            return dataSet;
        }
        /// <summary>
        /// 打印多个sheet表格
        /// </summary>
        /// <param name="dataSet"></param>
        /// <param name="sheetsName"></param>
        static private void Show(DataSet dataSet,string[] sheetsName) 
        {
           
            foreach (string name in sheetsName)
            {
                if (!dataSet.Tables.Contains(name)) 
                {
                    Console.WriteLine("{0} not found", name);
                    continue;
                }
                DataTable dt = dataSet.Tables[name];
                
                Console.WriteLine(dt.ToString());
                foreach (DataRow myRow in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        Console.Write(myRow[i]+" "); ;
                    }
                    Console.Write("\n");
                }
            }
        }
        /// <summary>
        /// 生成/添加markdown
        /// </summary>
        /// <param name="dataSet"></param>
        /// <param name="markdownFileName"> 生成文件的名字</param>
        /// <param name="sheetsName">必要sheet</param>
        /// <returns></returns>
        static private bool ETToMarkdown(DataSet dataSet,string markdownFileName, string[] sheetsName)
        {
            // 表格数据检查
            foreach (string sheet in sheetsName) 
            {
                if (!dataSet.Tables.Contains(sheet))
                {
                    Console.WriteLine("sheet数据获取失败:"+ sheet);
                    return false;
                }

            }

            var info = dataSet.Tables["sheet1"];
            string title = info.Rows[0]["主键"].ToString();
            var data = dataSet.Tables["sheet2"];

            List<MyRowData> inputKey = new List<MyRowData>();
            List<MyRowData> outputKey = new List<MyRowData>();

            foreach (DataRow row in data.Rows) 
            {
                if (row["接口类型"].ToString().Contains("输入"))
                { 
                    inputKey.Add(new MyRowData(){
                        KeyName = Regex.Replace(row["英文名称"].ToString(), @"[^a-zA-Z0-9\u4e00-\u9fa5\s]", ""),
                        KeyNameZCH = row["中文描述"].ToString(),
                        necessary = row["必填"].ToString().Contains("非")?"否":"是",
                        type = row["数据类型"].ToString().Equals("1") ? row["是否为数组"].ToString() : "数组",
                    });
                }
                else 
                {
                    Console.WriteLine("未定义行");
                }

            }

            using (StreamWriter sw = new StreamWriter(markdownFileName, true))
            {
                sw.WriteLine("## "+title);

                sw.WriteLine("输入 :\\");
                sw.WriteLine("| 字段名称（英文）|字段名称（中文） | 是否必输项 |  字段类型 | 取值说明 |");
                sw.WriteLine("|  ---- | ----  |  ----  |   ----  | ----  |");
                foreach (var row in inputKey)
                {
                    sw.WriteLine("| {0} | {1} | {2} | {3} | {4} |" , row.KeyName, row.KeyNameZCH, row.necessary,row.type," ");
                }
                sw.WriteLine("\n\n\n");
                sw.WriteLine("输出 :\\");
                sw.WriteLine("| 字段名称（英文）|字段名称（中文） | 是否必输项 |  字段类型 | 取值说明 |");
                sw.WriteLine("|  ---- | ----  |  ----  |   ----  | ----  |");

                foreach (var row in inputKey)
                {
                    sw.WriteLine("| {0} | {1} | {2} | {3} | {4} |", row.KeyName, row.KeyNameZCH, row.necessary, row.type, " ");
                }
                sw.WriteLine("\n\n\n");
                sw.WriteLine("请求示例 :\\");
                sw.WriteLine("\n\n\n");
            }
            return true;
        }

        static void Main(string[] args)
        {
            string xmlFileName = @"C:\。。。\Desktop";
            var list = getAllFile(xmlFileName);
            foreach (string file in list)
            {
                parseOneEt(file);

            }
            Console.WriteLine("done喵");
        }
    }
}
