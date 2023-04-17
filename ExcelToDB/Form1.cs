using System.Data;
using System.Diagnostics;
using System.Text;

namespace ExcelToDB
{
    public partial class Form1 : Form
    {

        private string fileName = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 选择Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "请选择excel";
            dialog.Filter = "*.xls | *.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = dialog.FileName;
                this.label1.Text = $"path:{fileName}";
            }
        }
        /// <summary>
        /// 生成sql
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (fileName == string.Empty)
            {
                MessageBox.Show("请选择文件", "系统提示");
                return;
            }
            string tableName = this.textBox1.Text.ToString().Trim();
            string tableNameDesc = this.textBox2.Text.ToString().Trim();
            if (tableName == string.Empty)
            {
                MessageBox.Show("请输入表名", "系统提示");
                return;
            }
            if (tableNameDesc == string.Empty)
            {
                MessageBox.Show("请输入表描述", "系统提示");
                return;
            }
            FileStream fs = File.OpenRead(fileName);
            DataTable dt = ExcelHelper.ExcelToTable(fileName, fs, true, null, 0, 8);
            Console.WriteLine(dt);
            StringBuilder sb = new StringBuilder();
            StringBuilder sbDesc = new StringBuilder();
            sb.Append($"CREATE TABLE {tableName} (\r\n");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                string fielName = dr[1].ToString();
                string dbType = dr[2].ToString();
                string size = dr[3].ToString();
                string isEmpty = dr[4].ToString();
                string isKey = dr[5].ToString();
                string desc = dr[6].ToString();
                string sqlItem = string.Empty;
                if (string.IsNullOrEmpty(size.Trim()))
                {
                    sqlItem = $"[{fielName}] {dbType} ";
                }
                else
                {
                    sqlItem = $"[{fielName}] {dbType}({size}) ";
                }

                if ("是".Equals(isKey))
                {
                    sqlItem += $"primary key ";
                }
                if ("否".Equals(isEmpty))
                {
                    sqlItem += $"not null ";
                }
                //if(!String.IsNullOrEmpty(desc))
                //{
                //    sqlItem += $"comment '{desc}' ";
                //}
                if (i < dt.Rows.Count - 1)
                {
                    sqlItem += ",";
                }
                sb.Append(sqlItem + "\r\n");
                #region 添加表字段注释
                sbDesc.Append($"execute sp_addextendedproperty 'MS_Description','{desc}','user','dbo','table','{tableName}','column','{fielName}'; \r\n ");
                #endregion

            }
            sb.Append(")");
            #region 生成表注释
            string tableDescSQl = $"execute sp_addextendedproperty 'MS_Description','{tableNameDesc}','user','dbo','table','{tableName}',null,null;";
            #endregion
            this.SqlContentBox.Text = sb.ToString() + "\r\n" + sbDesc.ToString() + "\r\n" + tableDescSQl;
        }

        /// <summary>
        /// 清空
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            this.SqlContentBox.Text = "";
            this.fileName = "";
            this.textBox1.Text = "";
            this.label1.Text = "";
        }

        /// <summary>
        /// 下载模板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            FileStream fs = null;
            FileStream fs2 = null;
            try
            {
                FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
                
                folderBrowser.Description = "请选择保存的地址";
                //folderBrowser.ShowNewFolderButton = true;

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    string filePath = folderBrowser.SelectedPath;
                    filePath = Path.Combine(filePath, "模板.xlsx");
                    fs = File.Create(filePath);
                    //string templeFile = Application.CommonAppDataPath + @"\assist\模板.xlsx";
                    string templeFile = Application.StartupPath + @"\assist\模板.xlsx";
                    fs2 = File.OpenRead(templeFile);
                    fs2.CopyTo(fs);
                    OpenFolderAndSelectedFile(filePath);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "系统异常");
                fs?.Close();
                fs2?.Close();
            }
           

        }


        /// <summary>
        /// 打开目录且选中文件
        /// </summary>
        /// <param name="filePathAndName">文件的路径和名称（比如：C:\Users\Administrator\test.txt）</param>
        private static void OpenFolderAndSelectedFile(string filePathAndName)
        {
            if (string.IsNullOrEmpty(filePathAndName)) return;

            Process process = new Process();
            ProcessStartInfo psi = new ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + filePathAndName;
            process.StartInfo = psi;

            //process.StartInfo.UseShellExecute = true;
            try
            {
                process.Start();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                process?.Close();

            }
        }
    }
}