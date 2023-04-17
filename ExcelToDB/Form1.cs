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
        /// ѡ��Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "��ѡ��excel";
            dialog.Filter = "*.xls | *.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = dialog.FileName;
                this.label1.Text = $"path:{fileName}";
            }
        }
        /// <summary>
        /// ����sql
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (fileName == string.Empty)
            {
                MessageBox.Show("��ѡ���ļ�", "ϵͳ��ʾ");
                return;
            }
            string tableName = this.textBox1.Text.ToString().Trim();
            string tableNameDesc = this.textBox2.Text.ToString().Trim();
            if (tableName == string.Empty)
            {
                MessageBox.Show("���������", "ϵͳ��ʾ");
                return;
            }
            if (tableNameDesc == string.Empty)
            {
                MessageBox.Show("�����������", "ϵͳ��ʾ");
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

                if ("��".Equals(isKey))
                {
                    sqlItem += $"primary key ";
                }
                if ("��".Equals(isEmpty))
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
                #region ��ӱ��ֶ�ע��
                sbDesc.Append($"execute sp_addextendedproperty 'MS_Description','{desc}','user','dbo','table','{tableName}','column','{fielName}'; \r\n ");
                #endregion

            }
            sb.Append(")");
            #region ���ɱ�ע��
            string tableDescSQl = $"execute sp_addextendedproperty 'MS_Description','{tableNameDesc}','user','dbo','table','{tableName}',null,null;";
            #endregion
            this.SqlContentBox.Text = sb.ToString() + "\r\n" + sbDesc.ToString() + "\r\n" + tableDescSQl;
        }

        /// <summary>
        /// ���
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
        /// ����ģ��
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
                
                folderBrowser.Description = "��ѡ�񱣴�ĵ�ַ";
                //folderBrowser.ShowNewFolderButton = true;

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    string filePath = folderBrowser.SelectedPath;
                    filePath = Path.Combine(filePath, "ģ��.xlsx");
                    fs = File.Create(filePath);
                    //string templeFile = Application.CommonAppDataPath + @"\assist\ģ��.xlsx";
                    string templeFile = Application.StartupPath + @"\assist\ģ��.xlsx";
                    fs2 = File.OpenRead(templeFile);
                    fs2.CopyTo(fs);
                    OpenFolderAndSelectedFile(filePath);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "ϵͳ�쳣");
                fs?.Close();
                fs2?.Close();
            }
           

        }


        /// <summary>
        /// ��Ŀ¼��ѡ���ļ�
        /// </summary>
        /// <param name="filePathAndName">�ļ���·�������ƣ����磺C:\Users\Administrator\test.txt��</param>
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