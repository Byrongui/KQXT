using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Forms = System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Data;

namespace Bozoneit.CQXI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filePath="";
        private string fileFullPath = "";
        private string fileName="";
        private string sheetName = "";
        private string path = "";
        private string path0 = "";//生成文件路径

        private List<RecordMode> RecordList = new List<RecordMode>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Btn_selectFile_Click(object sender, RoutedEventArgs e)
        {
            Forms.OpenFileDialog openFile = new Forms.OpenFileDialog();
            openFile.Filter = "2007以上版本|*.xlsx|2017以前版本|*.xls";
            if (openFile.ShowDialog() == Forms.DialogResult.OK)
            {
                fileFullPath = openFile.FileName ;
                fileName = openFile.SafeFileName;
                filePath = fileFullPath.Remove(fileFullPath.Length - fileName.Length, fileName.Length);
                TB_selectFilePath.Text = fileFullPath;
                if (TB_selectPath.Text.Trim() == "")
                {
                    TB_selectPath.Text = filePath;
                }
            }
        }

        private void Btn_selectPath_Click(object sender, RoutedEventArgs e)
        {
            Forms.FolderBrowserDialog folderBrowserDialog = new Forms.FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                // 对路径，进行校验
                TB_selectPath.Text = folderBrowserDialog.SelectedPath;
            }
        }
        
        private void Btn_Start_Click(object sender, RoutedEventArgs e)
        {
            if (fileFullPath.Trim() == "")
            {
                MessageBox.Show("咋不得先选个文件啊，先点上面红色的按钮,亲！");
                return;
            }
          
            Btn_Start.IsEnabled = false;
            Btn_Start.Content = "正在处理中...";
            path = TB_selectPath.Text;
            Task.Factory.StartNew(() => {
                showLog("$clear");
                showLog("开始处理");
                RecordList.Clear();
                pross();
                //Thread.Sleep(5000);
                Btn_Start.Dispatcher.Invoke(new Action(() => {
                    Btn_Start.IsEnabled = true;
                    Btn_Start.Content = "开始处理";
                    showLog("处理完成");
                    MessageBox.Show(string .Format ("处理完成，请查看文件:{0}",path0));
                }));

            });
        }

        #region 日志显示
        private void showLog(string messes)
        {
            this.RTB_log.Dispatcher.Invoke(new Action<string>((m) => {
                if (messes == "$clear")
                {
                    this.RTB_log.Document.Blocks.Clear();
                }
                else
                {
                    this.RTB_log.AppendText(m);
                    this.RTB_log.AppendText("\r\n");
                }
            }),messes);
        }
        #endregion

        #region 处理过程
        private void pross()
        {
            //判断临时目录是否存在
            if(!Directory.Exists("Temp"))
            {
                Directory.CreateDirectory("Temp");
            }
            //考备文件到临时目录
            string desFilePath
                = System.IO.Path.Combine("Temp", fileName);

            File.Copy(fileFullPath, desFilePath, true);

            //读取数据源
            DataTable dt = NpoiExcelHelper.ExcelToDataTable(desFilePath, ref sheetName, true);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //for(int j = 0; j < dt.Columns.Count; j++)

                RecordMode mode = new RecordMode();
                mode.name = dt.Rows[i][0].ToString();
                mode.date = dt.Rows[i][1].ToString();
                mode.startTime = dt.Rows[i][2].ToString();
                mode.endTime = dt.Rows[i][3].ToString();
                mode.qdTime = dt.Rows[i][4].ToString();
                mode.qtTime = dt.Rows[i][5].ToString();
                mode.bm = dt.Rows[i][6].ToString();
                mode.bz = dt.Rows[i][7].ToString();

                RecordList.Add(mode);
                if(mode.isYC)
                {
                    showLog(string.Format("异常：姓名：{0}，日期：{1},签到时间：{2}， 签退时间:{3}",mode.name, 
                        mode.date, mode.qdTime, mode .qtTime));
                }
            }

            //生成新数据
             path0 = System.IO.Path.Combine(path, "BZ_"+fileName);
            NpoiExcelHelper.ListToExcel(path0, desFilePath, RecordList, "", sheetName);

            File.Delete(desFilePath);

        }
        #endregion
    }
}
