using System;
using System.IO;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using NPOI.HSSF.UserModel;
using TdqqClient.Commands;
using TdqqClient.Services.Common;

namespace TdqqClient.ViewModels
{
    public class CbfbmViewModel:NotificationObject
    {
        #region 关闭窗口命令

        public DelegateCommand CloseCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
        }
        #endregion

        #region 绑定属性

        private string _fbfbm;

        public string Fbfbm
        {
            get { return _fbfbm; }
            set
            {
                _fbfbm = value;
                this.RaisePropertyChanged("Fbfbm");
            }
        }

        private int _startIndex;

        public int StartIndex
        {
            get { return _startIndex; }
            set
            {
                _startIndex = value;
                this.RaisePropertyChanged("StartIndex");
            }
        }
        
        #endregion

        #region 打开按钮命令

        public DelegateCommand ConfirmCommand { get; set; }

        private void StrartCbfbm(object parameter)
        {
            if (string.IsNullOrEmpty(Fbfbm) || Fbfbm.Length != 14) return;
            if (StartIndex <= 0) return;
            DialogHelper dialogHelper=new DialogHelper("xls");
            var openFilePath = dialogHelper.OpenFile("选择家庭成员信息表");
            if (string.IsNullOrEmpty(openFilePath)) return;
            if (StrartCbfbm(openFilePath))
            {
                MessageBox.Show(null, "编码成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "编码失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool StrartCbfbm(string openFilePath)
        {
            try
            {
                using (FileStream fileStream = new FileStream(openFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    //FileStream fileStream = new FileStream(fileSource, FileMode.Open, FileAccess.ReadWrite);
                    HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                    string homeName = string.Empty;
                    for (int i = 1; i <= sheetSource.LastRowNum; i++)
                    {
                        HSSFRow rowSource = (HSSFRow)sheetSource.GetRow(i);//获取一行数据
                        if (rowSource == null) break;
                        if (rowSource.GetCell(0) == null)
                        {
                            break;
                        }
                        if (homeName == rowSource.GetCell(0).ToString().Trim())
                        {
                            rowSource.GetCell(0).SetCellValue(IntToString(StartIndex));
                        }
                        else
                        {
                            StartIndex++;
                            homeName = rowSource.GetCell(0).ToString().Trim();
                            rowSource.GetCell(0).SetCellValue(IntToString(StartIndex));
                        }
                    }
                    FileStream fs = new FileStream(openFilePath, FileMode.Create, FileAccess.Write);

                    workbookSource.Write(fs);
                    fs.Close();
                    fileStream.Close();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            
        }
        private string IntToString(int ordervalue)
        {
            return Fbfbm + ordervalue.ToString("0000");
        }

        #endregion

        public CbfbmViewModel()
        {
            StartIndex = 1;
            CloseCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            ConfirmCommand.ExecuteAction = new Action<object>(StrartCbfbm);
        }
    }
}
