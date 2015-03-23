using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Aspose.Pdf.Facades;
using TdqqClient.Commands;
using TdqqClient.Services.Common;
using TdqqClient.Services.Database;
using TdqqClient.Services.Export;
using TdqqClient.Services.Export.ExportSingle;
using TdqqClient.Models;
using TdqqClient.Views;

namespace TdqqClient.ViewModels
{
    public class ExportViewModel:NotificationObject
    {
        #region 关闭窗口

        public DelegateCommand CloseCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
        }
        #endregion


        #region 绑定属性
        private List<FarmerModel> _farmerList;

        public List<FarmerModel> FarmerList
        {
            get { return _farmerList; }
            set
            {
                _farmerList = value;
                this.RaisePropertyChanged("FarmerList");
            }
        }

        private FarmerModel _selectFarmer;

        public FarmerModel SelectFarmer
        {
            get { return _selectFarmer; }
            set
            {
                _selectFarmer = value;
                this.RaisePropertyChanged("SelectFarmer");
            }
        }
        #endregion

        #region 输出成果命令

        public DelegateCommand ExportArchiveCommond { get; set; }
        public DelegateCommand ExportCertificationCommand { get; set; }

        private void ExportArchive(object parameter)
        {
            if (SelectFarmer == null) return;
            SelectFeatureViewModel selectFeatureVm=new SelectFeatureViewModel(_personDatabase);
            selectFeatureVm.Caption = "请选择村界要素类";
            SelectFeatureWindow selectFeatureV=new SelectFeatureWindow(selectFeatureVm);
            selectFeatureV.ShowDialog();
            if (string.IsNullOrEmpty(selectFeatureVm.SelectFeature)) return;
            DialogHelper dialogHelper=new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (ExportArchive(selectFeatureVm.SelectFeature, folderPath))
            {
                MessageBox.Show(null, "归档成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "归档失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ExportArchive(string cunEdge, string folderPath)
        {
            try
            {
                 
                 ExportBase export=new DkExport(_personDatabase,_selectFeature,_basicDatabase);
                 export.Export(SelectFarmer.Cbfmc,SelectFarmer.Cbfbm,folderPath);
                 export = new CoverExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new CbfExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new ContractExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new StatementExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new MapExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath,cunEdge);
                 export = new AcceptExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new GhbExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 export = new RegisterExport(_personDatabase, _selectFeature, _basicDatabase);
                 export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                 ConnectPdfFile(folderPath);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
                      
        }

        private void ConnectPdfFile(string folderPath)
        {
            var files = GetToMergeFileFullName(folderPath, SelectFarmer.Cbfbm.Substring(14), SelectFarmer.Cbfmc);
            List<Stream> fileStreams=new List<Stream>();
            foreach (var pdfFile in files)
            {
                fileStreams.Add(new FileStream(pdfFile,FileMode.Open));
            }
            var targetFile = folderPath + @"\" + SelectFarmer.Cbfbm.Substring(14) + "_" + SelectFarmer.Cbfmc +
                             "档案局资料.pdf";
            var outFileStream = new FileStream(targetFile, FileMode.Create);
            var pdfEditor = new PdfFileEditor();
            pdfEditor.Concatenate(fileStreams.ToArray(), outFileStream);
            outFileStream.Close();
            foreach (var fileStream in fileStreams)
            {
                fileStream.Close();
            }
            foreach (var file in files)
            {
                File.Delete(file);
            }
        }

        private IEnumerable<string> GetToMergeFileFullName(string folderPath, string shortCbfbm, string cbfmc)
        {
            string fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_00档案局目录.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_01承包方调查表.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_02承包地块调查表.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_03承包经营权合同书.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_04地块示意图.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_05户主声明书.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_06公示无异议声明书.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_07公示结果归户表.pdf";
            if (File.Exists(fileName)) yield return fileName;
            fileName = folderPath + @"\" + shortCbfbm + "_" + cbfmc + "_08承包经营权证登记簿.pdf";
            if (File.Exists(fileName)) yield return fileName;
        }

        private void ExportCertification(object parameter)
        {
            if (SelectFarmer == null) return;
            DialogHelper dialogHelper=new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath))return;
            if (ExportCertification(folderPath))
            {
                MessageBox.Show(null, "经营权证导出成功", 
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "经营权证导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private bool ExportCertification(string folderPath)
        {
            try
            {
                var export = new TdqqClient.Services.Export.ExportSingle.
                    JyqzExport(_personDatabase, _selectFeature, _basicDatabase);
                export.Export(SelectFarmer.Cbfmc, SelectFarmer.Cbfbm, folderPath);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            
        }
        #endregion

        private string _personDatabase;
        private string _selectFeature;
        private string _basicDatabase;

        private void InitGridInfo()
        {
            this.FarmerList=new List<FarmerModel>();
            var dt = SelectCbfbmOwnFields();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                FarmerList.Add(new FarmerModel()
                {
                    Cbfmc = dt.Rows[i][1].ToString().Trim(),
                    Cbfbm = dt.Rows[i][0].ToString().Trim()
                });
            }

        }
        private  System.Data.DataTable SelectCbfbmOwnFields()
        {
            var sqlString = string.Format("Select distinct CBFBM,CBFMC From {0} where CBFBM NOT LIKE  '{1}' order by CBFBM ",
            _selectFeature, "99999999999999%");
            var accessFactory = new MsAccessDatabase(_personDatabase);
            return accessFactory.Query(sqlString);
        }

        #region 初始化命令

        private void InitCommand()
        {
            this.CloseCommand=new DelegateCommand();
            this.ExportArchiveCommond=new DelegateCommand();
            this.ExportCertificationCommand=new DelegateCommand();
            CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            ExportArchiveCommond.ExecuteAction=new Action<object>(ExportArchive);
            ExportCertificationCommand.ExecuteAction=new Action<object>(ExportCertification);
        }
        #endregion

        public ExportViewModel(string personDatabase, string selectFeature, string basicDatabase)
        {
            _personDatabase = personDatabase;
            _selectFeature = selectFeature;
            _basicDatabase = basicDatabase;
            InitGridInfo();
            InitCommand();
        }

    }
}
