
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Aspose.Pdf.Facades;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Export.ExportTotal
{
    /// <summary>
    /// 导出按照农户所有
    /// </summary>
    class ArchivesExport:ExportBase
    {
        public ArchivesExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public void Export()
        {
            var archiveVm=new ArchiveViewModel();
            ArchiveView archiveV=new ArchiveView(archiveVm);
            archiveV.ShowDialog();
            if (string.IsNullOrEmpty(archiveVm.ArchiveFolder)) return;
            if (Export(archiveVm))
            {
                MessageBox.Show(null, "归档成功", 
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "归档失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(ArchiveViewModel archiveVm)
        {
            Wait wait=new Wait();
            wait.SetWaitCaption("归档中");
            Hashtable para=new Hashtable()
            {
                {"wait",wait},{"archiveVm",archiveVm},{"ret",false}
            };
            Thread t=new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }

        private void ExportF(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            var archiveVm = para["archiveVm"] as ArchiveViewModel;
            try
            {
                var dt = SelectCbfbmOwnFields();
                int rowCount = dt.Rows.Count;
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress((double)i/(double)rowCount);
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    var cbfbm = dt.Rows[i][0].ToString().Trim();
                    MergePdfFiles(cbfbm,cbfmc,archiveVm);
                }
                para["ret"] = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }

        }

        private void MergePdfFiles(string cbfbm, string cbfmc, ArchiveViewModel archiveVm)
        {
            var createFilePath = archiveVm.ArchiveFolder + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "档案局资料.pdf";
            var outFileStream = new FileStream(createFilePath, FileMode.Create);
            var inFileStream = GetToMergePdfFile(cbfbm, archiveVm);
            var pdfEditor = new PdfFileEditor();
            pdfEditor.Concatenate(inFileStream.ToArray(), outFileStream);
            outFileStream.Close();
            foreach (var stream in inFileStream)
            {
                stream.Close();
            }

        }

        private IEnumerable<Stream> GetToMergePdfFile(string cbfbm, ArchiveViewModel archiveVm)
        {
            List<Stream> mergeFileStrem=new List<Stream>();
            string filefullName;
            if (!string.IsNullOrEmpty(archiveVm.CoverFolder))
            {
                filefullName = GetExistFilePath(archiveVm.CoverFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName,FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.CbfFolder))
            {
                filefullName = GetExistFilePath(archiveVm.CbfFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.DkFolder))
            {
                filefullName = GetExistFilePath(archiveVm.DkFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.ContractFolder))
            {
                filefullName = GetExistFilePath(archiveVm.ContractFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.MapFolder))
            {
                filefullName = GetExistFilePath(archiveVm.MapFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.StatementFolder))
            {
                filefullName = GetExistFilePath(archiveVm.StatementFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.AcceptFolder))
            {
                filefullName = GetExistFilePath(archiveVm.AcceptFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.GhbFolder))
            {
                filefullName = GetExistFilePath(archiveVm.GhbFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            if (!string.IsNullOrEmpty(archiveVm.RegisterFolder))
            {
                filefullName = GetExistFilePath(archiveVm.RegisterFolder, cbfbm);
                mergeFileStrem.Add(new FileStream(filefullName, FileMode.Open));
            }
            return mergeFileStrem;
        }

        private string GetExistFilePath(string folderPath, string cbfbm)
        {
            var dir=new DirectoryInfo(folderPath);
            foreach (var file in dir.GetFiles("*.pdf"))
            {
                var filename = file.Name;
                if (filename.StartsWith(cbfbm.Substring(14)))
                {
                    return file.FullName;
                }
            }
            return string.Empty;
        }
    }
}
