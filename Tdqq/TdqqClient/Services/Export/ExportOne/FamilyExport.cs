using System;
using System.Collections;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Services.Export.ExportOne
{
    class FamilyExport:ExportBase,IExport
    {
        public FamilyExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("xls");
            var saveFilePath = dialogHelper.SaveFile("导出家庭成员信息表");
            if (string.IsNullOrEmpty(saveFilePath)) return;
            if (Export(saveFilePath))
            {
                MessageBox.Show(null, "导出家庭成员信息表成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "导出家庭成员信息表失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出家庭成员信息表");
            var para = new Hashtable()
            {
                {"wait",wait},{"saveFilePath",saveFilePath},{"ret",false}
            };
            var t = new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void ExportF(object p)
        {
            var para = p as Hashtable;
            var savedFilePath = para["saveFilePath"].ToString();
            var wait = para["wait"] as Wait;
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\家庭成员信息表.xls";
            File.Copy(templatePath, savedFilePath, true);
            var cbfs = Cbfs(false);
            int startRow = 3, endRow = 3;
            using (var fileStream = new System.IO.FileStream(savedFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbookSource = new HSSFWorkbook(fileStream);
                ICellStyle style = MergetStyle(workbookSource);
                var rowCount =cbfs.Count;
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress((double)i / (double)rowCount);
                    var cbfbm = cbfs[i].Cbfbm;
                    int familyCount;
                    FillOneFamily(workbookSource, cbfbm, ref endRow, out familyCount);
                    //合并单元格
                    ISheet sheet = workbookSource.GetSheetAt(0);
                    IRow row = sheet.GetRow(startRow);
                    //序号
                    sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));
                    ICell cell = row.GetCell(0);
                    cell.CellStyle = style;
                    cell.SetCellValue((i + 1).ToString());
                    //户主姓名
                    sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                    cell = row.GetCell(1);
                    cell.CellStyle = style;
                    cell.SetCellValue(cbfs[i].Cbfmc);
                    //家庭成员数量
                    sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                    cell = row.GetCell(2);
                    cell.CellStyle = style;
                    cell.SetCellValue(familyCount.ToString());
                    endRow++;
                    startRow = endRow;
                }
                EditExcel(workbookSource, endRow, 0);
                FileStream fs = new FileStream(savedFilePath, FileMode.Create, FileAccess.Write);
                workbookSource.Write(fs);
                fs.Close();
                fileStream.Close();
            }
            wait.CloseWait();
            para["ret"] = true;
            return;
        }
        private void FillOneFamily(IWorkbook workbook, string cbfbm, ref int endRow, out int familyCount)
        {
            var cbfJtcys = Cbfjtcys(cbfbm);
            familyCount = cbfJtcys.Count;
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 0; i < familyCount; i++)
            {
                IRow row = sheet.GetRow(endRow + i);
                row.GetCell(3).SetCellValue(cbfJtcys[i].Cyxm);
                row.GetCell(4).SetCellValue(cbfJtcys[i].Cyzjhm);
                row.GetCell(5).SetCellValue(Transcode.CodeToRelationship(cbfJtcys[i].Yhzgx));
                //endRow++;
            }
            endRow = endRow + familyCount - 1;
        }

        private void EditExcel(IWorkbook workbook, int endRow, int sheetIndex)
        {
            //base.EditExcel(workbook, lastRowIndex);
            HSSFSheet sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= endRow + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
        }
    }
}
