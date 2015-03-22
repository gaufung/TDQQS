using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using TdqqClient.Models;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Services.Export.ExportOne
{
    class ListExport:ExportBase,IExport
    {
        public ListExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("xls");
            var saveFilePath = dialogHelper.SaveFile("导出颁证清册");
            if (string.IsNullOrEmpty(saveFilePath)) return;
            if (Export(saveFilePath))
            {
                MessageBox.Show(null, "颁证清册导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "颁证清册导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出颁证清册");
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
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\颁证清册.xls";
            File.Copy(templatePath, savedFilePath, true);
            var wait = para["wait"] as Wait;
            try
            {
                var cbfs = Cbfs(true);
                int startRow = 5, endRow = 5;
                using (var fileStream = new FileStream(savedFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    ICellStyle style = MergetStyle(workbookSource);
                    var rowCount = cbfs.Count;
                    for (int i = 0; i < rowCount; i++)
                    {
                        wait.SetProgress(((double)i / (double)rowCount));
                        var cbfbm = cbfs[i].Cbfbm;
                        string cbfmc = cbfs[i].Cbfmc;
                        //double htmj = 0.0;
                        var fields = Fields(cbfbm);
                        var cbfjtcys = Cbfjtcys(cbfbm);
                        double scmj = 0.0;
                        var endRowField = FillFields(fields, workbookSource, endRow, ref scmj);
                        var endRowFamily = FillFamily(cbfjtcys, workbookSource, endRow);
                        endRow = Math.Max(endRowField, endRowFamily);
                        MergeCells(workbookSource, i + 1, cbfmc, cbfjtcys.Count, cbfbm, scmj, startRow, endRow, style);
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
            catch (Exception)
            {
                wait.CloseWait();
                para["ret"] = false;
                return;
            }
        }

        protected void EditExcel(IWorkbook workbook, int endRow, int sheetIndex)
        {
            //删除多余行
            HSSFSheet sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= endRow + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
        }
        private void MergeCells(IWorkbook workbook, int familyIndex, string cbfmc, int familyCount, string cbfbm, double scmj,
            int startRow, int endRow, ICellStyle style)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = sheet.GetRow(startRow);
            ICell cell;
            //合并序号单元格
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));
            cell = row.GetCell(0);
            cell.CellStyle = style;
            cell.SetCellValue(familyIndex);
            //承包方名称
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
            cell = row.GetCell(1);
            cell.CellStyle = style;
            cell.SetCellValue(cbfmc);
            //家庭成员数量
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
            cell = row.GetCell(2);
            cell.CellStyle = style;
            cell.SetCellValue(familyCount);
            //合同证书
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 4, 4));
            cell = row.GetCell(4);
            cell.CellStyle = style;
            cell.SetCellValue(cbfbm + "J");
            //实测总面积
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 8, 8));
            cell = row.GetCell(8);
            cell.CellStyle = style;
            cell.SetCellValue(scmj);
            //签字
            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 9, 9));

        }
        

        private int FillFields(List<FieldModel> fields , IWorkbook workbook, int endRow, ref double scmj)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 0; i < fields.Count; i++)
            {
                IRow row = sheet.GetRow(endRow + i);
                row.GetCell(5).SetCellValue(fields[i].Dkmc);
                row.GetCell(6).SetCellValue(fields[i].Dkbm);
                
                row.GetCell(7).SetCellValue(fields[i].Scmj);
                scmj += fields[i].Scmj;
            }
            return endRow + fields.Count - 1;
        }

        private int FillFamily(List<CbfjtcyModel> cbfjtcys , IWorkbook workbook, int endRow)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 0; i < cbfjtcys.Count; i++)
            {
                IRow row = sheet.GetRow(endRow + i);
                row.GetCell(3).SetCellValue(cbfjtcys[i].Cyxm);
            }
            return endRow + cbfjtcys.Count - 1;
        }
    }
}
