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
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;

namespace TdqqClient.Services.Export.ExportOne
{
    class DExport:ExportBase,IExport    
    {
        public DExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("xls");
            var savedPath = dialogHelper.SaveFile("地块信息公示表");
            if (string.IsNullOrEmpty(savedPath)) return;
            if (Export(savedPath))
            {
                MessageBox.Show(null, "地块信息表导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "地块信息表导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }         
        }

        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出地块信息表");
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
            var saveFilePath = para["saveFilePath"].ToString();
            var wait = para["wait"] as Wait;
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\二榜公示表.xls";
            try
            {
                File.Copy(templatePath, saveFilePath, true);
                int startRowIndex = 6;
                var cbfs = Cbfs(true);
                using (var fileStream = new FileStream(saveFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    var workbookSource = new HSSFWorkbook(fileStream);
                    var style = MergetStyle(workbookSource);
                    var rowCount = cbfs.Count;
                    int current = 0;
                    foreach (var cbfModel in cbfs)
                    {
                        wait.SetProgress(((double)current++/ (double)rowCount));
                        Export(workbookSource, cbfModel, ref startRowIndex, style, 1);
                    }
                    EditExcel(workbookSource, startRowIndex, 0);
                    var fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                    workbookSource.Write(fs);
                    fs.Close();
                    fileStream.Close();
                }
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }
        }
        private void Export(HSSFWorkbook workbook,CbfModel cbfModel, ref int startRowIndex, ICellStyle style, int index)
        {          
            List<FieldModel> cbdks = Fields(cbfModel.Cbfbm);
            double htmjSum = 0.0;
            double scmjSum = 0.0;
            var cbfmc = cbfModel.Cbfmc;
            var sheet = workbook.GetSheetAt(0);
            for (int j = 0; j < cbdks.Count; j++)
            {

                var rowSource = (HSSFRow)sheet.GetRow(startRowIndex);
                if (cbdks.Count == 1) rowSource.Height = 1200;
                //DKMC
                rowSource.GetCell(4).SetCellValue(cbdks[j].Dkmc);
                //DKBM
                var dkbm = cbdks[j].Dkbm;
                if (dkbm != string.Empty) dkbm = dkbm.Substring(14, 5);
                rowSource.GetCell(5).SetCellValue(dkbm);
                //四至
                rowSource.GetCell(6).SetCellValue(EditSz(cbdks[j].Dkdz));
                rowSource.GetCell(7).SetCellValue(EditSz(cbdks[j].Dknz));
                rowSource.GetCell(8).SetCellValue(EditSz(cbdks[j].Dkxz));
                rowSource.GetCell(9).SetCellValue(EditSz(cbdks[j].Dkbz));
                //合同面积              
                htmjSum += cbdks[j].Htmj;                
                scmjSum += cbdks[j].Scmj;
                rowSource.GetCell(10).SetCellValue(cbdks[j].Htmj.ToString("f"));
                rowSource.GetCell(11).SetCellValue(cbdks[j].Scmj.ToString("f"));
                startRowIndex++;
            }
            var rowSet = (HSSFRow)sheet.GetRow(startRowIndex - cbdks.Count);
            //设置编号和合并单元格
            sheet.AddMergedRegion(new CellRangeAddress(startRowIndex - cbdks.Count,
                startRowIndex - 1, 0, 0));
            var cellIndex = (HSSFCell)rowSet.GetCell(0);
            cellIndex.SetCellValue(index + 1);
            cellIndex.CellStyle = style;
            //承包方名称
            var cellcbfmc = (HSSFCell)rowSet.GetCell(1);
            cellcbfmc.SetCellValue(cbfmc);
            sheet.AddMergedRegion(new CellRangeAddress(startRowIndex - cbdks.Count,
startRowIndex - 1, 1, 1));
            //设置地块汇总情况
            var cellHtmj = (HSSFCell)rowSet.GetCell(2);
            var cellScmj = (HSSFCell)rowSet.GetCell(3);
            var htmjhz = string.Format("合计：\n {0}块 \n {1}亩", cbdks.Count, htmjSum.ToString("f"));
            var scmjhz = string.Format("合计：\n {0}块 \n {1}亩", cbdks.Count, scmjSum.ToString("f"));
            cellHtmj.SetCellValue(htmjhz);
            cellScmj.SetCellValue(scmjhz);
            cellHtmj.CellStyle = style;
            cellScmj.CellStyle = style;
            sheet.AddMergedRegion(new CellRangeAddress(startRowIndex - cbdks.Count,
startRowIndex - 1, 2, 2));
            sheet.AddMergedRegion(new CellRangeAddress(startRowIndex - cbdks.Count,
startRowIndex - 1, 3, 3));
        }

        protected void EditExcel(IWorkbook workbook, int lastRowIndex, int sheetIndex)
        {
            //删除多余行
            var sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= lastRowIndex + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
            //删除最后一行
            lastRowIndex = sheetSource.LastRowNum;
            sheetSource.ShiftRows(lastRowIndex, lastRowIndex, -1);
            //增加制表信息
            lastRowIndex = sheetSource.LastRowNum;
            var lastRow = (HSSFRow)sheetSource.CreateRow(lastRowIndex);
            lastRow.Height = 500;
            //合并样式
            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            //合并单元格
            sheetSource.AddMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 0, 2));
            sheetSource.AddMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 3, 6));
            sheetSource.AddMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 7, 9));
            sheetSource.AddMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 10, 12));
            //填写内容
            var cell = lastRow.CreateCell(0);
            cell.SetCellValue("制表人：_______________");
            cell.CellStyle = style;
            cell = lastRow.CreateCell(3);
            cell.SetCellValue("制表日期：__________ 年 ______ 月 ______ 日");
            cell.CellStyle = style;
            cell = lastRow.CreateCell(7);
            cell.SetCellValue("审核人:_____________");
            cell = lastRow.CreateCell(10);
            cell.CellStyle = style;
            cell.SetCellValue("审核日期:__________ 年______ 月 ______ 日");
        }
    }
}
