using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using TdqqClient.Models;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Services.Export.ExportOne
{
    class OpenExport:ExportBase,IExport
    {
        public OpenExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("xls");
            var saveFilePath = dialogHelper.SaveFile("导出公示表");
            if (string.IsNullOrEmpty(saveFilePath)) return;
            if (Export(saveFilePath))
            {
                MessageBox.Show(null, "公示表导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "公式表导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出公示表");
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
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\土地承包经营权确权公示表.xls";
            File.Copy(templatePath, savedFilePath, true);
            try
            {
                //按照承包方名称
                var cbfs = Cbfs(false);
                int startRow = 4, endRow = 4;
                using (var fileStream = new FileStream(savedFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    ICellStyle style = MergetStyle(workbookSource);
                    var rowCount = cbfs.Count;
                    for (int i = 0; i < rowCount; i++)
                    {
                        wait.SetProgress((double)i / (double)rowCount);
                        var cbfbm = cbfs[i].Cbfbm;

                        //
                        string cbfmc = cbfs[i].Cbfmc;
                        var fields = Fields(cbfbm);
                        var cbfjtcys = Cbfjtcys(cbfbm);
                        double htmj = 0.0;
                        double scmj = 0.0;
                        var endRowField = FillFields(fields, workbookSource, endRow, ref htmj, ref scmj);
                        var endRowFamily = FillFamily(cbfjtcys, workbookSource, endRow);
                        endRow = Math.Max(endRowField, endRowFamily);
                        MergeCells(workbookSource, i + 1, cbfmc, cbfjtcys.Count, scmj, htmj, startRow, endRow, style);
                        endRow++;
                        startRow = endRow;
                    }
                    EditExcel(workbookSource, endRow, 0);
                    FileStream fs = new FileStream(savedFilePath, FileMode.Create, FileAccess.Write);
                    workbookSource.Write(fs);
                    fs.Close();
                    fileStream.Close();
                }
                ExportIndexTable(savedFilePath);
                para["ret"] = true;
            }
            catch (Exception e)
            {
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }
        }

        private void MergeCells(IWorkbook workbookSource, int p1, string cbfmc, int p2, double scmj, double htmj, int startRow, int endRow, ICellStyle style)
        {
            HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
            HSSFRow rowSet = (HSSFRow)sheetSource.GetRow(startRow);

            //处理编号合并单元格
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));
            HSSFCell cellIndex = (HSSFCell)rowSet.GetCell(0);
            cellIndex.SetCellValue(p1);
            cellIndex.CellStyle = style;
            //处理户主信息合并单元格
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
            HSSFCell cellCbfmc = (HSSFCell)rowSet.GetCell(1);
            cellCbfmc.SetCellValue(cbfmc);
            cellCbfmc.CellStyle = style;
            //处理家庭成员个数
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
            HSSFCell cellJtcygs = (HSSFCell)rowSet.GetCell(2);
            cellJtcygs.SetCellValue(p2);
            cellJtcygs.CellStyle = style;
            //处理实测总面积
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 13, 13));
            HSSFCell cellScmj = (HSSFCell)rowSet.GetCell(13);
            cellScmj.SetCellValue(scmj.ToString("f"));
            cellScmj.CellStyle = style;
            //处理合同总面积
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 15, 15));
            HSSFCell cellHtmj = (HSSFCell)rowSet.GetCell(15);
            cellHtmj.SetCellValue(htmj.ToString("f"));
            cellHtmj.CellStyle = style;
            //盖章签字
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 17, 17));
            HSSFCell cellQzgz = (HSSFCell)rowSet.GetCell(17);
            cellQzgz.SetCellValue(string.Empty);
            cellQzgz.CellStyle = style;
        }

        private int FillFamily(List<CbfjtcyModel> cbfjtcys , IWorkbook workbookSource, int endRow)
        {
            HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
            for (int k = 0; k < cbfjtcys.Count; k++)
            {
                HSSFRow rowSource = (HSSFRow)sheetSource.GetRow(endRow + k);
                rowSource.GetCell(3).SetCellValue(cbfjtcys[k].Cyxm);
                rowSource.GetCell(4).SetCellValue(cbfjtcys[k].Cyzjhm);
                rowSource.GetCell(5).SetCellValue(Transcode.CodeToRelationship(cbfjtcys[k].Yhzgx));
            }
            return endRow + cbfjtcys.Count - 1;
        }

        private int FillFields(List<FieldModel> fields , IWorkbook workbookSource, int endRow, ref double htmj, ref double scmj)
        {
            HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
            //填写承包地块信息，承包方名称
            for (int j = 0; j < fields.Count; j++)
            {

                HSSFRow rowSource = (HSSFRow)sheetSource.GetRow(endRow + j);
                //地块名称
                rowSource.GetCell(6).SetCellValue(fields[j].Dkmc);
                //地块编码
                var dkbm = fields[j].Dkbm;
                if (dkbm != string.Empty)
                {
                    dkbm = dkbm.Substring(14, 5);
                }
                rowSource.GetCell(7).SetCellValue(dkbm);
                //四至
                rowSource.GetCell(8).SetCellValue(EditSz(fields[j].Dkdz));
                rowSource.GetCell(9).SetCellValue(EditSz(fields[j].Dknz));
                rowSource.GetCell(10).SetCellValue(EditSz(fields[j].Dkxz));
                rowSource.GetCell(11).SetCellValue(EditSz(fields[j].Dkbz));
                //合同面积和实测面积

                htmj += fields[j].Htmj;
                scmj += fields[j].Scmj;
                ICell cellscmj = rowSource.GetCell(12);
                ICell cellhtmj = rowSource.GetCell(14);
                if (fields[j].Htmj > fields[j].Scmj)
                {
                    cellhtmj.CellStyle = LessStyle(workbookSource);
                    cellscmj.CellStyle = LessStyle(workbookSource);
                }
                cellscmj.SetCellValue(fields[j].Scmj.ToString("f"));
                cellhtmj.SetCellValue(fields[j].Htmj.ToString("f"));
                //耕地类型
                rowSource.GetCell(16).SetCellValue(Transcode.CodeToDklb(fields[j].Dklb));
            }
            return endRow + fields.Count - 1;
        }

       

        private ICellStyle LessStyle(IWorkbook workbook)
        {
            HSSFFont font = (HSSFFont)workbook.CreateFont();
            HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
            font.Color = HSSFColor.RED.index;
            style.SetFont(font);
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
            style.WrapText = true;
            return style;
        }

        private void EditExcel(IWorkbook workbook, int endRow, int sheetIndex)
        {
            //删除多余行
            HSSFSheet sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= endRow + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
            //删除最后一行
            endRow = sheetSource.LastRowNum;
            sheetSource.ShiftRows(endRow, endRow, -1);
            //增加制表信息
            endRow = sheetSource.LastRowNum;
            var lastRow = (HSSFRow)sheetSource.CreateRow(endRow);
            lastRow.Height = 500;
            //合并样式
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            //合并单元格
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 0, 3));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 4, 7));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 8, 11));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 12, 17));
            //填写内容
            var cell = lastRow.CreateCell(0);
            cell.SetCellValue("制表人：_______________");
            cell.CellStyle = style;
            cell = lastRow.CreateCell(4);
            cell.SetCellValue("制表日期：__________ 年 ______ 月 ______ 日");
            cell.CellStyle = style;
            cell = lastRow.CreateCell(8);
            cell.SetCellValue("审核人:_____________");
            cell = lastRow.CreateCell(12);
            cell.CellStyle = style;
            cell.SetCellValue("审核日期:________ 年______ 月 ______ 日");
        }
    }
}
