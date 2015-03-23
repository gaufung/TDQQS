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
    class SignExport:ExportBase,IExport
    {
        public SignExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("xls");
            var saveFilePath = dialogHelper.SaveFile("导出签字表");
            if (string.IsNullOrEmpty(saveFilePath)) return;
            if (Export(saveFilePath))
            {
                MessageBox.Show(null, "签字表导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "签字表导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出签字表");
            var para = new Hashtable()
            {
                {"wait",wait},{"saveFilePath",saveFilePath},{"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void ExportF(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            var saveFilePath = para["saveFilePath"].ToString();
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\土地承包经营权确权签字表.xls";
            File.Copy(templatePath, saveFilePath, true);
            try
            {
                var cbfs = Cbfs(false);
                var startRow = 4;
                var endRow = 4;

                using (var fileStream = new FileStream(saveFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    //设置格式要素
                    HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    int rowCount = cbfs.Count;
                    var style = MergetStyle(workbookSource);
                    for (int i = 0; i < rowCount; i++)
                    {
                        wait.SetProgress(((double)i / (double)rowCount));
                        var cbfbm = cbfs[i].Cbfbm;
                        var fields = Fields(cbfbm);
                        var cbfjtcys = Cbfjtcys(cbfbm);
                        var htmj = 0.0;
                        var scmj = 0.0;
                        var endRowField = FillFieldSign(fields, workbookSource, endRow, ref htmj, ref scmj);
                        var endRowFamily = FillFamily(cbfjtcys, workbookSource, endRow);
                        endRow = Math.Max(endRowField, endRowFamily);
                        MergeCellsSign(workbookSource, i + 1, cbfs[i].Cbfmc, cbfjtcys.Count, scmj, htmj, startRow, endRow, style);
                        endRow++;
                        startRow = endRow;
                    }
                    EditExcelSign(workbookSource, endRow);
                    FileStream fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                    workbookSource.Write(fs);
                    fs.Close();
                    fileStream.Close();
                }
                ExportIndexTable(saveFilePath);
                wait.CloseWait();
                para["ret"] = true;

            }
            catch (Exception)
            {
                wait.CloseWait();
                para["ret"] = false;
                return;
            }
        }

        private void EditExcelSign(HSSFWorkbook workbookSource, int endRow)
        {
            //throw new NotImplementedException();
            //删除多余行
            HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
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
            ICellStyle style = workbookSource.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            //合并单元格
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 0, 4));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 5, 9));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 10, 13));
            sheetSource.AddMergedRegion(new CellRangeAddress(endRow, endRow, 14, 19));
            //填写内容
            var cell = lastRow.CreateCell(0);
            cell.CellStyle = style;
            cell.SetCellValue("制表人：_______________");

            cell = lastRow.CreateCell(5);
            cell.CellStyle = style;
            cell.SetCellValue("制表日期：__________ 年 ______ 月 ______ 日");
            cell = lastRow.CreateCell(10);
            cell.CellStyle = style;
            cell.SetCellValue("审核人:_____________");
            cell = lastRow.CreateCell(14);
            cell.CellStyle = style;
            cell.SetCellValue("审核日期:________ 年______ 月 ______ 日");
        }

        private void MergeCellsSign(HSSFWorkbook workbookSource, int p1, string cbfmc, int p2, double scmj, double htmj, int startRow, int endRow, ICellStyle style)
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
            //备注
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 18, 18));
            HSSFCell cellbz = (HSSFCell)rowSet.GetCell(18);
            cellbz.SetCellValue(string.Empty);
            cellbz.CellStyle = style;
            //盖章签字
            // sheetSource.AddMergedRegion(new Region(startRow, 19, endRow, 19));
            sheetSource.AddMergedRegion(new CellRangeAddress(startRow, endRow, 19, 19));
            HSSFCell cellQzgz = (HSSFCell)rowSet.GetCell(19);
            cellQzgz.SetCellValue(string.Empty);
            cellQzgz.CellStyle = style;
        }

        private int FillFamily(List<CbfjtcyModel> cbfjtcys , HSSFWorkbook workbookSource, int endRow)
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

        private int FillFieldSign(List<FieldModel> fields , HSSFWorkbook workbookSource, int endRow, ref double htmj, ref double scmj)
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
                rowSource.GetCell(12).SetCellType(CellType.NUMERIC);
                rowSource.GetCell(12).SetCellValue(fields[j].Scmj.ToString("f"));
                rowSource.GetCell(14).SetCellType(CellType.NUMERIC);
                rowSource.GetCell(14).SetCellValue(fields[j].Htmj.ToString("f"));
                //耕地类型
                rowSource.GetCell(16).SetCellValue(Transcode.CodeToDklb(fields[j].Dklb));
                //是否基本农田
                rowSource.GetCell(17).SetCellValue(Transcode.CodeToSfjbnt(fields[j].Sfjbnt));
                //地块备注
                rowSource.GetCell(18).SetCellValue(fields[j].Dkbz);
            }
            return endRow + fields.Count - 1;
        }
       
    }
}
