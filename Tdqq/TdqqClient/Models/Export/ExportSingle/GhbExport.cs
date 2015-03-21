using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.Common;
using TdqqClient.Services.Export;

namespace TdqqClient.Models.Export.ExportSingle
{
    class GhbExport:ExportBase
    {
        public GhbExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_07公示结果归户表.xls";
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\农村土地承包经营权公示结果归户表.xls";
            File.Copy(templatePath,saveFilePath,true);
            ExportInfo(saveFilePath,cbfmc,cbfbm);
            ExportField(saveFilePath,cbfmc,cbfbm);
            Export2Pdf.Excel2Pdf(saveFilePath);
            File.Delete(saveFilePath);
        }

        private void ExportInfo(string saveFilePath,string cbfmc,string cbfbm)
        {
            var dtjtcy = SelectCbf_JtcyByCbfbm(cbfbm);
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            var rowFbf = SelectFbfInfo();
            var fbfmc = rowFbf[1].ToString().Trim();
            var fbffzr = rowFbf[2].ToString().Trim();
            var fbfdz = rowFbf[6].ToString().Trim();
            using (var fileStream = new FileStream(saveFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                #region 导出其他

                IWorkbook workbook = new HSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0);
                HSSFRow rowSource = (HSSFRow)sheet.GetRow(3);
                rowSource.GetCell(2).SetCellValue(fbfmc);
                rowSource.GetCell(8).SetCellValue(fbffzr);
                //填写承包方信息
                rowSource = (HSSFRow)sheet.GetRow(4);
                rowSource.GetCell(2).SetCellValue(cbfbm + "J");
                rowSource = (HSSFRow)sheet.GetRow(6);
                rowSource.GetCell(3).SetCellValue(cbfmc);
                var lxdh = string.IsNullOrEmpty(rowCbf[8].ToString().Trim()) ? "/" : rowCbf[8].ToString().Trim();
                rowSource.GetCell(8).SetCellValue(lxdh);
                rowSource = (HSSFRow)sheet.GetRow(7);
                //身份证号码
                rowSource.GetCell(8).SetCellValue(rowCbf[5].ToString());

                rowSource = (HSSFRow)sheet.GetRow(10);
                rowSource.GetCell(2).SetCellValue(fbfdz);
                var yzbm = string.IsNullOrEmpty(rowCbf[7].ToString().Trim()) ? "/" : rowCbf[7].ToString().Trim().Trim();
                rowSource.GetCell(9).SetCellValue(yzbm);
                //填写家庭成员信息
                sheet = (HSSFSheet)workbook.GetSheetAt(1);
                var start_row_index = 2;
                rowSource = (HSSFRow)sheet.GetRow(0);
                rowSource.GetCell(9).SetCellValue(dtjtcy.Rows.Count);
                for (int i = 0; i < dtjtcy.Rows.Count; i++)
                {
                    rowSource = (HSSFRow)sheet.GetRow(start_row_index + i);
                    rowSource.GetCell(0).SetCellValue(dtjtcy.Rows[i][3].ToString());
                    rowSource.GetCell(3).SetCellValue(Transcode.CodeToRelationship(dtjtcy.Rows[i][8].ToString()));
                    var sfzh = string.IsNullOrEmpty(dtjtcy.Rows[i][4].ToString().Trim()) ? "/" : dtjtcy.Rows[i][4].ToString().Trim();
                    rowSource.GetCell(5).SetCellValue(sfzh);
                    var cybz = string.IsNullOrEmpty(dtjtcy.Rows[i][6].ToString().Trim()) ? "/" : dtjtcy.Rows[i][6].ToString().Trim();
                    rowSource.GetCell(8).SetCellValue(cybz);
                }
                #endregion

                var gsjs = string.IsNullOrEmpty(rowCbf[13].ToString().Trim()) ? "/" : rowCbf[13].ToString().Trim();
                sheet.GetRow(19).GetCell(2).SetCellValue(gsjs);


                sheet.GetRow(25).GetCell(4).SetCellValue(GetDcy());
                sheet.GetRow(25).GetCell(8).SetCellValue(GetShrq(-2).ToLongDateString());
                sheet.GetRow(32).GetCell(8).SetCellValue(GetShrq(-2).ToLongDateString());
                //  sheet.GetRow(32).GetCell(4).SetCellValue(row[1].ToString());
                sheet.GetRow(39).GetCell(8).SetCellValue(GetShrq().ToLongDateString());
                FileStream fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }

        private void ExportField(string saveFilePath, string cbfmc, string cbfbm)
        {
            var dtCbdk = SelectFieldsByCbfbm(cbfbm);
            if (dtCbdk == null) return;
            using (FileStream fileStream = new FileStream(saveFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                ICellStyle style = MergetStyle(workbook);
                ISheet sheet = workbook.GetSheetAt(0);
                IRow rowSource;
                var mjSum = 0.0;
                int start_row_index = 14;
                int row_gap = 1;
                for (int i = 0; i < dtCbdk.Rows.Count; i++)
                {
                    //填写四至
                    StringBuilder sz = new StringBuilder();
                    sz.Append("东：" + EditSz(dtCbdk.Rows[i][4].ToString().Trim()) + "\n");
                    sz.Append("南：" + EditSz(dtCbdk.Rows[i][5].ToString().Trim()) + "\n");
                    sz.Append("西：" + EditSz(dtCbdk.Rows[i][6].ToString().Trim()) + "\n");
                    sz.Append("北：" + EditSz(dtCbdk.Rows[i][7].ToString().Trim()) + "\n");                   
                    ICell cell;
                    //填写地块名称
                    cell = sheet.GetRow(start_row_index).GetCell(0);
                    //cell.CellStyle = style;
                    cell.SetCellValue(dtCbdk.Rows[i][1].ToString());
                    //填写地块编码
                    cell = sheet.GetRow(start_row_index).GetCell(1);
                    // cell.CellStyle = style;
                    cell.SetCellValue(dtCbdk.Rows[i][3].ToString().Trim().Substring(14, 5));

                    cell = sheet.GetRow(start_row_index).GetCell(2);
                    cell.SetCellValue(sz.ToString());
                    //保留有效位数
                    double htmj, scmj;
                    if (string.IsNullOrEmpty(dtCbdk.Rows[i][16].ToString()))
                    {
                        htmj = 0.0;
                    }
                    else
                    {
                        htmj = Convert.ToDouble(double.Parse(dtCbdk.Rows[i][16].ToString()).ToString("f"));
                    }
                    if (string.IsNullOrEmpty(dtCbdk.Rows[i][17].ToString()))
                    {
                        scmj = 0.0;
                    }
                    else
                    {
                        scmj = Convert.ToDouble(double.Parse(dtCbdk.Rows[i][17].ToString()).ToString("f"));
                    }
                    mjSum += scmj;
                    //填写合同面积
                    cell = sheet.GetRow(start_row_index).GetCell(4);
                    // cell.CellStyle = style;
                    cell.SetCellValue(htmj);
                    //填写实测面积
                    cell = sheet.GetRow(start_row_index).GetCell(5);
                    //  cell.CellStyle = style;
                    cell.SetCellValue(scmj);
                    //土地用途
                    cell = sheet.GetRow(start_row_index).GetCell(6);
                    //    cell.CellStyle = style;
                    cell.SetCellValue(Transcode.CodeToTdyt(dtCbdk.Rows[i][13].ToString()));
                    //地力等级
                    cell = sheet.GetRow(start_row_index).GetCell(7);
                    //    cell.CellStyle = style;
                    cell.SetCellValue(Transcode.CodeToDldj(dtCbdk.Rows[i][12].ToString()));
                    //土地备注
                    cell = sheet.GetRow(start_row_index).GetCell(8);
                    //    cell.CellStyle = style;
                    var dkbz = string.IsNullOrEmpty(dtCbdk.Rows[i][8].ToString().Trim())
                        ? "/"
                        : dtCbdk.Rows[i][8].ToString().Trim();
                    cell.SetCellValue(dkbz);
                    start_row_index += row_gap;
                }
                //添加承包地块信息
                rowSource = (HSSFRow)sheet.GetRow(11);
                rowSource.GetCell(3).SetCellValue(dtCbdk.Rows.Count + "块");
                rowSource.GetCell(4).SetCellValue(mjSum.ToString("f") + "亩");
                EditExcel(workbook, start_row_index, 0);
                FileStream fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }
        protected  void EditExcel(IWorkbook workbook, int lastRowIndex, int sheetIndex)
        {
            var sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= lastRowIndex + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
        }
        
    }
}
