using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class GhbExport:ExportBase
    {
        public GhbExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_07公示结果归户表.xls";
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\农村土地承包经营权公示结果归户表.xls";
            File.Copy(templatePath, saveFilePath, true);
            ExportInfo(saveFilePath, cbfmc, cbfbm);
            ExportField(saveFilePath, cbfmc, cbfbm);
            Export2Pdf.Excel2Pdf(saveFilePath);
            File.Delete(saveFilePath);
        }

        private void ExportInfo(string saveFilePath, string cbfmc, string cbfbm)
        {
            var cbfjtcys =Cbfjtcys(cbfbm);
            var cbf = Cbf(cbfbm);
            var fbf = Fbf();
            if (fbf == null) return;
            var fbfmc = fbf.Fbfmc;
            var fbffzr = fbf.Fbffzrxm;
            var fbfdz = fbf.Fbfdz;
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
                var lxdh = string.IsNullOrEmpty(cbf.Lxdh) ? "/" : cbf.Lxdh;
                rowSource.GetCell(8).SetCellValue(lxdh);
                rowSource = (HSSFRow)sheet.GetRow(7);
                //身份证号码
                rowSource.GetCell(8).SetCellValue(cbf.Cbfzjhm);

                rowSource = (HSSFRow)sheet.GetRow(10);
                rowSource.GetCell(2).SetCellValue(fbfdz);
                var yzbm = string.IsNullOrEmpty(cbf.Yzbm) ? "/" : cbf.Yzbm;
                rowSource.GetCell(9).SetCellValue(yzbm);
                //填写家庭成员信息
                sheet = (HSSFSheet)workbook.GetSheetAt(1);
                var start_row_index = 2;
                rowSource = (HSSFRow)sheet.GetRow(0);
                rowSource.GetCell(9).SetCellValue(cbfjtcys.Count);
                for (int i = 0; i < cbfjtcys.Count; i++)
                {
                    rowSource = (HSSFRow)sheet.GetRow(start_row_index + i);
                    rowSource.GetCell(0).SetCellValue(cbfjtcys[i].Cyxm);
                    rowSource.GetCell(3).SetCellValue(Transcode.CodeToRelationship(cbfjtcys[i].Yhzgx));
                    var sfzh = string.IsNullOrEmpty(cbfjtcys[i].Cyzjhm) ? "/" : cbfjtcys[i].Cyzjhm;
                    rowSource.GetCell(5).SetCellValue(sfzh);
                    var cybz = string.IsNullOrEmpty(cbfjtcys[i].Cybz) ? "/" :cbfjtcys[i].Cybz;
                    rowSource.GetCell(8).SetCellValue(cybz);
                }
                #endregion

                var dcSh = DcSh();
                var gsjs = string.IsNullOrEmpty(dcSh.Gsjs) ? "/" : dcSh.Gsjs;
                sheet.GetRow(19).GetCell(2).SetCellValue(gsjs);


                sheet.GetRow(25).GetCell(4).SetCellValue(dcSh.Cbfdcy);
                sheet.GetRow(25).GetCell(8).SetCellValue(dcSh.Gsshrq.Add(new TimeSpan(-2,0,0,0)).ToLongDateString());
                sheet.GetRow(32).GetCell(8).SetCellValue(dcSh.Gsshrq.Add(new TimeSpan(-2, 0, 0, 0)).ToLongDateString());
                sheet.GetRow(39).GetCell(8).SetCellValue(dcSh.Gsshrq.ToLongDateString());
                var fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }

        private void ExportField(string saveFilePath, string cbfmc, string cbfbm)
        {
            var fields = Fields(cbfbm);
            if (fields == null) return;
            using (var fileStream = new FileStream(saveFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                ICellStyle style = MergetStyle(workbook);
                ISheet sheet = workbook.GetSheetAt(0);
                IRow rowSource;
                var mjSum = 0.0;
                int start_row_index = 14;
                int row_gap = 1;
                for (int i = 0; i < fields.Count; i++)
                {
                    //填写四至
                    StringBuilder sz = new StringBuilder();
                    sz.Append("东：" + EditSz(fields[i].Dkdz) + "\n");
                    sz.Append("南：" + EditSz(fields[i].Dknz) + "\n");
                    sz.Append("西：" + EditSz(fields[i].Dkxz) + "\n");
                    sz.Append("北：" + EditSz(fields[i].Dkbz) + "\n");
                    ICell cell;
                    //填写地块名称
                    cell = sheet.GetRow(start_row_index).GetCell(0);
                    //cell.CellStyle = style;
                    cell.SetCellValue(fields[i].Dkmc);
                    //填写地块编码
                    cell = sheet.GetRow(start_row_index).GetCell(1);
                    // cell.CellStyle = style;
                    cell.SetCellValue(fields[i].Dkbm.Substring(14, 5));

                    cell = sheet.GetRow(start_row_index).GetCell(2);
                    cell.SetCellValue(sz.ToString());
                    //保留有效位数
                    double htmj = fields[i].Htmj;
                    double scmj=fields[i].Scmj;                    
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
                    cell.SetCellValue(Transcode.CodeToTdyt(fields[i].Tdyt));
                    //地力等级
                    cell = sheet.GetRow(start_row_index).GetCell(7);
                    //    cell.CellStyle = style;
                    cell.SetCellValue(Transcode.CodeToDldj(fields[i].Dldj));
                    //土地备注
                    cell = sheet.GetRow(start_row_index).GetCell(8);
                    //    cell.CellStyle = style;
                    var dkbz = string.IsNullOrEmpty(fields[i].Dkbz)
                        ? "/"
                        : fields[i].Dkbz;
                    cell.SetCellValue(dkbz);
                    start_row_index += row_gap;
                }
                //添加承包地块信息
                rowSource = (HSSFRow)sheet.GetRow(11);
                rowSource.GetCell(3).SetCellValue(fields.Count + "块");
                rowSource.GetCell(4).SetCellValue(mjSum.ToString("f") + "亩");
                EditExcel(workbook, start_row_index, 0);
                FileStream fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }
        protected void EditExcel(IWorkbook workbook, int lastRowIndex, int sheetIndex)
        {
            var sheetSource = (HSSFSheet)workbook.GetSheetAt(sheetIndex);
            for (int i = sheetSource.LastRowNum; i >= lastRowIndex + 1; i--)
            {
                sheetSource.ShiftRows(i, i + 1, -1);
            }
        }
    }
}
