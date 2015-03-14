using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Cells;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.Common;
using TdqqClient.Services.Database;

namespace TdqqClient.Models.Export
{
    class ExportBase
    {
        protected string PersonDatabase;
        protected string SelectFeature;
        protected string BasicDatabase;

        public ExportBase(string personDatabase, string selectFeature, string basicDatabase)
        {
            PersonDatabase = personDatabase;
            SelectFeature = selectFeature;
            BasicDatabase = basicDatabase;
        }

        /// <summary>
        /// 将模板拷贝到要保存的目录下，同名文件并覆盖
        /// </summary>
        /// <param name="templatePath">模板的路径</param>
        /// <param name="fileType">文件的类型</param>
        /// <param name="title">对话框的标题</param>
        /// <returns>拷贝后的文件的路径</returns>
        protected string CopyTemplateToSave(string templatePath, string fileType,string title)
        {
            var dialogHelper=new DialogHelper(fileType);
            var savePath = dialogHelper.SaveFile(title);
            File.Copy(templatePath, savePath, true);
            return savePath;
        }

        /// <summary>
        /// 获取所有农户的信息
        /// </summary>
        /// <returns></returns>
        protected IEnumerable<FarmerModel> Farmers()
        {
            var sqlString = string.Format("select distinct CBFBM,CBFMC from {0} order by CBFBM", SelectFeature);
            IDatabaseService pDatabaseService = new MsAccessDatabase(PersonDatabase);
            var dt = pDatabaseService.Query(sqlString);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                yield return new FarmerModel()
                {
                    Index = i+1,
                    Cbfbm = dt.Rows[i][0].ToString().Trim(),
                    Cbfmc = dt.Rows[i][1].ToString().Trim()
                };
            }
        }

        /// <summary>
        /// 导出姓名索引表
        /// </summary>
        /// <param name="excelPath"></param>
        protected void ExportIndexTable(string excelPath)
        {
            var farmers = Farmers();
            var sortedFarmers = farmers.OrderBy(f => f.Cbfmc);
            using (var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                //在姓名排序模板的sheet2中
                ISheet sheet = workbook.GetSheetAt(1);
                int startRow = 2;
                int rowCount = 7;
                int index = 0;
                foreach (var farmer in sortedFarmers)
                {
                    int currentRow = startRow + index / rowCount;
                    IRow row = sheet.GetRow(currentRow);
                    row.GetCell(index % rowCount * 2).SetCellValue((farmer.Index).ToString());
                    row.GetCell(index % rowCount * 2 + 1).SetCellValue(farmer.Cbfmc);
                    index++;
                }
                int endRow = startRow + index / rowCount + 1;
                for (int i = sheet.LastRowNum; i >= endRow + 1; i--)
                {
                    sheet.ShiftRows(i, i + 1, -1);
                }
                var fs = new FileStream(excelPath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }

        /// <summary>
        /// 将Excel转换成pdf文件(同名文件)
        /// </summary>
        /// <param name="excelFilePath">文件路径</param>
        /// <param name="isDelete">是否删除源文件</param>
        protected void ConvertXlsToPdf(string excelFilePath, bool isDelete)
        {
            if (!File.Exists(excelFilePath)) return;
            var targetFile = System.IO.Path.GetFileNameWithoutExtension(excelFilePath) + @".pdf";
            var li = new License();
            string path = AppDomain.CurrentDomain.BaseDirectory + @"\License.lic";
            li.SetLicense(path);
            using (var fileStream = new FileStream(excelFilePath, FileMode.Open))
            {
                var workbook = new Aspose.Cells.Workbook();
                workbook.Open(fileStream);
                workbook.Save(targetFile, FileFormatType.Pdf);
                fileStream.Close();
            }
            if (isDelete)
            {
                File.Delete(excelFilePath);
            }
        }

        /// <summary>
        /// 编辑四至，如果获取的四至中有下划线开始，去掉下划线
        /// </summary>
        /// <param name="sz">从数据库中获取的四至信息</param>
        /// <returns>去掉的下划线的结果</returns>
        protected string EditSz(string sz)
        {
            return sz.StartsWith("_") ? sz.Substring(1) : sz;
        }


        /// <summary>
        /// 获取调查日期
        /// </summary>
        /// <returns>返回日期</returns>
        protected DateTime GetDcrq()
        {
            var sqlString = string.Format("select distinct CBFDCRQ from CBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            if (dt != null && dt.Rows.Count != 0)
            {
                return string.IsNullOrEmpty(dt.Rows[0][0].ToString().Trim())
                    ? System.DateTime.Now
                    : Convert.ToDateTime(dt.Rows[0][0].ToString().Trim());
            }
            return System.DateTime.Now;
        }
        /// <summary>
        /// 获取调查日期
        /// </summary>
        /// <param name="gapDay">间隔的天数</param>
        /// <returns>返回的时间</returns>
        protected DateTime GetDcrq(int gapDay)
        {
            var dcrq = GetDcrq();
            var timeSpan = new TimeSpan(gapDay, 0, 0, 0);
            return dcrq.Add(timeSpan);
        }

        /// <summary>
        /// 获取审核日期
        /// </summary>
        /// <returns>数据库中的时间，如果没有，则返回系统当前时间</returns>
        protected DateTime GetShrq()
        {
            var sqlString = string.Format("select distinct GSSHRQ from CBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            if (dt != null && dt.Rows.Count != 0)
            {
                return string.IsNullOrEmpty(dt.Rows[0][0].ToString().Trim())
                    ? System.DateTime.Now
                    : Convert.ToDateTime(dt.Rows[0][0].ToString().Trim());
            }
            return System.DateTime.Now;
        }

        /// <summary>
        /// 合并单元格样式
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns>样式</returns>
        protected ICellStyle MergetStyle(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.CENTER;
            style.VerticalAlignment = VerticalAlignment.CENTER;
            style.BorderBottom = BorderStyle.THIN;
            style.BorderRight = BorderStyle.THIN;
            style.BorderLeft = BorderStyle.THIN;
            style.BorderTop = BorderStyle.THIN;
            style.WrapText = true;
            return style;
        }

        /// <summary>
        /// 获取与审核日期相关的日期
        /// </summary>
        /// <param name="gapDay">间隔的期间</param>
        /// <returns></returns>
        protected DateTime GetShrq(int gapDay)
        {
            var shrq = GetShrq();
            TimeSpan timeSpan = new TimeSpan(gapDay, 0, 0, 0);
            return shrq.Add(timeSpan);
        }
        /// <summary>
        /// 获取调查员
        /// </summary>
        /// <returns>调查员</returns>
        protected string GetDcy()
        {
            var sqlString = string.Format("select distinct CBFDCY from CBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            if (dt != null && dt.Rows.Count != 0)
            {
                return dt.Rows[0][0].ToString();
            }
            return string.Empty;
        }
        /// <summary>
        /// 获取审核人
        /// </summary>
        /// <returns>审核人</returns>
        protected string GetShr()
        {
            var sqlString = string.Format("select distinct GSSHR from CBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            if (dt != null && dt.Rows.Count != 0)
            {
                return dt.Rows[0][0].ToString();
            }
            return string.Empty;
        }

        #region 获取信息
        /// <summary>
        /// 选择出发包方所有信息,用一行数据保存
        /// </summary>
        /// <returns></returns>
        protected DataRow SelectFbfInfo()
        {
            var sqlString =
                string.Format(
                    "select FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS from FBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            return (dt == null || dt.Rows.Count != 1) ? null : dt.Rows[0];          
        }
        /// <summary>
        /// 从承包方家庭成员信息中获取信息
        /// </summary>
        /// <param name="cbfbm">承包方编码</param>
        /// <returns>获取信息</returns>
        protected System.Data.DataTable SelectCbf_JtcyByCbfbm(string cbfbm)
        {
            var sqlString =
              string.Format(
                  "select CBFBM,CBFMC,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,CYBZ,YHZGX,CYSZC,YZBM,LXDH,SFGYR" +
                  " from CBF_JTCY where trim(CBFBM) = '{0}' order by YHZGX", cbfbm);
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            return dt;
        }

        /// <summary>
        /// 根据某一户承包方编码，获取该农户的所有地块
        /// </summary>
        /// <param name="cbfbm">承包方编码</param>
        /// <returns>筛选出来的地块</returns>
        protected System.Data.DataTable SelectFieldsByCbfbm(string cbfbm)
        {
            var sqlString =
             string.Format(
                 "select CBFMC,DKMC,YHTMJ,DKBM,DKDZ,DKNZ,DKXZ,DKBZ,DKBZXX,ZJRXM,DKLB,TDLYLX,DLDJ,TDYT,SFJBNT,CBJYQQDFS,HTMJ,SCMJ" +
                 " from {0} where trim(CBFBM) = '{1}' order by DKBM ", SelectFeature, cbfbm);
            var accessFactory = new MsAccessDatabase(PersonDatabase);
            return accessFactory.Query(sqlString);
        }
        /// <summary>
        /// 筛选出拥有地块的承包方编码
        /// </summary>
        /// <returns>结果</returns>
        protected System.Data.DataTable SelectCbfbmOwnFields()
        {
            var sqlString = string.Format("Select distinct CBFBM,CBFMC From {0} where CBFBM NOT LIKE  '{1}' order by CBFBM ",
            SelectFeature, "99999999999999%");
            var accessFactory = new MsAccessDatabase(PersonDatabase);
            return accessFactory.Query(sqlString);
        }
        /// <summary>
        /// 筛选出承包方信息，根据承包方名称
        /// </summary>
        /// <param name="cbfbm">承包方编码</param>
        /// <returns>一行记录</returns>
        protected DataRow SelectCbfInfoByCbfbm(string cbfbm)
        {
            var sqlString =
                string.Format(
                    "select CBFBM,CBFLX,CBFMC,CYXB,CBFZJLX,CBFZJHM,CBFDZ,YZBM,LXDH,CBFCYSL,CBFDCRQ,CBFDCY,CBFDCJS,GSJS,GSJSR,GSSHRQ," +
                    "GSSHR from CBF where trim(CBFBM) = '{0}'", cbfbm);
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            return (dt == null || dt.Rows.Count != 1) ? null : dt.Rows[0];
        }
        #endregion

        /// <summary>
        /// 导出单个农户结果结果
        /// </summary>
        /// <param name="cbfmc">承包方名称</param>
        /// <param name="cbfbm">承包方编码</param>
        /// <param name="folderPath">文件夹路径</param>
        /// <param name="edgeFeature">边界要素</param>
        public virtual void Export(string cbfmc,string cbfbm,string folderPath,string edgeFeature="")
        {
            
        }
    }
}
