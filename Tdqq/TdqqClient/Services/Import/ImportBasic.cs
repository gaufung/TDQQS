using TdqqClient.Services.Database;

namespace TdqqClient.Services.Import
{
    /// <summary>
    /// 导入基础数据库
    /// </summary>
    class ImportBasic
    {
      
        private readonly ImportCbfjtcy _importCbfjtcy;
        private readonly  ImportFbf _importFbf;

        public ImportBasic(string basicDatabase)
        {
            IDatabaseService pDatabaseService = new MsAccessDatabase(basicDatabase);
            _importCbfjtcy=new ImportCbfjtcy(new ImportToDb(pDatabaseService));
            _importFbf=new ImportFbf(new ImportToDb(pDatabaseService));            
        }

        public void ImportFbf()
        {
            _importFbf.Import();
        }

        public void ImportCbfjtcy()
        {
            _importCbfjtcy.Import();
        }
    }
}
