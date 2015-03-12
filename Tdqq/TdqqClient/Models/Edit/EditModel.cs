using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geodatabase;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;

namespace TdqqClient.Models.Edit
{
    /// <summary>
    /// 编辑按钮的抽象基类
    /// </summary>
    class EditModel
    {
        protected string PersonDatabase;
        protected string SelectFeature;
        protected string BasicDatabase;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="personDatabase">个人地理数据库</param>
        /// <param name="selectFeature">选择的要素类</param>
        /// <param name="basicDatabase">基础数据库</param>
        public EditModel(string personDatabase, string selectFeature, string basicDatabase)
        {
            PersonDatabase = personDatabase;
            SelectFeature = selectFeature;
            BasicDatabase = basicDatabase;
        }

        public virtual void Edit(object parameter){}

        protected bool CheckEditFieldsExist()
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            var flag = pFeatureClass.FieldExistCheck("CBFMC", "YHTMJ",
                "DKMC", "YSDM",
                "DKBZXX", "ZJRXM",
                "FBFBM", "SYQXZ",
                "DKLB", "DLDJ",
                "TDYT", "SFJBNT",
                "TDLYLX", "CBJYQQDFS",
                "HTMJ", "SCMJ",
                "BSM", "DKDZ",
                "DKNZ", "DKXZ",
                "DKBZ", "DKBM");
            pAeFactory.ReleaseFeautureClass(pFeatureClass);
            return flag;
        }   

    }
}
