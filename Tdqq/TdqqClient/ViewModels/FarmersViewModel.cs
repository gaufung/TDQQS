using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using TdqqClient.Commands;
using TdqqClient.Models;
using TdqqClient.Services.Database;

namespace TdqqClient.ViewModels
{
    public class FarmersViewModel:NotificationObject
    {

        #region 属性字段
        private string _searchFarmer;

        public string SearchFarmer
        {
            get { return _searchFarmer; }
            set
            {
                _searchFarmer = value;
                this.RaisePropertyChanged("SearchFarmer");
            }
        }
        private FarmerModel _selectFarmer;

        public FarmerModel SelectFarmer
        {
            get { return _selectFarmer; }
            set
            {
                _selectFarmer = value;
                this.RaisePropertyChanged("SelectFarmer");
            }
        }
        private List<FarmerModel> _farmerList;

        public List<FarmerModel> FarmerList
        {
            get { return _farmerList; }
            set
            {
                _farmerList = value;
                this.RaisePropertyChanged("FarmerList");
            }
        }
        
        #endregion

        #region 移动和关闭窗口
        
        public DelegateCommand MouseLeftButtonDownCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
        }

        public DelegateCommand WindowMoveCommand { get; set; }

        private void MoveWindow(object parameter)
        {
            this.OnMovingRequest();
        }   
        #endregion

        private readonly string _personDatabase;
        private readonly string _selectFeauture;
        private readonly IMap _pMap;

        public FarmersViewModel(string personDatabase,string selectFeature,IMap pMap)
        {
            _personDatabase = personDatabase;
            _selectFeauture = selectFeature;
            _pMap = pMap;
            InitGridInfos();
            InitCommands();
        }

        private void InitGridInfos()
        {
            FarmerList=new List<FarmerModel>();
            var sqlString = string.Format("select distinct CBFBM,CBFMC from {0}", _selectFeauture);
            IDatabaseService pDatabaseService = new MsAccessDatabase(_personDatabase);
            var dt = pDatabaseService.Query(sqlString);
            if (dt == null) return ;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                FarmerList.Add(new FarmerModel()
                {
                    Cbfbm = dt.Rows[i][0].ToString(),
                    Cbfmc = dt.Rows[i][1].ToString()
                });
            }
        }

        private void InitCommands()
        {
            MouseLeftButtonDownCommand=new DelegateCommand();
            WindowMoveCommand=new DelegateCommand();
            SearchCommand=new DelegateCommand();
            SelectFarmerCommand=new DelegateCommand();
            MouseLeftButtonDownCommand.ExecuteAction=new Action<object>(CloseWindow);
            WindowMoveCommand.ExecuteAction=new Action<object>(MoveWindow);
            SearchCommand.ExecuteAction=new Action<object>(Search);
            SelectFarmerCommand.ExecuteAction = new Action<object>(SelectChangeFarmer);
        }
        #region 查询操作

        public DelegateCommand SearchCommand { get; set; }

        private void Search(object parameter)
        {
            if (string.IsNullOrEmpty(SearchFarmer)) return;
            foreach (var farmerModel in FarmerList)
            {
                if (farmerModel.Cbfmc==SearchFarmer)
                {

                    SelectFarmer = farmerModel;
                    MessageBox.Show(null, "查询到此人", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }              
             MessageBox.Show(null, "未查询到此人", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);            
        }

        #endregion

        #region 选择操作

        public DelegateCommand SelectFarmerCommand { get; set; }

        private void SelectChangeFarmer(object parameter)
        {
            if (SelectFarmer == null) return;
            var pFeatureLayer = _pMap.get_Layer(0) as IFeatureLayer;
            var pSection = pFeatureLayer as IFeatureSelection;
            IQueryFilter queryFilter = new SpatialFilterClass();
            queryFilter.WhereClause = string.Format("trim(CBFBM)={0}", SelectFarmer.Cbfbm);
            pSection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
            IActiveView pview = _pMap as IActiveView;
            pview.Refresh();
        }
        #endregion

    }
}
