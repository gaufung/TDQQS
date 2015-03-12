using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using TdqqClient.Commands;
using TdqqClient.Models;
using TdqqClient.Services.AE;
using TdqqClient.Services.Database;

namespace TdqqClient.ViewModels
{
    public class FieldsViewModel : NotificationObject
    {
        #region 属性字段

        private string _dkmc;

        public string Dkmc
        {
            get { return _dkmc; }
            set
            {
                _dkmc = value;
                this.RaisePropertyChanged("Dkmc");
            }
        }

        private string _cbfmc;

        public string Cbfmc
        {
            get { return _cbfmc; }
            set
            {
                _cbfmc = value;
                this.RaisePropertyChanged("Cbfmc");
            }
        }

        private string _dkbm;

        public string Dkbm
        {
            get { return _dkbm; }
            set
            {
                _dkbm = value;
                this.RaisePropertyChanged("Dkbm");
            }
        }
        private double _scmj;

        public double Scmj
        {
            get { return _scmj; }
            set
            {
                _scmj = value;
                this.RaisePropertyChanged("Scmj");
            }
        }

        private double _htmj;

        public double Htmj
        {
            get { return _htmj; }
            set
            {
                _htmj = value;
                this.RaisePropertyChanged("Htmj");
            }
        }

        private string _dkdz;

        public string Dkdz
        {
            get { return _dkdz; }
            set
            {
                _dkdz = value;
                this.RaisePropertyChanged("Dkdz");
            }
        }
        private string _dknz;

        public string Dknz
        {
            get { return _dknz; }
            set
            {
                _dknz = value;
                this.RaisePropertyChanged("Dknz");
            }
        }
        private string _dkxz;

        public string Dkxz
        {
            get { return _dkxz; }
            set
            {
                _dkxz = value;
                this.RaisePropertyChanged("Dkxz");
            }
        }
        private string _dkbz;

        public string Dkbz
        {
            get { return _dkbz; }
            set
            {
                _dkbz = value;
                this.RaisePropertyChanged("Dkbz");
            }
        }
        private FieldModel _selectField;

        public FieldModel SelectField
        {
            get { return _selectField; }
            set
            {
                _selectField = value;
                this.RaisePropertyChanged("SelectField");
            }
        }

        private List<FieldModel> _filedList;

        public List<FieldModel> FieldList
        {
            get { return _filedList; }
            set
            {
                _filedList = value;
                this.RaisePropertyChanged("FieldList");
            }
        }        
        #endregion

        private readonly string _personDatabase;
        private readonly string _selectFeauture;
        private readonly IMap _pMap;

        public FieldsViewModel(string personDatabase,string selectFeature,IMap pMap)
        {
            _personDatabase = personDatabase;
            _selectFeauture = selectFeature;
            _pMap = pMap;
            InitGridList();
            InitCommands();
        }

        /// <summary>
        /// 填充Grid
        /// </summary>
        private void InitGridList()
        {
            IDatabaseService pDatabaseService=new MsAccessDatabase(_personDatabase);
            string sqlString = string.Format("Select OBJECTID,DKMC,CBFMC,DKBM,SCMJ,HTMJ,DKDZ,DKNZ,DKXZ,DKBZ,YHTMJ from {0}",
                _selectFeauture);
            var dt = pDatabaseService.Query(sqlString);
            if (dt == null) return;
            FieldList=new List<FieldModel>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                FieldList.Add(new FieldModel()
                {
                   Dkmc = dt.Rows[i][1].ToString(),
                   Cbfmc = dt.Rows[i][2].ToString(),
                   Dkbm = dt.Rows[i][3].ToString(),
                   Scmj = string.IsNullOrEmpty(dt.Rows[i][4].ToString().Trim())?
                   0.0 : double.Parse(dt.Rows[i][4].ToString().Trim()),
                   Htmj = string.IsNullOrEmpty(dt.Rows[i][5].ToString().Trim()) ?
                   0.0 : double.Parse(dt.Rows[i][5].ToString().Trim()),
                   Dkbz = dt.Rows[i][6].ToString(),
                   Dknz = dt.Rows[i][7].ToString(),
                   Dkxz = dt.Rows[i][8].ToString(),
                   Dkdz = dt.Rows[i][9].ToString(),
                   Yhtmj = string.IsNullOrEmpty(dt.Rows[i][10].ToString().Trim()) ?
                   0.0 : double.Parse(dt.Rows[i][10].ToString().Trim())
               
                });
            }
        }

        private void InitCommands()
        {
            MouseLeftButtonDownCommand = new DelegateCommand();
            WindowMoveCommand=new DelegateCommand();
            SelectChangedCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            MouseLeftButtonDownCommand.ExecuteAction = new Action<object>(CloseWindow);
            WindowMoveCommand.ExecuteAction=new Action<object>(MoveWindow);
            SelectChangedCommand.ExecuteAction=new Action<object>(SelectChangeField);
            ConfirmCommand.ExecuteAction=new Action<object>(Save);
        }

        #region 关闭和移动窗口

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

        #region 选择内容改变

        public DelegateCommand SelectChangedCommand { get; set; }

        private void SelectChangeField(object parameter)
        {
            if (SelectField == null) return;
            this.Dkmc = SelectField.Dkmc;
            this.Cbfmc = SelectField.Cbfmc;
            this.Dkbm = SelectField.Dkbm.Substring(14,5);
            this.Scmj = Convert.ToDouble(SelectField.Scmj.ToString("F"));
            this.Htmj = Convert.ToDouble(SelectField.Htmj.ToString("F"));
            this.Dkdz = SelectField.Dkdz;
            this.Dknz = SelectField.Dknz;
            this.Dkxz = SelectField.Dkxz;
            this.Dkbz = SelectField.Dknz;
            FreshMap(SelectField.Dkbm);
        }

        private void FreshMap(string dkbm)
        {
            IFeatureLayer pFeatureLayer = _pMap.get_Layer(0) as IFeatureLayer;
            IFeatureSelection pSection = pFeatureLayer as IFeatureSelection;
            IQueryFilter queryFilter = new SpatialFilterClass();
            queryFilter.WhereClause = string.Format("trim(DKBM) = {0}", dkbm);
            pSection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
            IAeFactory aeFactory = new PersonalGeoDatabase(_personDatabase);
            IFeatureClass pFeatureClass = aeFactory.OpenFeatureClasss(_selectFeauture);
            IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
            var fieldIndex = pFeatureClass.Fields.FindField("DKBM");
            IFeature pFeature;  
            while ((pFeature = pFeatureCursor.NextFeature()) != null)
            {
                if (pFeature.Value[fieldIndex].ToString() == dkbm)
                    break;
            }
            if (pFeature == null) return;
            var topo = pFeature.Shape as ITopologicalOperator;
            var topoGeometry = topo.Buffer(400);
            var pEnvelope = topoGeometry.Envelope;
            var pview = _pMap as IActiveView;
            if (pview == null) return;
            pview.Extent = pEnvelope;
            pview.Refresh();
            aeFactory.ReleaseFeautureClass(pFeatureClass);
        }
        #endregion

        #region 保存内容

        public DelegateCommand ConfirmCommand { get; set; }

        private void Save(object parameter)
        {
            if (SelectField == null) return;
            IDatabaseService pDatabaseService=new MsAccessDatabase(_personDatabase);
            var sqlString =
                string.Format(
                    "Update {0} Set DKMC='{1}',CBFMC='{2}',SCMJ={3},HTMJ={4},DKDZ='{5}',DKNZ='{6}',DKXZ='{7}',DKBZ='{8}' Where DKBM ='{9}'",
                    _selectFeauture, SelectField.Dkmc, SelectField.Cbfmc, SelectField.Scmj, SelectField.Htmj,
                    SelectField.Dkdz,
                    SelectField.Dknz, SelectField.Dkxz, SelectField.Dkbz,SelectField.Dkbm);
            var ret = pDatabaseService.Execute(sqlString);
            if (ret == -1)
            {
                MessageBox.Show(null, "更新失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show(null, "更新成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        #endregion


    }
    
}
