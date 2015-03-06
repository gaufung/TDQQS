using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using TdqqClient.Commands;
using TdqqClient.Services.AE;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.ViewModels
{
    class MainWindowViewModel:NotificationObject
    {
        #region 地图开关

        private bool _isClosed;

        public bool IsClosed
        {
            get { return _isClosed; }
            set
            {
                _isClosed = value;
                this.RaisePropertyChanged("IsClosed");
            }
        }
        private bool _isOpen;

        public bool IsOpen
        {
            get { return _isOpen; }
            set
            {
                _isOpen = value;
                this.RaisePropertyChanged("IsOpen");
            }
        } 
        #endregion

        #region 个人地理数据库位置和选择的要素类
        private string _personDatabase;
        private string _selectFeauture;
        private string _basicDatabase;
        #endregion

        public MainWindowViewModel()
        {
            IsClosed = false;
            IsOpen = false;
        }

        private AxMapControl _axMapControl;
        public MainWindowViewModel(AxMapControl axMapControl)
        {
            IsClosed = false;
            IsOpen = false;
            _axMapControl = axMapControl;
            InitDelegateCommand();
            OpenDatabaseCommand.ExecuteAction = new Action<object>(Open);
            CloseMapCommand.ExecuteAction=new Action<object>(Close);
        }

        private void InitDelegateCommand()
        {
            OpenDatabaseCommand=new DelegateCommand();
            CloseMapCommand=new DelegateCommand();
        }
        /// <summary>
        /// 打开地图操作
        /// </summary>
        public DelegateCommand OpenDatabaseCommand { get; set; }

        /// <summary>
        /// 关闭地图操作
        /// </summary>
        public DelegateCommand CloseMapCommand { get; set; }

        private void Open(object parameter)
        {
            var dialogHelper=new DialogHelper("mdb");
            var ret=dialogHelper.OpenFile("请选择个人地理数据库");
            if (string.IsNullOrEmpty(ret)) return;  
            _personDatabase = ret;
            var selectFeatureVm=new SelectFeatureWindowViewModel(_personDatabase);
            var selectFeatureV=new SelectFeatureWindow(selectFeatureVm);
            selectFeatureV.ShowDialog();
            ret = selectFeatureVm.SelectFeature;
            if (string.IsNullOrEmpty(ret)) return;
            _selectFeauture = ret;
            if (Check(_personDatabase, _selectFeauture))
            {
                if (!LoadMap())
                {
                    _personDatabase = _selectFeauture = string.Empty;
                }
                else
                {
                    _basicDatabase = CopyBasicDatabaseIfNotExist(_personDatabase);
                }
            }
            else
            {
                _personDatabase = _selectFeauture = string.Empty;
            }
            //检查Button是否可用
            ButtonStateCheck();
        }

        #region 数据检查

        private bool Check(string personDatabase, string selectFeature)
        {
            if (!TypeCheck(personDatabase, selectFeature, esriGeometryType.esriGeometryPolygon))
            {
                System.Windows.Forms.MessageBox.Show(null,
                    "打开的非地块要素类", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!NullValueCheck(personDatabase, selectFeature, "SHAPE_Length"))
            {
                System.Windows.Forms.MessageBox.Show(null,
                   "存在空地块", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (MessageBox.Show(null, "是否进行拓扑检查", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (!TopoCheck(personDatabase, selectFeature)) return false;
            }
            return true;
        }

        /// <summary>
        /// 检查要素类是否满足要求
        /// </summary>
        /// <param name="personDatabse"></param>
        /// <param name="selectFeature"></param>
        /// <param name="toCheckEsriGeometryType">要检查的类型</param>
        /// <returns>返回是否成功</returns>
        private bool TypeCheck(string personDatabse, string selectFeature, esriGeometryType toCheckEsriGeometryType)
        {
            return Services.Check.ValidCheck.CheckFeatureClassType(personDatabse, selectFeature,
                toCheckEsriGeometryType);
        }
        /// <summary>
        /// 检查是否存在空字段
        /// </summary>
        /// <param name="personDatabase"></param>
        /// <param name="selecreFeature"></param>
        /// <param name="toCheckField">要检查的字段</param>
        /// <returns>是否通过</returns>
        private bool NullValueCheck(string personDatabase, string selecreFeature, string toCheckField)
        {
            return
                Services.Check.ValidCheck.PersonDatabaseNullField(personDatabase, selecreFeature, "SHAPE_Length");
        }

        private bool TopoCheck(string personDatabase, string selectFeature)
        {
            Hashtable para = new Hashtable();
            var count = AeHelper.FeautureCount(personDatabase, selectFeature);
            if (count == -1) return false;
            Wait wait = new Wait();
            wait.SetWaitCaption("检查拓扑重叠");
            para["pd"] = personDatabase;
            para["sf"] = selectFeature;
            para["count"] = count;
            para["w"] = wait;
            para["ret"] = false;
            Thread t = new Thread(new ParameterizedThreadStart(TopoCheck));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];

        }
        /// <summary>
        /// 线程调用
        /// </summary>
        /// <param name="p"></param>
        private void TopoCheck(object p)
        {
            Hashtable para = p as Hashtable;
            var personDabase = para["pd"].ToString();
            var selectFeature = para["sf"].ToString();
            var wait = para["w"] as Wait;
            var count = (int)para["count"];
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeature);
            try
            {
                IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
                IFeature pFeature;
                int currentIndex = 0;
                while ((pFeature = pFeatureCursor.NextFeature()) != null)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)count));
                    var topoGeometry = pFeature.Shape;
                    ISpatialFilter pSpatialFilter = new SpatialFilterClass();
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelOverlaps;
                    pSpatialFilter.Geometry = topoGeometry;
                    IFeatureCursor mFeatureCursor = pFeatureClass.Search(pSpatialFilter, false);
                    IFeature feature = mFeatureCursor.NextFeature();
                    if (feature != null)
                    {
                        para["ret"] = false;
                        wait.CloseWait();
                        return;
                    }
                    Marshal.ReleaseComObject(mFeatureCursor);
                }
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = true;
            }
            finally
            {
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
                wait.CloseWait();
            }

        } 
        #endregion

        /// <summary>
        /// 按钮的可用性检查
        /// </summary>
        private void ButtonStateCheck()
        {
            if (_personDatabase == string.Empty || _selectFeauture == string.Empty)
            {
                IsOpen = false;
                IsClosed = true;
            }
            else
            {
                IsOpen = true;
                IsClosed = false;
            }
        }

        private bool LoadMap()
        {
            IAeFactory pAeFactory=new PersonalGeoDatabase(_personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(_selectFeauture);
            bool flag;
            try
            {
                ISimpleFillSymbol pSimpleFillSymbol = new SimpleFillSymbolClass();
                pSimpleFillSymbol.Style = esriSimpleFillStyle.esriSFSSolid;
                pSimpleFillSymbol.Color = AeHelper.GetRgb(180, 180, 0);
                var simpleRender = new ESRI.ArcGIS.Carto.SimpleRendererClass();
                simpleRender.Symbol = pSimpleFillSymbol as ISymbol;
                var featureLayer = new FeatureLayerClass();
                featureLayer.FeatureClass = pFeatureClass;
                featureLayer.Renderer = simpleRender;
                _axMapControl.Map.AddLayer(featureLayer);
                flag = true;

            }
            catch (Exception)
            {
                flag = false;
                System.Windows.Forms.MessageBox.Show(null,
                   "地图加载失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
            }
            return flag;
        }

        /// <summary>
        /// 复制家庭成员基础信息，如果不存在的话
        /// </summary>
        /// <param name="personDatabase">个人地理数据库的位置</param>
        /// <returns>返回家庭成员库的地址</returns>
        private string CopyBasicDatabaseIfNotExist(string personDatabase)
        {
            int floderIndex = personDatabase.LastIndexOf('\\');
            int nameIndex = personDatabase.LastIndexOf('.');
            string floderPath = personDatabase.Substring(0, floderIndex + 1);
            string fileName = personDatabase.Substring(floderIndex + 1, nameIndex - floderIndex - 1);
            string templateBasicDatabase = AppDomain.CurrentDomain.BaseDirectory + @"\template\基础数据模板.mdb";
            string currentBasicDatabasePath = floderPath + fileName + "_基础数据库.mdb";
            if (!File.Exists(currentBasicDatabasePath))
            {
                File.Copy(templateBasicDatabase, currentBasicDatabasePath);
            }
            return currentBasicDatabasePath;
        }
        private void Close(object parameter)
        {
            while (_axMapControl.Map.LayerCount != 0)
            {
                var pLayer = _axMapControl.Map.Layer[0];
                _axMapControl.Map.DeleteLayer(pLayer);
            }
            _personDatabase = _basicDatabase = _selectFeauture = string.Empty;
            ButtonStateCheck();
        }

    }
}
