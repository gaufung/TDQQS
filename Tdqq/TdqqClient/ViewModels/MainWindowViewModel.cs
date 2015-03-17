using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessor;
using TdqqClient.Commands;
using TdqqClient.Models.Edit;
using TdqqClient.Models.Export.ExportOne;
using TdqqClient.Models.Export.ExportTotal;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Services.Common;
using TdqqClient.Views;
using MessageBox = System.Windows.Forms.MessageBox;

namespace TdqqClient.ViewModels
{
    partial class MainWindowViewModel:NotificationObject
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
        public Window ParentWindow { get; set; }
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

        private readonly  AxMapControl _axMapControl;
        public MainWindowViewModel(AxMapControl axMapControl)
        {
            IsClosed = false;
            IsOpen = false;
            _axMapControl = axMapControl;
            _isPointEdit = false;
            InitDelegateCommand();
            OpenDatabaseCommand.ExecuteAction = new Action<object>(Open);
            CloseMapCommand.ExecuteAction=new Action<object>(Close);
            EditFieldsCommand.ExecuteAction = new Action<object>(EditFields);
            SetDefaultCommand.ExecuteAction=new Action<object>(SetDefault);
            DkbmCommand.ExecuteAction=new Action<object>(SortDkbm);
            ScmjCommand.ExecuteAction=new Action<object>(SetScmj);
            HtmjCommand.ExecuteAction=new Action<object>(SetHtmj);
            UpdateCommand.ExecuteAction = new Action<object>(UnpdateCbfbm);
            ReplaceCommand.ExecuteAction = new Action<object>(RepalceCbfmc);
            DeleteValidPointCommand.ExecuteAction = new Action<object>(DeleatePoint);
            ValidTopoCommand.ExecuteAction = new Action<object>(ValidTopo);
            JzdCommand.ExecuteAction=new Action<object>(Jzd);
            JzxCommand.ExecuteAction=new Action<object>(Jzx);
            SzCommmand.ExecuteAction = new Action<object>(Sz);
            FieldsInfoCommand.ExecuteAction=new Action<object>(FieldsInfo);
            FarmerInfoCommand.ExecuteAction = new Action<object>(FarmerInfo);
            StartPointCommand.ExecuteAction = new Action<object>(ChangePointEditState);
            StopPointCommand.ExecuteAction = new Action<object>(ChangePointEditState);
            ExportACommand.ExecuteAction = new Action<object>(ExportA);
            ExportDCommand.ExecuteAction = new Action<object>(ExportD);
            ExportFamilyCommand.ExecuteAction = new Action<object>(ExportFamily);
            ExportOpenCommand.ExecuteAction = new Action<object>(ExportOpen);
            ExportSignCommand.ExecuteAction = new Action<object>(ExportSign);
            ExportListCommand.ExecuteAction = new Action<object>(ExportList);
            ExportPostCommand.ExecuteAction = new Action<object>(ExportPost);
            ExportDelegateCommand.ExecuteAction = new Action<object>(ExportDelegate);
            ExportJyqzCommand.ExecuteAction = new Action<object>(ExportJyqz);
            ExportConverCommand.ExecuteAction = new Action<object>(ExportCover);
            ExportCbfCommand.ExecuteAction = new Action<object>(ExportCbf);
            ExportMapCommand.ExecuteAction = new Action<object>(ExportMap);
            ExportContractCommand.ExecuteAction = new Action<object>(ExportContract);
            ExportStatementCommand.ExecuteAction = new Action<object>(ExportStatement);
            ExportAcceptCommand.ExecuteAction = new Action<object>(ExportAccept);
            ExportGhbCommand.ExecuteAction = new Action<object>(ExportGhb);
            ExportDkCommand.ExecuteAction=new Action<object>(ExportDk);
            ExportRegisterCommand.ExecuteAction = new Action<object>(ExportRegister);
            ExportArchiveCommand.ExecuteAction=new Action<object>(ExportArchive);
            ExportFarmerArchiveCommand.ExecuteAction=new Action<object>(ExportFarmerArchive);
            InitAxMapControlEvent();
        }

        private void InitDelegateCommand()
        {
            OpenDatabaseCommand=new DelegateCommand();
            CloseMapCommand=new DelegateCommand();
            EditFieldsCommand=new DelegateCommand();
            SetDefaultCommand=new DelegateCommand();
            DkbmCommand=new DelegateCommand();
            ScmjCommand=new DelegateCommand();
            HtmjCommand=new DelegateCommand();
            UpdateCommand=new DelegateCommand();
            ReplaceCommand=new DelegateCommand();
            DeleteValidPointCommand=new DelegateCommand();
            ValidTopoCommand=new DelegateCommand();
            JzdCommand=new DelegateCommand();
            JzxCommand=new DelegateCommand();
            SzCommmand = new DelegateCommand();
            FieldsInfoCommand=new DelegateCommand();
            FarmerInfoCommand=new DelegateCommand();
            StartPointCommand=new DelegateCommand();
            StopPointCommand=new DelegateCommand();
            ExportACommand = new DelegateCommand();
            ExportDCommand = new DelegateCommand();
            ExportFamilyCommand = new DelegateCommand();
            ExportOpenCommand = new DelegateCommand();
            ExportSignCommand = new DelegateCommand();
            ExportListCommand = new DelegateCommand();
            ExportPostCommand = new DelegateCommand();
            ExportDelegateCommand = new DelegateCommand();
            ExportJyqzCommand = new DelegateCommand();
            ExportConverCommand = new DelegateCommand();
            ExportCbfCommand = new DelegateCommand();
            ExportDkCommand = new DelegateCommand();
            ExportMapCommand = new DelegateCommand();
            ExportStatementCommand = new DelegateCommand();
            ExportAcceptCommand = new DelegateCommand();
            ExportContractCommand = new DelegateCommand();
            ExportGhbCommand=new DelegateCommand();
            ExportRegisterCommand = new DelegateCommand();
            ExportArchiveCommand=new DelegateCommand();
            ExportFarmerArchiveCommand=new DelegateCommand();
        }

        #region 打开和关闭地图

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
           var dialogHelper = new DialogHelper("mdb");           
            var ret = dialogHelper.OpenFile("请选择个人地理数据库");           
            if (string.IsNullOrEmpty(ret)) return;
            _personDatabase = ret;
            var selectFeatureVm = new SelectFeatureViewModel(_personDatabase);
            var selectFeatureV = new SelectFeatureWindow(selectFeatureVm);
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

        private bool Check(string personDatabase,string selectFeature)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeature);
            if (!pFeatureClass.CheckType(esriGeometryType.esriGeometryPolygon))
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
                if (!TopoCheck(personDatabase, selectFeature))
                {
                    System.Windows.Forms.MessageBox.Show(null,
                 "地块重叠", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
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
            Wait wait = new Wait();
            wait.SetWaitCaption("检查拓扑重叠");
            para["pd"] = personDatabase;
            para["sf"] = selectFeature;
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
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeature);
            var count = pFeatureClass.Count();
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
            IAeFactory pAeFactory = new PersonalGeoDatabase(_personDatabase);
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
        #endregion

        #region 编辑和设置默认值

        public DelegateCommand EditFieldsCommand { get; set; }
        public DelegateCommand SetDefaultCommand { get; set; }


        private void EditFields(object parameter)
        {
            
            EditModel edit=new EditFields(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }

        private void SetDefault(object parameter)
        {
            
            EditModel edit=new EditSetDefault(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }
        #endregion

        #region 设置地块编码

        public DelegateCommand DkbmCommand { get; set; }
        private void SortDkbm(object parameter)
        {
            
            EditModel edit=new EditDkbm(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }
        #endregion

        #region 设置合同面积

        public DelegateCommand HtmjCommand { get; set; }
        private void SetHtmj(object parameter)
        {
            EditModel edit=new EditHtmj(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }

        #endregion

        #region 设置实测面积

        public DelegateCommand ScmjCommand { get; set; }
        private void SetScmj(object parameter)
        {
           
            EditModel edit=new EditScmj(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);

        }
        #endregion

        #region 更新承包方编码

        public DelegateCommand UpdateCommand { get; set; }  
        private void UnpdateCbfbm(object parameter)
        {
            
            EditModel edit=new EditCbfbm(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }
        #endregion

        #region 替换承包方名称

        public DelegateCommand ReplaceCommand { get; set; }

        private void RepalceCbfmc(object parameter) 
        {
            EditModel edit=new EditCbfmc(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }
        #endregion

        #region 删除无效节点

        public DelegateCommand DeleteValidPointCommand { get; set; }
        private void DeleatePoint(object parameter)
        {
           
            EditModel edit=new EditInvalidatePoint(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }       
        
        #endregion

        #region 拓扑检验

        public DelegateCommand ValidTopoCommand { get; set; }

        private void ValidTopo(object parameter)
        {
            
            EditModel edit=new EditTopo(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
            
        }       
        #endregion

        #region 提取界址点

        public DelegateCommand JzdCommand { get; set; }

        private void Jzd(object parameter)
        {
            /*
            if (CreateJzd())
            {
                MessageBox.Show(null, "提取界址点成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "提取界址点失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
             */
            EditModel edit=new EditPoints(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
            
        }

        private bool CreateJzd()
        {
            Wait wait=new Wait();
            wait.SetWaitCaption("提取界址点");
            Hashtable para=new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t=new Thread(new ParameterizedThreadStart(CreateJzd));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }
        public  void CreateJzd(object p)
        {
            string tmpDir = AppDomain.CurrentDomain.BaseDirectory + "\\TMP";
            if (!Directory.Exists(tmpDir))
                Directory.CreateDirectory(tmpDir);
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            IAeFactory pAeFactory=new PersonalGeoDatabase(_personDatabase);
            IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
            string jzdFeature = _selectFeauture + "_JZD";
            IFeatureClass inputFC = pAeFactory.OpenFeatureClasss(_selectFeauture);
            try
            {
                FeatureVerticesToPoints fvtp = new FeatureVerticesToPoints();
                fvtp.in_features = inputFC;
                fvtp.out_feature_class = tmpDir + "\\" + jzdFeature + "_T.shp";
                Geoprocessor GP = new Geoprocessor();
                GP.OverwriteOutput = true;
                GP.Execute(fvtp, null);
                AddXY axy = new AddXY();
                axy.in_features = tmpDir + "\\" + jzdFeature + "_T.shp";
                GP.Execute(axy, null);
                Dissolve dlv = new Dissolve();
                pAeFactory.DeleteIfExist(jzdFeature);
                dlv.dissolve_field = "POINT_X;POINT_Y";
                dlv.multi_part = "SINGLE_PART";
                dlv.in_features = tmpDir + "\\" + jzdFeature + "_T.shp";
                dlv.out_feature_class = _personDatabase + "\\" + jzdFeature;
                GP.Execute(dlv, null);
                IFeatureClass outFC = pAeFactory.OpenFeatureClasss(jzdFeature);
                IFields fields = outFC.Fields;
                int j = 0;
                while (fields.FieldCount != j)
                {
                    IField field = fields.get_Field(j);
                    if (field.Type != esriFieldType.esriFieldTypeOID &&
                        field.Type != esriFieldType.esriFieldTypeGeometry)
                    {
                        outFC.DeleteField(field);
                    }
                    else
                    {
                        j++;
                    }
                }
                var addFields = JzdAddFields();
                foreach (var fieldEdit in addFields)
                {
                    outFC.AddField(fieldEdit);
                }
                int count = outFC.Count();
                int current = 0;
                IWorkspaceEdit workspaceEdit = workspace as IWorkspaceEdit;
                workspaceEdit.StartEditing(false);
                workspaceEdit.StartEditOperation();
                IFeatureCursor featureCursor = outFC.Update(null, false);
                IFeature feature = featureCursor.NextFeature();
                while (feature != null)
                {
                    wait.SetProgress((double) current++/(double) count);
                    feature.set_Value(2, feature.get_Value(0));
                    feature.set_Value(3, "211021");
                    feature.set_Value(4, feature.get_Value(0).ToString());
                    feature.set_Value(5, "3");
                    feature.set_Value(6, "9");
                    featureCursor.UpdateFeature(feature);
                    feature = featureCursor.NextFeature();
                }
                featureCursor.Flush();
                workspaceEdit.StopEditOperation();
                workspaceEdit.StopEditing(true);
                para["ret"] = true;

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                para["ret"] = false;
            }
            finally
            {
                pAeFactory.ReleaseFeautureClass(inputFC);
                wait.CloseWait();
            }
        }

        private IEnumerable<IFieldEdit> JzdAddFields()
        {
            List<IFieldEdit> addFields = new List<IFieldEdit>();
            var pField = new FieldClass();
            var pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "BSM";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger;
            pFieldEdit.Length_2 = 10;
            addFields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "YSDM";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 6;
            addFields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JZDH";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 10;
            addFields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JZDLX";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 1;
            addFields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JBLX";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 1;
            addFields.Add(pFieldEdit);
            return addFields;
        }
        #endregion

        #region 提取界址线

        public DelegateCommand JzxCommand { get; set; }

        private void Jzx(object parameter)
        {
           
            EditModel edit=new EditEdges(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }
        #endregion

        #region 提取四至

        public DelegateCommand SzCommmand { get; set; }

        private void Sz(object parameter)
        {
           
            EditModel edit=new EditSz(_personDatabase,_selectFeauture,_basicDatabase);
            edit.Edit(parameter);
        }

        #endregion

        #region 地块信息查询
        
        public DelegateCommand FieldsInfoCommand { get; set; }

        private void FieldsInfo(object parameter)
        {
            FieldsViewModel fieldsVm=new FieldsViewModel(_personDatabase,_selectFeauture,_axMapControl.Map);
            FieldsView fieldsV=new FieldsView(fieldsVm);
            fieldsV.Owner = ParentWindow;
            fieldsV.Show();
        }
        #endregion

        #region 农户信息查询

        public DelegateCommand FarmerInfoCommand { get; set; }

        private void FarmerInfo(object parameter)
        {
            FarmersViewModel farmersVm=new FarmersViewModel(_personDatabase,_selectFeauture,_axMapControl.Map);
            FarmersView farmerV=new FarmersView(farmersVm);
            farmerV.Owner = ParentWindow;
            farmerV.Show();

        }
        #endregion

        #region 点选地块操作

        #region 前期准备工作
        private bool _isPointEdit;
        private FieldView _PointEdit = null;
        public DelegateCommand StartPointCommand { get; set; }
        public DelegateCommand StopPointCommand { get; set; }

        private void ChangePointEditState(object parameter)
        {
            _isPointEdit = !_isPointEdit;
        }

        #endregion

        #region 点选事件函数

        private void InitAxMapControlEvent()
        {

            _axMapControl.OnMouseDown += (object sender, IMapControlEvents2_OnMouseDownEvent e)=>
            {
                //如果是按住滚轮，平移整个
                if (e.button == 4)
                {
                    _axMapControl.MousePointer = esriControlsMousePointer.esriPointerHand;
                    this._axMapControl.Pan();
                }
                this._axMapControl.MousePointer = esriControlsMousePointer.esriPointerDefault;
                if (_isPointEdit)
                {
                    this._axMapControl.MousePointer = esriControlsMousePointer.esriPointerHand;
                    IPoint pPoint = new PointClass();
                    pPoint.PutCoords(e.mapX, e.mapY);
                    var fieldVm = GetPointFeature(_axMapControl.Map, pPoint);
                    if (_PointEdit == null || _PointEdit.IsDisposed())
                    {
                        _PointEdit = new FieldView();
                        _PointEdit.SetDataContext(fieldVm);
                        _PointEdit.Owner = ParentWindow;
                        _PointEdit.Show();
                        //winPointEdit.Topmost = true;
                    }
                    else
                    {
                        _PointEdit.SetDataContext(fieldVm);
                    }
                   

                }
            };
        }
        public FieldViewModel GetPointFeature(IMap pMap, IPoint pPoint)
        {
            IFeatureLayer pFeatureLayer = pMap.get_Layer(0) as IFeatureLayer;
            IFeatureSelection pSection = pFeatureLayer as IFeatureSelection;
            ISpatialFilter pSpatialFilter = new SpatialFilterClass();
            pSpatialFilter.Geometry = pPoint;
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin;
            pSection.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
            IActiveView pActiveView = pMap as IActiveView;
            pActiveView.Refresh();
            ISelection selection = pMap.FeatureSelection;
            IEnumFeature pEnumFeature = (IEnumFeature)selection;
            FieldViewModel filedVm=new FieldViewModel(_personDatabase,_selectFeauture);
            var pfeature =pEnumFeature.Next();
            filedVm.Cbfmc = pfeature.Value[pfeature.Fields.FindField("CBFMC")].ToString();
            filedVm.Cbfbm = pfeature.Value[pfeature.Fields.FindField("CBFBM")].ToString();
            filedVm.Dkmc = pfeature.Value[pfeature.Fields.FindField("DKMC")].ToString();
            filedVm.Dkbm = pfeature.Value[pfeature.Fields.FindField("DKBM")].ToString();
            filedVm.Dkdz = pfeature.Value[pfeature.Fields.FindField("DKDZ")].ToString();
            filedVm.Dknz = pfeature.Value[pfeature.Fields.FindField("DKNZ")].ToString();
            filedVm.Dkxz = pfeature.Value[pfeature.Fields.FindField("DKXZ")].ToString();
            filedVm.Dkbz = pfeature.Value[pfeature.Fields.FindField("DKBZ")].ToString();
            try
            {
                filedVm.Yhtmj =
                    string.IsNullOrEmpty(pfeature.Value[pfeature.Fields.FindField("YHTMJ")].ToString().Trim())
                        ? 0.0
                        : Convert.ToDouble(pfeature.Value[pfeature.Fields.FindField("YHTMJ")].ToString().Trim());
                filedVm.Htmj = string.IsNullOrEmpty(pfeature.Value[pfeature.Fields.FindField("HTMJ")].ToString())
                    ? 0.0 : Convert.ToDouble(Convert.ToDouble(pfeature.Value[pfeature.Fields.FindField("HTMJ")].ToString()).ToString("f"));
                filedVm.Scmj = string.IsNullOrEmpty(pfeature.Value[pfeature.Fields.FindField("SCMJ")].ToString())
                    ? 0.0
                    : Convert.ToDouble(Convert.ToDouble(pfeature.Value[pfeature.Fields.FindField("SCMJ")].ToString()).ToString("f"));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw;
            }
           
            return filedVm;
        }

        #endregion
        #endregion

        #region 输出成果

        /// <summary>
        /// 导出发包方调查表
        /// </summary>
        public DelegateCommand ExportACommand { get; set; }

        private void ExportA(object parameter)
        {
            AExport export=new AExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }
        /// <summary>
        /// 导出地块信息表
        /// </summary>
        public DelegateCommand ExportDCommand { get; set; }

        private void ExportD(object parameter)
        {
            DExport export=new DExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }

        /// <summary>
        /// 导出家庭成员信息
        /// </summary>
        public DelegateCommand ExportFamilyCommand { get; set; }

        private void ExportFamily(object parameter)
        {
            FamilyExport export=new FamilyExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }
        /// <summary>
        /// 导出公示表
        /// </summary>
        public DelegateCommand ExportOpenCommand { get; set; }

        private void ExportOpen(object parameter)
        {
            OpenExport export=new OpenExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }
        /// <summary>
        /// 导出签字表
        /// </summary>
        public DelegateCommand ExportSignCommand { get; set; }

        private void ExportSign(object parameter)
        {
            SignExport export=new SignExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }

        
        /// <summary>
        /// 导出颁证清册
        /// </summary>
        public DelegateCommand ExportListCommand { get; set; }

        private void ExportList(object parameter)
        {
            ListExport export=new ListExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }
        /// <summary>
        /// 导出公示公告
        /// </summary>
        public DelegateCommand ExportPostCommand { get; set; }

        private void ExportPost(object parameter)
        {
            PostExport export=new PostExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }

        /// <summary>
        /// 导出委托书
        /// </summary>
        public DelegateCommand ExportDelegateCommand { get; set; }

        private void ExportDelegate(object parameter)
        {
            DelegateExport export=new DelegateExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export(parameter);
        }

        public DelegateCommand ExportJyqzCommand { get; set; }

        private void ExportJyqz(object parameter)
        {
            JyqzsExport export=new JyqzsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportConverCommand { get; set; }

        private void ExportCover(object parameter)
        {
            CoversExport export=new CoversExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportCbfCommand { get; set; }

        private void ExportCbf(object parameter)
        {
            CbfsExport export=new CbfsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportDkCommand  { get; set; }

        private void ExportDk(object parameter)
        {
            DksExport export=new DksExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportContractCommand { get; set; }

        private void ExportContract(object parameter)
        {
            ContractsExport export=new ContractsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportStatementCommand { get; set; }

        private void ExportStatement(object parameter)
        {
            StatementsExport export=new StatementsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportAcceptCommand { get; set; }

        private void ExportAccept(object parameter)
        {
            AcceptsExport export=new AcceptsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportMapCommand { get; set; }

        private void ExportMap(object parameter)
        {
            MapsExport export=new MapsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportGhbCommand { get; set; }

        private void ExportGhb(object parameter)
        {
            GhbsExport export=new GhbsExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportRegisterCommand { get; set; }

        private void ExportRegister(object parameter)
        {
            RegistersExport export=new RegistersExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }

        public DelegateCommand ExportArchiveCommand { get; set; }

        private void ExportArchive(object parameter)
        {
            ArchivesExport export=new ArchivesExport(_personDatabase,_selectFeauture,_basicDatabase);
            export.Export();
        }
        public DelegateCommand ExportFarmerArchiveCommand { get; set; }

        private void ExportFarmerArchive(object parameter)
        {
            ExportViewModel exportVm=new ExportViewModel(_personDatabase,_selectFeauture,_basicDatabase);
            ExportView exportV=new ExportView(exportVm);
            exportV.ShowDialog();
        }
        #endregion
    }
}
