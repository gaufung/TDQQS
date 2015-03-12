using System;

using System.IO;

using System.Windows.Forms;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessor;
using TdqqClient.Services.AE;


namespace TdqqClient.Models.Edit
{
    class EditPoints:EditModel
    {
        public EditPoints(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
           // base.Edit(parameter);
            if (CreateJZD(PersonDatabase,SelectFeature,SelectFeature+"_JZD"))
            {
                MessageBox.Show(null, "提取界址点成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "提取界址点失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }       
        private  bool CreateJZD(string databaseUrl, string input, string output)
        {
            //DelectExcessVertex(databaseUrl,input);
            string tmpDir = AppDomain.CurrentDomain.BaseDirectory + "\\TMP";
            if (!Directory.Exists(tmpDir))
                Directory.CreateDirectory(tmpDir);
            IAeFactory pAeFactory=new PersonalGeoDatabase(databaseUrl);
            IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
            if (workspace != null)
            {
                try
                {
                    IFeatureClass inputFC = workspace.OpenFeatureClass(input);

                    // delIfExist(workspace, esriDatasetType.esriDTFeatureClass, output + "_T");

                    FeatureVerticesToPoints fvtp = new FeatureVerticesToPoints();
                    fvtp.in_features = inputFC;
                    fvtp.out_feature_class = tmpDir + "\\" + output + "_T.shp";
                    Geoprocessor GP = new Geoprocessor();
                    GP.OverwriteOutput = true;
                    GP.Execute(fvtp, null);
                    AddXY axy = new AddXY();
                    axy.in_features = tmpDir + "\\" + output + "_T.shp";

                    GP.Execute(axy, null);

                    Dissolve dlv = new Dissolve();
                    pAeFactory.DeleteIfExist(output);
                    dlv.dissolve_field = "POINT_X;POINT_Y";
                    dlv.multi_part = "SINGLE_PART";
                    dlv.in_features = tmpDir + "\\" + output + "_T.shp";
                    dlv.out_feature_class = databaseUrl + "\\" + output;
                    GP.Execute(dlv, null);
                    IFeatureClass outFC = workspace.OpenFeatureClass(output);
                    IFields fields = outFC.Fields;
                    int j = 0;
                    while (fields.FieldCount != j)
                    {
                        IField field = fields.get_Field(j);
                        if (field.Type != esriFieldType.esriFieldTypeOID && field.Type != esriFieldType.esriFieldTypeGeometry)
                        {
                            outFC.DeleteField(field);
                        }
                        else
                        {
                            j++;
                        }
                    }
                    var pField = new FieldClass();
                    var pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "BSM";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger;
                    pFieldEdit.Length_2 = 10;
                    outFC.AddField(pFieldEdit);

                    pField = new FieldClass();
                    pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "YSDM";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
                    pFieldEdit.Length_2 = 6;
                    outFC.AddField(pFieldEdit);

                    pField = new FieldClass();
                    pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "JZDH";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
                    pFieldEdit.Length_2 = 10;
                    outFC.AddField(pFieldEdit);

                    pField = new FieldClass();
                    pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "JZDLX";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
                    pFieldEdit.Length_2 = 1;
                    outFC.AddField(pFieldEdit);

                    pField = new FieldClass();
                    pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "JBLX";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
                    pFieldEdit.Length_2 = 1;
                    outFC.AddField(pFieldEdit);

                    IWorkspaceEdit workspaceEdit = workspace as IWorkspaceEdit;
                    workspaceEdit.StartEditing(false);
                    workspaceEdit.StartEditOperation();
                    IFeatureCursor featureCursor = outFC.Update(null, false);
                    IFeature feature = featureCursor.NextFeature();
                    while (feature != null)
                    {
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
                    return true;
                }
                catch (Exception ex)
                {
                                        
                }
            }
            return false;
        }

        }
        
}
