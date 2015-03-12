using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessor;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    class EditEdges:EditModel
    {
        public EditEdges(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            if (CreateJzx())
            {
                MessageBox.Show(null, "界址线提取成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "界址线提取失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CreateJzx()
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("提取界址线");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(CreateJzx));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }
        public void CreateJzx(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
            string jzxFeauture = SelectFeature + "_JZX";
            try
            {

                IFeatureClass inputFC = workspace.OpenFeatureClass(SelectFeature);
                string outputT = PersonDatabase + "\\T_" + jzxFeauture;
                pAeFactory.DeleteIfExist("T_" + jzxFeauture);
                PolygonToLine ptl = new PolygonToLine();
                ptl.in_features = inputFC;
                ptl.out_feature_class = outputT;
                Geoprocessor GP = new Geoprocessor();
                GP.ResetEnvironments();
                GP.OverwriteOutput = true;
                GP.Execute(ptl, null);
                pAeFactory.DeleteIfExist(jzxFeauture);
                SplitLine sl = new SplitLine();
                sl.in_features = outputT;
                sl.out_feature_class = PersonDatabase + "\\" + jzxFeauture;
                GP.ResetEnvironments();
                GP.OverwriteOutput = true;
                GP.Execute(sl, null);
                pAeFactory.DeleteIfExist("T_" + jzxFeauture);
                IFeatureClass outFC = workspace.OpenFeatureClass(jzxFeauture);

                var fields = JzxFieldEdits();
                foreach (var fieldEdit in fields)
                {
                    outFC.AddField(fieldEdit);
                }

                int leftFidIndex = outFC.FindField("LEFT_FID");
                int rightFidIndex = outFC.FindField("RIGHT_FID");
                int cbfmcIndex = inputFC.FindField("CBFMC");
                int zjrIndex = inputFC.FindField("ZJRXM");
                IWorkspaceEdit workspaceEdit = workspace as IWorkspaceEdit;
                workspaceEdit.StartEditing(false);
                workspaceEdit.StartEditOperation();

                IFeatureCursor featureCursor = outFC.Update(null, false);
                IFeature feature = featureCursor.NextFeature();
                IFeature featureTmp = null;
                int left_fid = -1;
                int right_fid = -1;
                string left_pldwqlr = "";
                string right_pldwqlr = "";
                string left_pldwzjr = "";
                string right_pldwzjr = "";
                int count = outFC.Count();
                int current = 0;
                while (feature != null)
                {
                    wait.SetProgress((double)current++ / (double)count);
                    left_fid = (int)feature.get_Value(leftFidIndex);
                    right_fid = (int)feature.get_Value(rightFidIndex);
                    if (left_fid < 0)
                    {
                        left_pldwqlr = "";
                        left_pldwzjr = "";
                    }
                    else
                    {
                        featureTmp = inputFC.GetFeature(left_fid);
                        left_pldwqlr = featureTmp.get_Value(cbfmcIndex) as string;
                        left_pldwzjr = featureTmp.get_Value(zjrIndex) as string;
                    }
                    if (right_fid < 0)
                    {
                        right_pldwqlr = "";
                        right_pldwzjr = "";
                    }
                    else
                    {
                        featureTmp = inputFC.GetFeature(right_fid);
                        right_pldwqlr = featureTmp.get_Value(cbfmcIndex) as string;
                        right_pldwzjr = featureTmp.get_Value(zjrIndex) as string;
                    }


                    feature.set_Value(5, feature.OID);
                    feature.set_Value(6, "211031");
                    feature.set_Value(7, "600001");
                    if (left_fid < 0 || right_fid < 0)
                    {
                        feature.set_Value(8, "03");
                        feature.set_Value(9, "1");
                    }
                    else
                    {
                        feature.set_Value(8, "01");
                        feature.set_Value(9, "2");
                    }
                    feature.set_Value(10, "");

                    feature.set_Value(11, left_pldwqlr + "," + right_pldwqlr);
                    feature.set_Value(12, right_pldwzjr + "," + right_pldwzjr);

                    featureCursor.UpdateFeature(feature);
                    feature = featureCursor.NextFeature();
                }
                featureCursor.Flush();
                workspaceEdit.StopEditOperation();
                workspaceEdit.StopEditing(true);

                outFC.DeleteField(outFC.Fields.get_Field(outFC.FindField("LEFT_FID")));
                outFC.DeleteField(outFC.Fields.get_Field(outFC.FindField("RIGHT_FID")));
                para["ret"] = true;
            }
            catch (Exception ex)
            {
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }
        }

        private IEnumerable<IFieldEdit> JzxFieldEdits()
        {
            List<IFieldEdit> fields = new List<IFieldEdit>();
            var pField = new FieldClass();
            var pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "BSM";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger;
            pFieldEdit.Length_2 = 10;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "YSDM";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 6;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JXXZ";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 6;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JZXLB";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 6;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JZXWZ";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 1;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "JZXSM";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 300;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "PLDWQLR";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 100;
            fields.Add(pFieldEdit);

            pField = new FieldClass();
            pFieldEdit = pField as IFieldEdit;
            pFieldEdit.Name_2 = "PLDWZJR";
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString;
            pFieldEdit.Length_2 = 100;
            fields.Add(pFieldEdit);
            return fields;
        }   
    }
}
