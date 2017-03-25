using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using Inventor;
using Application = Inventor.Application;

namespace interactive_mesh
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("0b3929a0-ba85-4215-83d8-844c130e6d0f")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        // Inventor application object.
        private Application mApp;
        private ButtonDefinition mButton;

        private ClientGraphics mClientGraphics;
        private GraphicsDataSets mGraphicsDataSets;
        private GraphicsNode mGraphicsNode;

        private int N;
        private int mNumTriangles;
        private double[] mVertices;
        private int[] mIndices;
        private double[] mNormals;

        public StandardAddInServer()
        {
            N = 512;
            mNumTriangles = 2 * (N - 1) * (N - 1);

            mVertices = new double[3 * N * N];
            for (int i = 0, k = 0; i < N; ++i)
            {
                var x = i * 10.0 / (N - 1);
                for (int j = 0; j < N; ++j)
                {
                    var y = j * 10.0 / (N - 1);
                    var r = Math.Sqrt(x * x + y * y);
                    var z = 0.3 * Math.Sin(2 * Math.PI * r);

                    mVertices[k++] = x;
                    mVertices[k++] = y;
                    mVertices[k++] = z;
                }
            }

            mIndices = new int[3 * mNumTriangles];
            for (int r = 0, k = 0; r < N - 1; ++r)
            {
                var i0 = r * N;
                var i1 = i0 + N;
                for (int c = 0; c < N - 1; ++c)
                {
                    var t0 = i0 + c;
                    var t1 = i1 + c;

                    // NOTE: The extra +1 on each line is because Inventor expects 1-based indices.
                    mIndices[k++] = t0 + 1;
                    mIndices[k++] = t0 + 1 + 1;
                    mIndices[k++] = t1 + 1 + 1;

                    mIndices[k++] = t0 + 1;
                    mIndices[k++] = t1 + 1 + 1;
                    mIndices[k++] = t1 + 1;
                }
            }

            mNormals = new double[3 * mNumTriangles];
            for (int t = 0; t < mNumTriangles; ++t)
            {
                mNormals[3 * t] = 0;
                mNormals[3 * t + 1] = 0;
                mNormals[3 * t + 2] = 1;
            }

        }

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            mApp = addInSiteObject.Application;

            // Create a button and connect it to the CreateMesh method.
            var ribbon = mApp.UserInterfaceManager.Ribbons["Part"];
            var tab = ribbon.RibbonTabs["id_TabTools"];
            var panel = tab.RibbonPanels.Add("Update", "ToolsTabUpdatePanel", "SampleClientId", "id_PanelP_ToolsMeasure");

            var cd = mApp.CommandManager.ControlDefinitions;
            mButton = cd.AddButtonDefinition("TestMesh", "TestMesh", CommandTypesEnum.kShapeEditCmdType,
                null, "TestMesh", "TestMesh", null, null);

            panel.CommandControls.AddButton(mButton);
            mButton.OnExecute += CreateMesh;
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            mApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        public void CreateMesh(NameValueMap nv)
        {
            var N = 512;
            var T0 = DateTime.Now;

            if (null != mGraphicsDataSets)
            {
                mGraphicsDataSets.Delete();
                mGraphicsDataSets = null;

                mGraphicsNode.Delete();
                mGraphicsNode = null;

                mApp.ActiveView.Update();
                return;
            }

            var doc = mApp.ActiveDocument as PartDocument;
            if (null == mClientGraphics)
            {
                mClientGraphics = doc.ComponentDefinition.ClientGraphicsCollection.Add("MyTest");
            }

            mGraphicsDataSets = doc.GraphicsDataSetsCollection.Add("MyTest");

            var setupTimeSeconds = (DateTime.Now - T0).TotalSeconds;
            T0 = DateTime.Now;


            var msg = string.Format("N = {0}, {1} triangles\n", N, mNumTriangles)
                      + string.Format("Inventor setup time: {0} sec\n", setupTimeSeconds);

            T0 = DateTime.Now;

            //var transaction = mApp.TransactionManager.StartTransaction(mApp.ActiveDocument, "test mesh");

            try
            {
                var dataSetsTimeSeconds = (DateTime.Now - T0).TotalSeconds;
                T0 = DateTime.Now;

                var gcs = mGraphicsDataSets.CreateCoordinateSet(1) as GraphicsCoordinateSet;
                gcs.PutCoordinates(mVertices);

                var coordSetSeconds = (DateTime.Now - T0).TotalSeconds;
                T0 = DateTime.Now;

                var gis = mGraphicsDataSets.CreateIndexSet(2) as GraphicsIndexSet;
                gis.PutIndices(mIndices);

                var indexSetSeconds = (DateTime.Now - T0).TotalSeconds;
                T0 = DateTime.Now;

                var gns = mGraphicsDataSets.CreateNormalSet(3) as GraphicsNormalSet;
                gns.PutNormals(mNormals);

                var normalSetSeconds = (DateTime.Now - T0).TotalSeconds;
                T0 = DateTime.Now;

                mGraphicsNode = mClientGraphics.AddNode(1) as GraphicsNode;
                var triangles = mGraphicsNode.AddTriangleGraphics() as TriangleGraphics;

                triangles.CoordinateSet = gcs;
                triangles.CoordinateIndexSet = gis;
                triangles.NormalSet = gns;
                triangles.NormalBinding = NormalBindingEnum.kPerItemNormals;
                triangles.NormalIndexSet = gis;
                //triangles.ColorSet = oGraphicsColorSet;
                //triangles.ColorBinding = ColorBindingEnum.kPerVertexColors;
                //triangles.ColorIndexSet = oGraphicsIndexSet;

                var trianglesSeconds = (DateTime.Now - T0).TotalSeconds;
                T0 = DateTime.Now;

                //transaction.End();

                msg = msg + string.Format("GraphicsDataSetsCollection.Add time: {0} sec\n", dataSetsTimeSeconds)
                    + string.Format("Coordinate set creation time: {0} sec\n", coordSetSeconds)
                    + string.Format("Index set creation time: {0} sec\n", indexSetSeconds)
                    + string.Format("Normal set creation time: {0} sec\n", normalSetSeconds)
                    + string.Format("Triangle node creation time: {0} sec\n", trianglesSeconds);

            }
            catch (Exception e)
            {
                msg += string.Format("Exception: {0}", e.ToString());
            }
            finally
            {
                //transaction.Abort();
            }

            mApp.ActiveView.Update();

            System.Windows.MessageBox.Show(msg);



        }
    }
}
