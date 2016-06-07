using System;
using System.Runtime.InteropServices;
using System.Windows;
using Inventor;
using Application = Inventor.Application;

namespace InventorAddIn1
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [Guid("f9132243-0bb0-41dc-8ce7-c2e0042bb5c5")]
    public class StandardAddInServer : ApplicationAddInServer
    {

        // Inventor application object.
        private Application mApp;
        private ButtonDefinition mButton;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            mApp = addInSiteObject.Application;

            // Create a button and connect it to the CreateSurface method.
            var ribbon = mApp.UserInterfaceManager.Ribbons["Part"];
            var tab = ribbon.RibbonTabs["id_TabTools"];
            var panel = tab.RibbonPanels.Add("Update", "ToolsTabUpdatePanel", "SampleClientId", "id_PanelP_ToolsMeasure");

            var cd = mApp.CommandManager.ControlDefinitions;
            mButton = cd.AddButtonDefinition("TestSurface", "TestSurface", CommandTypesEnum.kShapeEditCmdType,
                null, "TestSurface", "TestSurface", null, null);

            panel.CommandControls.AddButton(mButton);
            mButton.OnExecute += CreateSurface;
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

        public void CreateSurface(NameValueMap nv)
        {
            var T0 = DateTime.Now;
            var T1 = DateTime.Now;
            var T2 = DateTime.Now;

            var N = 100;
            var Nu = N;   // number of u control points
            var Lx = 10.0;  // x-extent in cm
            var Nv = N;   // number of v control points
            var Ly = 10.0;  // y-extent in cm

            var tg = mApp.TransientGeometry;
            var partDoc = (PartDocument)mApp.ActiveDocument;
            var cd = partDoc.ComponentDefinition;
            var features = cd.Features;

            int dimU = 3 + Nu;
            int dimV = 3 + Nv;
            int dimC = 3 * Nu * Nv;
            var uKnots = new double[dimU];
            var vKnots = new double[dimV];
            var controls = new double[dimC];
            // NOTE: The Inventor documentation does not seem to mention this, 
            // but providing an empty array of weights apparently causes 
            // Inventor to assume all points are equally weighted.
            //var weights = new double[Nu * Nv];
            var weights = new double[0];

            for (int i = 0; i <= 2; ++i)
            {
                uKnots[i] = vKnots[i] = 0.0;
            }

            for (int u = 3; u < Nu; ++u)
            {
                uKnots[u] = (u - 2) * (1.0f / (Nu - 2));
            }
            for (int v = 3; v < Nv; ++v)
            {
                vKnots[v] = (v - 2) * (1.0f / (Nv - 2));
            }

            for (int u = Nu; u < Nu + 3; ++u)
            {
                uKnots[u] = 1.0;
            }
            for (int v = Nv; v < Nv + 3; ++v)
            {
                vKnots[v] = 1.0;
            }

            for (int u = 0; u < Nu; ++u)
            {
                var x = u*Lx/(Nu - 1);
                for (int v = 0; v < Nv; ++v)
                {
                    var y = v * Ly / (Nv - 1);

                    var idx = 3 * (v * Nu + u);
                    controls[idx] = x;
                    controls[idx + 1] = y;
                    var r = Math.Sqrt(x*x + y*y);
                    controls[idx + 2] = 0.3 * Math.Cos(2 * r);
                }
            }

            var polesu0 = new double[3 * Nv]; // u = 0 edge
            var polesuN = new double[3 * Nv]; // u = Nu-1 edge
            var polesv0 = new double[3 * Nu]; // v = 0 edge
            var polesvN = new double[3 * Nu]; // v = Nv-1 edge
            for (int u = 0; u < Nu; ++u)
            {
                // The v0 edge runs from (0, 0) to (Nu - 1, 0)
                int v0base = 3 * u;
                // The vN edge runs from (Nu - 1, Nv - 1) to (0, Nv - 1)
                int vNbase = 3 * ((Nu - 1 - u) + Nu * (Nv - 1));
                for (int k = 0; k < 3; ++k)
                {
                    polesv0[3 * u + k] = controls[v0base + k];
                    polesvN[3 * u + k] = controls[vNbase + k];
                }
            }
            for (int v = 0; v < Nv; ++v)
            {
                // The u0 edge runs from (0, Nv - 1) to (0, 0)
                int u0base = 3 * Nu * (Nv - 1 - v);
                // The uN edge runs from (Nu - 1, 0) to (Nu - 1, Nv - 1)
                int uNbase = 3 * (Nu - 1 + Nu * v);
                for (int k = 0; k < 3; ++k)
                {
                    polesu0[3 * v + k] = controls[u0base + k];
                    polesuN[3 * v + k] = controls[uNbase + k];
                }
            }

            var order = new int[2] { 3, 3 };
            var isPeriodic = new bool[2] { false, false };

            var transaction = mApp.TransactionManager.StartTransaction(mApp.ActiveDocument, "Create B-Spline Surface");
            try
            {
                var curvev0 = tg.CreateBSplineCurve(3, ref polesv0, ref uKnots, ref weights, isPeriodic[0]);
                var curveuN = tg.CreateBSplineCurve(3, ref polesuN, ref vKnots, ref weights, isPeriodic[1]);

                // TODO: Should the knots be reversed for curvevN and curve u0?
                // We can get away with not reversing them now because they are always symmetric.
                var curvevN = tg.CreateBSplineCurve(3, ref polesvN, ref uKnots, ref weights, isPeriodic[0]);
                var curveu0 = tg.CreateBSplineCurve(3, ref polesu0, ref vKnots, ref weights, isPeriodic[1]);

                var topSurface = tg.CreateBSplineSurface(ref order, ref controls, ref uKnots, ref vKnots, ref weights, ref isPeriodic);

                if (topSurface == null)
                {
                    throw new Exception("TransientGeometry.CreateBSplineSurface returned null");
                }

                var bodyDef = mApp.TransientBRep.CreateSurfaceBodyDefinition();

                var corners = new VertexDefinition[4];
                corners[0] = bodyDef.VertexDefinitions.Add(tg.CreatePoint(controls[0], controls[1], controls[2]));
                int c1 = 3 * (Nu - 1);
                corners[1] = bodyDef.VertexDefinitions.Add(tg.CreatePoint(controls[c1], controls[c1 + 1], controls[c1 + 2]));
                int c2 = c1 + 3 * Nu * (Nv - 1);
                corners[2] = bodyDef.VertexDefinitions.Add(tg.CreatePoint(controls[c2], controls[c2 + 1], controls[c2 + 2]));
                int c3 = 3 * Nu * (Nv - 1);
                corners[3] = bodyDef.VertexDefinitions.Add(tg.CreatePoint(controls[c3], controls[c3 + 1], controls[c3 + 2]));

                var edges = new EdgeDefinition[4];
                edges[0] = bodyDef.EdgeDefinitions.Add(corners[0], corners[1], curvev0);
                edges[1] = bodyDef.EdgeDefinitions.Add(corners[1], corners[2], curveuN);
                edges[2] = bodyDef.EdgeDefinitions.Add(corners[2], corners[3], curvevN);
                edges[3] = bodyDef.EdgeDefinitions.Add(corners[3], corners[0], curveu0);

                var lumpDef = bodyDef.LumpDefinitions.Add();
                var shellDef = lumpDef.FaceShellDefinitions.Add();

                var topFace = shellDef.FaceDefinitions.Add(topSurface, false);
                // TODO: What is this ID for?
                topFace.AssociativeID = 501;

                var topFaceLoop = topFace.EdgeLoopDefinitions.Add();
                topFaceLoop.EdgeUseDefinitions.Add(edges[3], false);
                topFaceLoop.EdgeUseDefinitions.Add(edges[2], false);
                topFaceLoop.EdgeUseDefinitions.Add(edges[1], true);
                topFaceLoop.EdgeUseDefinitions.Add(edges[0], true);

                T1 = DateTime.Now;
                NameValueMap errors;
                var newBody = bodyDef.CreateTransientSurfaceBody(out errors);
                T2 = DateTime.Now;

                if (newBody == null)
                {
                    throw new Exception("SurfaceBodyDefinition.CreateTransientSurfaceBody returned null");
                }

                NonParametricBaseFeatureDefinition baseDef = features.NonParametricBaseFeatures.CreateDefinition();
                ObjectCollection objColl = mApp.TransientObjects.CreateObjectCollection();
                objColl.Add(newBody);
                baseDef.BRepEntities = objColl;
                baseDef.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType; 
                NonParametricBaseFeature baseFeature = features.NonParametricBaseFeatures.AddByDefinition(baseDef);

                transaction.End();
            }
            catch (Exception exc)
            {
                transaction.Abort();
            }

            var Tfinal = DateTime.Now;

            var d01 = T1 - T0;
            var msg01 = string.Format("Time before CreateTransientSurfaceBody: {0} sec", d01.TotalSeconds);
            var d12 = T2 - T1;
            var msg12 = string.Format("CreateTransientSurfaceBody time: {0} sec", d12.TotalSeconds);
            var d2f = Tfinal - T2;
            var msg2f = string.Format("Time after CreateTransientSurfaceBody: {0} sec", d2f.TotalSeconds);

            MessageBox.Show(msg01 + "\n" + msg12 + "\n" + msg2f);
        }
    }
}
