using System;
using System.Linq;
//using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
//using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Media.Media3D;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        //[MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context)
        {
            Patient patient = context.Patient;
            Course course = context.Course;

            PlanSetup planSetup = context.PlanSetup;

            //foreach (PlanSetup planSetup in course.PlanSetups)//.Where(pl => pl.RTPrescription != null))
            //{
            StructureSet PlanningStructureSet = planSetup.StructureSet;
            var PlanStructures = PlanningStructureSet.Structures;

            Beam beam1 = planSetup.Beams.Where(b => b.Meterset.Value > 0).FirstOrDefault();
            Point3D isoctr = GetIsocenter(beam1);

            var BODY_Y = new List<double>();
            var B_Y_Abs = new List<double>();
            var B_Y = new List<double>();
            var C_Y = new List<double>();
            double B_Y_Min = 0;
            double B_Y_Min_Non_Abs = 0;

            Structure BODY = null;
            Structure COUCH_SURFACE = null;

            foreach (Structure str in PlanStructures.Where(s => !s.IsEmpty))
            {
                switch (str.Id.ToLower().Equals("body"))
                {
                    case true:
                        BODY = str;
                        var bodyMesh = BODY.MeshGeometry;
                        Point3DCollection meshPositionsBODY = new Point3DCollection();
                        meshPositionsBODY = bodyMesh.Positions;

                        foreach (Point3D pointBODY in meshPositionsBODY)
                        {
                            BODY_Y.Add(pointBODY.Y);
                        }
                        break;
                    default:
                        switch (str.Id.ToLower().Equals("couchsurface"))
                        {
                            case true:
                                COUCH_SURFACE = str;
                                var couchsurfaceMesh = COUCH_SURFACE.MeshGeometry;
                                Point3DCollection meshPositionsCOUCH_SURFACE = new Point3DCollection();
                                meshPositionsCOUCH_SURFACE = couchsurfaceMesh.Positions;

                                foreach (Point3D pointCOUCH_SURFACE in meshPositionsCOUCH_SURFACE)
                                {
                                    B_Y_Abs.Add(Math.Abs(isoctr.Y - pointCOUCH_SURFACE.Y));
                                    B_Y.Add(isoctr.Y - pointCOUCH_SURFACE.Y);
                                }
                                B_Y_Min = B_Y_Abs.Min() / 10;
                                int B_Y_ABs_idx = B_Y_Abs.IndexOf(B_Y_Min * 10);
                                B_Y_Min_Non_Abs = B_Y[B_Y_ABs_idx];
                                break;
                        }
                        break;
                }
            }
            foreach (ReferencePoint rp in planSetup.ReferencePoints)
            {
                switch (rp.HasLocation(planSetup))
                {
                    case true:

                        if (rp.Id.ToLower().Contains("apex") || rp.Id.ToLower().Contains("right") || rp.Id.ToLower().Contains("left"))
                        {
                            VVector RefP = rp.GetReferencePointLocation(planSetup);
                            RefP = context.Image.DicomToUser(RefP, planSetup);
                            C_Y.Add(Math.Abs(RefP.y - (isoctr.Y - B_Y_Min_Non_Abs)));
                        }
                        break;
                }
            }
            if (C_Y.Count() < 1)
            {
                foreach (Structure str in PlanStructures.Where(s => !s.IsEmpty))
                {
                    if (str.Id.ToLower().EndsWith("apex") || str.Id.ToLower().EndsWith("right") || str.Id.ToLower().EndsWith("left"))
                    {
                        VVector RefP = str.CenterPoint;
                        RefP = context.Image.DicomToUser(RefP, planSetup);
                        C_Y.Add(Math.Abs(RefP.y - (isoctr.Y - B_Y_Min_Non_Abs)));
                    }
                }
            }

            double BODY_Y_Max = BODY_Y.Max();
            double BODY_Y_Min = BODY_Y.Min();
            double C_Y_Min = C_Y.Min() / 10;

            var BODY_Y_MaxMin = new List<double>
                {
                    Math.Abs(BODY_Y_Max - (isoctr.Y - B_Y_Min_Non_Abs)),
                    Math.Abs(BODY_Y_Min - (isoctr.Y - B_Y_Min_Non_Abs))
                };

            double Add_Min = BODY_Y_MaxMin.Min();

            double A = (Math.Abs((BODY_Y_Max - BODY_Y_Min) + Add_Min) / 10);
            //double A = Math.Abs(Math.Abs(BODY_Y_Min) - (isoctr.Y - B_Y_Min_Non_Abs)) / 10;

            string ptInfo = "Course ID: " + course.Id + " - Treatment Plan ID: " + planSetup.Id;

            MessageBox.Show(ptInfo + Environment.NewLine + Environment.NewLine
                + "Table surface to Farthest skin surface distance (A) = " + Environment.NewLine + "\t" + A.ToString("0.00") + " cm" + Environment.NewLine + Environment.NewLine 
                + "Table surface to isocenter or prostate / prostatic bed center distance (B) = " + Environment.NewLine + "\t" + B_Y_Min.ToString("0.00") + " cm" + Environment.NewLine + Environment.NewLine 
                + "Farthest Skin surface to prostate / prostatic bed or isocenter distance (A – B) = " + Environment.NewLine + "\t" + (A - B_Y_Min).ToString("0.00") + " cm" + Environment.NewLine + Environment.NewLine 
                + "Table surface to closest transponder distance (C) = " + Environment.NewLine + "\t" + (C_Y_Min).ToString("0.00") + " cm" + Environment.NewLine + Environment.NewLine 
                + "(Farthest) Skin surface to closest transponder distance (A – C) = " + Environment.NewLine + "\t" + (A - C_Y_Min).ToString("0.00") + " cm"
                , "Patient Name: " + patient.Name);
            //}
        }
        public static Point3D GetIsocenter(Beam beam)
        {
            Point3D iso = new Point3D
            {
                X = beam.IsocenterPosition.x,
                Y = beam.IsocenterPosition.y,
                Z = beam.IsocenterPosition.z
            };
            return iso;
        }
    }
}
