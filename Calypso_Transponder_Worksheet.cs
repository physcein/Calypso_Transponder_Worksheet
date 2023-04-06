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

            var beamList = new List<Beam>();
            //beamList.Add(planSetup.Beams.Where(b => b.Meterset.Value > 0).FirstOrDefault());
            beamList.Add(planSetup.Beams.FirstOrDefault());
            Point3D isoctr;

            int isoN = 0;

            //if (beamList[0] != null)
            //{
            isoctr = GetIsocenter(beamList[0]);
            VVector beamOri_ = beamList[0].IsocenterPosition;
            VVector beamOri = PlanningStructureSet.Image.DicomToUser(beamOri_, planSetup); 
            Point3D isoctr1 = GetTheIsocenter(beamOri);
            isoN = 1;
            //}
            //else
            //{
            //    VVector Origin = context.Image.DicomToUser(planSetup.StructureSet.Image.UserOrigin, planSetup);

            //    isoctr = new Point3D
            //    {
            //        X = Origin.x,
            //        Y = Origin.y,
            //        Z = Origin.z
            //    };
            //    isoN = 2;
            //}

            string iso_coord = "\n" + "Beam Isocenter Location " + " (cm) \n\t\t\t= \t(" + (isoctr1.X / 10).ToString("0.00") + ", " + (isoctr1.Y / 10).ToString("0.00") + ", " + (isoctr1.Z / 10).ToString("0.00") + ")" + Environment.NewLine;

            //MessageBox.Show(iso_coord);


            var BODY_Y = new List<double>();
            var B_Y_Abs = new List<double>();
            var B_Y = new List<double>();
            var C_Y = new List<double>();
            double B_Y_Min = 0;
            double B_Y_Min_Non_Abs = 0;

            Structure BODY = null;
            Structure COUCH_SURFACE = null;

            //string test = "";

            foreach (Structure str in PlanStructures.Where(s => !s.IsEmpty))
            {
                switch (str.Id.ToLower().Equals("body"))
                {
                    case true:
                        BODY = str;
                        var body_Mesh = BODY.MeshGeometry;
                        Point3DCollection meshPositions_BODY = new Point3DCollection();
                        meshPositions_BODY = body_Mesh.Positions;

                        foreach (Point3D pointBODY in meshPositions_BODY)
                        {
                            BODY_Y.Add(pointBODY.Y);
                        }
                        break;
                    default:
                        switch (str.Id.ToLower().Equals("couchsurface"))
                        {
                            case true:
                                COUCH_SURFACE = str;
                                var couchsurface_Mesh = COUCH_SURFACE.MeshGeometry;
                                Point3DCollection meshPositions_COUCH_SURFACE = new Point3DCollection();
                                meshPositions_COUCH_SURFACE = couchsurface_Mesh.Positions;

                                foreach (Point3D pointCOUCH_SURFACE in meshPositions_COUCH_SURFACE)
                                {
                                    B_Y_Abs.Add(Math.Abs(isoctr.Y - pointCOUCH_SURFACE.Y));
                                    B_Y.Add(isoctr.Y - pointCOUCH_SURFACE.Y);
                                }
                                int B_Y_ABs_idx = B_Y_Abs.IndexOf(B_Y_Abs.Min());
                                B_Y_Min = Math.Round(B_Y_Abs.Min() / 10, 2);
                                B_Y_Min_Non_Abs = B_Y[B_Y_ABs_idx];
                                break;
                        }
                        break;
                }
            }
            string Transponder_RP_Coord = "";
            //double rfpy = 0;
            foreach (ReferencePoint rp in planSetup.ReferencePoints.OrderBy(p => p.Id))
            {
                switch (rp.HasLocation(planSetup))
                {
                    case true:

                        if (rp.Id.ToLower().EndsWith("apex") || (rp.Id.ToLower().EndsWith("right") || rp.Id.ToLower().EndsWith("rt")) || (rp.Id.ToLower().EndsWith("left") || rp.Id.ToLower().EndsWith("lt")))
                        {
                            VVector RefP = rp.GetReferencePointLocation(planSetup);
                            VVector RefP1 = context.Image.DicomToUser(RefP, planSetup);
                            //if(rp.Id.ToLower().EndsWith("apex"))
                            //{
                            Transponder_RP_Coord += Environment.NewLine + "Transponder Reference Point Location" + Environment.NewLine;
                            //}
                            Transponder_RP_Coord += "\t" + rp.Id + " (cm) \t= \t(" + (RefP1.x / 10).ToString("0.00") + ", " + (RefP1.y / 10).ToString("0.00") + ", " + (RefP1.z / 10).ToString("0.00") + ")" + Environment.NewLine;

                            //RefP = context.Image.DicomToUser(RefP, planSetup);
                            //rfpy = RefP.y;
                            //MessageBox.Show(rfpy.ToString());

                            C_Y.Add(Math.Abs(RefP.y - (isoctr.Y - B_Y_Min_Non_Abs)));
                        }
                        break;
                }
            }

            string Transponder_Contour_Coord = "";
            if (C_Y.Count() < 1)
            {
                foreach (Structure str in PlanStructures.Where(s => !s.IsEmpty))
                {
                    if (str.Volume < 0.5 /*&& (str.Id.ToLower().StartsWith("trans") || str.Id.ToLower().StartsWith("tp"))*/ && (str.Id.ToLower().EndsWith("apex") || (str.Id.ToLower().EndsWith("right") || str.Id.ToLower().EndsWith("rt")) || (str.Id.ToLower().EndsWith("left") || str.Id.ToLower().EndsWith("lt"))))
                    {
                        VVector TPstrP = str.CenterPoint;
                        VVector TPstrP1 = context.Image.DicomToUser(TPstrP, planSetup);

                        //if (str.Id.ToLower().EndsWith("apex"))
                        //{
                        Transponder_Contour_Coord += "Transponder Contour CenterPoint Location" + Environment.NewLine;
                        //}
                        Transponder_Contour_Coord += "\t" + str.Id + " (cm) \t= \t(" + (TPstrP1.x / 10).ToString("0.00") + ", " + (TPstrP1.y / 10).ToString("0.00") + ", " + (TPstrP1.z / 10).ToString("0.00") + ")" + Environment.NewLine;

                        //TPstrP = context.Image.DicomToUser(TPstrP, planSetup);
                        C_Y.Add(Math.Abs(TPstrP.y - (isoctr.Y - B_Y_Min_Non_Abs)));
                        //test += (TPstrP.y - (isoctr.Y - B_Y_Min_Non_Abs)).ToString("0.00") + ", " + Math.Abs(TPstrP.y - (isoctr.Y - B_Y_Min_Non_Abs)).ToString("0.00") + "\n";
                    }
                }
            }
            //MessageBox.Show(test);

            double BODY_Y_Max = BODY_Y.Max();
            double BODY_Y_Min = BODY_Y.Min();

            var BODY_Y_MaxMin = new List<double>
                {
                    Math.Abs(BODY_Y_Max - (isoctr.Y - B_Y_Min_Non_Abs)),
                    Math.Abs(BODY_Y_Min - (isoctr.Y - B_Y_Min_Non_Abs))
                };

            double Add_Min = BODY_Y_MaxMin.Min();

            //double A = Math.Round(Math.Abs(BODY_Y_Max - BODY_Y_Min) / 10, 2);
            double A = Math.Round(Math.Abs((BODY_Y_Max - BODY_Y_Min) + Add_Min) / 10, 2);
            //B_Y_Min = Math.Round(B_Y_Min, 2);
            double C_Y_Min = Math.Round(C_Y.Min() / 10, 2);

            string ptInfo = "";

            if (isoN == 1)
            {
                ptInfo = "<><><><><><><><><><><><><><><><>" + Environment.NewLine + "Course ID: " + course.Id + Environment.NewLine + "Treatment Plan ID: " + planSetup.Id + Environment.NewLine + "<><><><><><><><><><><><><><><><>" + Environment.NewLine + "@ For \"BEAM ISOCENTER\" : " + Environment.NewLine + iso_coord + Transponder_RP_Coord + Transponder_Contour_Coord;

                MessageBox.Show(ptInfo + Environment.NewLine
                    + "Table Surface to Farthest Skin Surface Distance (A)" + Environment.NewLine + "\t      A = " + A + " cm" + Environment.NewLine + Environment.NewLine
                    + "Table Surface to Isocenter or Prostate / Prostatic Bed Center Distance (B)" + Environment.NewLine + "\t      B = " + B_Y_Min + " cm" + Environment.NewLine + Environment.NewLine
                    + "Farthest Skin Surface to Prostate / Prostatic Bed or Isocenter Distance (A – B)" + Environment.NewLine + "\tA – B = " + (A - B_Y_Min) + " cm" + Environment.NewLine + Environment.NewLine
                    + "Table Surface to Closest Transponder Distance (C)" + Environment.NewLine + "\t      C = " + (C_Y_Min) + " cm" + Environment.NewLine + Environment.NewLine
                    + "(Farthest) Skin Surface to Closest Transponder Distance (A – C)" + Environment.NewLine + "\tA – C = " + (A - C_Y_Min) + " cm"
                    , "Patient Name: " + patient.Name);
            }
            else if (isoN == 2)
            {
                ptInfo = "<><><><><><><><><><><><><><><><>" + Environment.NewLine + "Course ID: " + course.Id + Environment.NewLine + "Treatment Plan ID: " + planSetup.Id + Environment.NewLine + "<><><><><><><><><><><><><><><><>" + Environment.NewLine + "@ For \"USER ORIGIN\" : " + Environment.NewLine + iso_coord + Transponder_RP_Coord + Transponder_Contour_Coord;

                MessageBox.Show(ptInfo + Environment.NewLine
                    + "Table Surface to Farthest Skin Surface Distance (A)" + Environment.NewLine + "\t      A = " + A + " cm" + Environment.NewLine + Environment.NewLine
                    + "Table Surface to User Origin or Prostate / Prostatic Bed Center Distance (B)" + Environment.NewLine + "\t      B = " + B_Y_Min + " cm" + Environment.NewLine + Environment.NewLine
                    + "Farthest Skin Surface to Prostate / Prostatic Bed or User Origin Distance (A – B)" + Environment.NewLine + "\tA – B = " + (A - B_Y_Min) + " cm" + Environment.NewLine + Environment.NewLine
                    + "Table Surface to Closest Transponder Distance (C)" + Environment.NewLine + "\t      C = " + (C_Y_Min) + " cm" + Environment.NewLine + Environment.NewLine
                    + "(Farthest) Skin Surface to Closest Transponder Distance (A – C)" + Environment.NewLine + "\tA – C = " + (A - C_Y_Min) + " cm"
                    , "Patient Name: " + patient.Name);
            }
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

        public static Point3D GetTheIsocenter(VVector vector)
        {
            Point3D iso = new Point3D
            {
                X = vector.x,
                Y = vector.y,
                Z = vector.z
            };
            return iso;
        }
    }
}
