using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(106)]
    public class CloudCalloutType : ShapeType
    {
        public CloudCalloutType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.round;


            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.round;

            this.Path = "ar,7165,4345,13110,1950,7185,1080,12690,475,11732,4835,17650,1080,12690,2910,17640,2387,9757,10107,20300,2910,17640,8235,19545,766,12382,14412,21597,8235,19545,14280,18330,12910,11080,18695,18947,14280,18330,18690,15045,14822,5862,21597,15082,18690,15045,2095,7665,15772,2592,21105,9865,20895,7665,19140,2715,14330,,19187,6595,19140,2715,14910,1170,10992,,15357,5945,14910,1170,11250,1665,6692,650,12025,7917,11250,1665,7005,2580,1912,1972,8665,11162,7005,2580,1950,7185xear,7165,4345,13110,1080,12690,2340,13080nfear475,11732,4835,17650,2910,17640,3465,17445nfear7660,12382,14412,21597,7905,18675,823519545nfear7660,12382,14412,21597,14280,18330,14400,17370nfear12910,11080,18695,18947,18690,15045,17070,11475nfear15772,2592,2115,9865,20175,9015,20895,7665nfear14330,,19187,6595,19200,3345,19140,2715nfear14330,,19187,6595,14910,1170,14550,1980nfear10992,,15357,5945,11250,1665,11040,2340nfear1912,1972,8665,11162,7650,3270,7005,2580nfear1912,1972,8665,11162,1950,7185,2070,7890nfem@23@37qx@35@24@23@36@34@24@23@37xem@16@33qx@31@17@16@32@30@17@16@33xem@38@29qx@27@39@38@28@26@39@38@29xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("sum #0 0 10800");
            this.Formulas.Add("sum #1 0 10800");
            this.Formulas.Add("cosatan2 10800 @0 @1");
            this.Formulas.Add("sinatan2 10800 @0 @1");
            this.Formulas.Add("sum @2 10800 0");
            this.Formulas.Add("sum @3 10800 0");
            this.Formulas.Add("sum @4 0 #0");
            this.Formulas.Add("sum @5 0 #1");
            this.Formulas.Add("mod @6 @7 0");
            this.Formulas.Add("prod 600 11 1");
            this.Formulas.Add("sum @8 0 @9");
            this.Formulas.Add("prod @10 1 3");
            this.Formulas.Add("prod 600 3 1");
            this.Formulas.Add("sum @11 @12 0");
            this.Formulas.Add("prod @13 @6 @8");
            this.Formulas.Add("prod @13 @7 @8");
            this.Formulas.Add("sum @14 #0 0");
            this.Formulas.Add("sum @15 #1 0");
            this.Formulas.Add("prod 600 8 1");
            this.Formulas.Add("prod @11 2 1");
            this.Formulas.Add("sum @18 @19 0");
            this.Formulas.Add("prod @20 @6 @8");
            this.Formulas.Add("prod @20 @7 @8");
            this.Formulas.Add("sum @21 #0 0");
            this.Formulas.Add("sum @22 #1 0");
            this.Formulas.Add("prod 600 2 1");
            this.Formulas.Add("sum #0 600 0");
            this.Formulas.Add("sum #0 0 600");
            this.Formulas.Add("sum #1 600 0");
            this.Formulas.Add("sum #1 0 600");
            this.Formulas.Add("sum @16 @25 0");
            this.Formulas.Add("sum @16 0 @25");
            this.Formulas.Add("sum @17 @25 0");
            this.Formulas.Add("sum @17 0 @25");
            this.Formulas.Add("sum @23 @12 0");
            this.Formulas.Add("sum @23 0 @12");
            this.Formulas.Add("sum @24 @12 0");
            this.Formulas.Add("sum @24 0 @12");
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");

            this.AdjustmentValues = "1350,25920";

            this.ConnectorLocations = "67,10800;10800,21577;21582,10800;10800,1235;@38,@39";

            this.TextboxRectangle = "2977,3262,17087,17337";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1"
            };
            this.Handles.Add(HandleOne);



        }
    }
}
