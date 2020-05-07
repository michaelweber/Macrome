using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using System.Reflection;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    public class ShapeType : IVisitable
    {
        public enum JoinStyle
        {
            miter,
            round,
            bevel,
            none
        }
        //inner Class
        public class Handle
        {

            public Handle()
            { }

            
            [Obsolete("Use default constuctor")]
            public Handle(string pos, string xRange) 
            {
                this.position = pos;
                this.xrange = xRange;
            }
  
            
            public string position = null;
            public string xrange = null;
            public string switchHandle = null;
            public string yrange = null;
            public string polar = null;
            public string radiusrange = null; 

        }

        /// <summary>
        /// This string describes a sequence of commands that define the shape’s path.<br/>
        /// This string describes both the pSegmentInfo array and pVertices array in the shape’s geometry properties.
        /// </summary>
        public string Path;


        /// <summary>
        /// This specifies a list of formulas whose calculated values are referenced by other properties. <br/>
        /// Each formula is listed on a separate line. Formulas are ordered, with the first formula having index 0. <br/>
        /// This section can be omitted if the shape doesn’t need any guides.
        /// </summary>
        public List<string> Formulas;


        /// <summary>
        /// Specifies a comma-delimited list of parameters, or adjustment values, 
        /// used to define values for a parameterized formula. <br/>
        /// These values represent the location of an adjust handle and may be 
        /// referenced by the geometry of an adjust handle or as a parameter guide function.
        /// </summary>
        public string AdjustmentValues;


        /// <summary>
        /// These values specify the location of connection points on the shape’s path. <br/>
        /// The connection points are defined by a string consisting of pairs of x and y values, delimited by commas.
        /// </summary>
        public string ConnectorLocations;


        public string ConnectorType;

        public bool PreferRelative;

        public bool ExtrusionOk = false;

        /// <summary>
        /// This section specifies the properties of each adjust handle on the shape. <br/>
        /// One adjust handle is specified per line. <br/>
        /// The properties for each handle correspond to values of the ADJH structure 
        /// contained in the pAdjustHandles array in the shape’s geometry properties.
        /// </summary>
        public List<Handle> Handles;


        /// <summary>
        /// Specifies one or more text boxes inscribed inside the shape. <br/>
        /// A textbox is defined by one or more sets of numbers specifying (in order) the left, top, right, and bottom points of the rectangle. <br/>
        /// Multiple sets are delimited by a semicolon. <br/>
        /// If omitted, the text box is the same as the geometry’s bounding box.
        /// </summary>
        public string TextboxRectangle;


        /// <summary>
        /// 
        /// </summary>
        public bool ShapeConcentricFill;

        /// <summary>
        /// Specifies what join style the shape has. <br/>
        /// Since there is no UI for changing the join style, 
        /// all shapes of this type will always have the specified join style.
        /// </summary>
        public JoinStyle Joins;

        /// <summary>
        /// Specifies the (x,y) coordinates of the limo stretch point.<br/>
        /// Some shapes that have portions that should be constrained to a fixed aspect ratio, are designed with limo-stretch to keep those portions at the fixed aspect ratio.<br/>
        /// </summary>
        public string Limo;
     
        /// <summary>
        /// Associated with each connection site, there is a direction which specifies at what angle elbow and curved connectors should attach to it<br/>
        /// </summary>
        public string ConnectorAngles;

        /// <summary>
        /// Specifies if a shape of this type is filled by default
        /// </summary>
        public bool Filled = true;

        /// <summary>
        /// Specifies if a shape of this type is stroked by default
        /// </summary>
        public bool Stroked = true;

        /// <summary>
        /// Speicfies the locked properties of teh shape.
        /// By default nothing is locked.
        /// </summary>
        public ProtectionBooleans Lock = new ProtectionBooleans(0);

        public bool LockShapeType;

        public bool TextPath;

        public bool TextKerning;

        public uint TypeCode
        {
            get 
            {
                uint ret = 0;

                var attrs = this.GetType().GetCustomAttributes(typeof(OfficeShapeTypeAttribute), false);
                OfficeShapeTypeAttribute attr = null;

                if (attrs.Length > 0)
                {
                    attr = attrs[0] as OfficeShapeTypeAttribute;
                }

                if (attr != null)
                {
                    ret = attr.TypeCode;
                }

                return ret;
            }
        }
	


        private static Dictionary<uint, Type> TypeToShapeClassMapping = new Dictionary<uint, Type>();


        static ShapeType()
        {
            UpdateTypeToShapeClassMapping(Assembly.GetExecutingAssembly(), typeof(ShapeType).Namespace);
        }


        public static ShapeType GetShapeType(uint typeCode)
        {
            ShapeType result;
            Type cls;

            if (TypeToShapeClassMapping.TryGetValue(typeCode, out cls))
            {
                var constructor = cls.GetConstructor(new Type[] {});

                if (constructor == null)
                {
                    throw new Exception(string.Format(
                        "Internal error: Could not find a matching constructor for class {0}",
                        cls));
                }

                try
                {
                    result = (ShapeType)constructor.Invoke(new object[] {});
                }
                catch (TargetInvocationException e)
                {
                    TraceLogger.DebugInternal(e.InnerException.ToString());
                    throw e.InnerException;
                }
            }
            else
            {
                result = null;
            }

            return result;
        }


        /// <summary>
        /// Updates the Dictionary used for mapping Office shape type codes to Office ShapeType classes.
        /// This is done by querying all classes in the specified assembly filtered by the specified
        /// namespace and looking for attributes of type OfficeShapeTypeAttribute.
        /// </summary>
        /// 
        /// <param name="assembly">Assembly to scan</param>
        /// <param name="ns">Namespace to scan or null for all namespaces</param>
        public static void UpdateTypeToShapeClassMapping(Assembly assembly, string ns)
        {
            foreach (var t in assembly.GetTypes())
            {
                if (ns == null || t.Namespace == ns)
                {
                    var attrs = t.GetCustomAttributes(typeof(OfficeShapeTypeAttribute), false);
                    OfficeShapeTypeAttribute attr = null;

                    if (attrs.Length > 0)
                    {
                        attr = attrs[0] as OfficeShapeTypeAttribute;
                    }

                    if (attr != null)
                    {
                        TypeToShapeClassMapping.Add(attr.TypeCode, t);
                    }
                }
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ShapeType>)mapping).Apply(this);
        }

        #endregion
    }
}
