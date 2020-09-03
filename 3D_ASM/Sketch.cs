using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.swconst;

namespace ASM_3D
{
    class Sketch
    {
        public float x1, x2, y1, y2;
        public int arcType;


        static public int GetArcType(string angle)
        {
            switch (angle)
            {
                case "1":
                    return (int)swTangentArcTypes_e.swForward;
                case "2":
                    return (int)swTangentArcTypes_e.swLeft;
                case "3":
                    return (int)swTangentArcTypes_e.swBack;
                case "4":
                    return (int)swTangentArcTypes_e.swRight;
                default:
                    return 0;
                
            }
        }
    }


}

