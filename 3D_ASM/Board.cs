using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASM_3D
{
    class Board
    {
        public List<Component> components;
        public List<Circle> circles;
        public List<Point> point;
        public List<Object> sketh, cutout;

        public double thickness;
    }
}
