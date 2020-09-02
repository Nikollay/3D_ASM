using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3D_ASM
{
    class Board
    {
        public List<Line> lines=new List<Line>();
        public List<Arc> arcs = new List<Arc>();
        public List<Circle> circles = new List<Circle>();
        public double thickness;
    }
}
