using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DA4ShelfBuilderPlugin.Models
{
    public class MidPoint2DModel
    {
        private double _x;
        private double _y;
        public double X
        {
            get { return _x; }
            set { _x = value; }
        }        
        public double Y
        {
            get { return _y; }
            set { _y = value; }
        }
    }
}
