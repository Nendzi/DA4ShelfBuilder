using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DA4ShelfBuilderPlugin.Models
{
    public class ConnectionDataModel
    {
        private double _ditance;
        private string _position;

        public double Distance
        {
            get { return _ditance; }
            set { _ditance = value; }
        }
        public string Position
        {
            get { return _position; }
            set { _position = value; }
        }
    }
}
