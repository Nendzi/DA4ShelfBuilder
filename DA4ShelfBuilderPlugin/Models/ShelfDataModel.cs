using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DA4ShelfBuilderPlugin.Models
{
    public class ShelfDataModel
    {
        private double _length;
        private string _orientation;
        private MidPoint2DModel _midPoint;
        private double _depth;
        private double _thickness;
        private int _material;
        private List<ConnectionDataModel> _connectionList = new List<ConnectionDataModel>();
        private bool _connectionOnBegin;
        private bool _connectionOnEnd;

        public double Length
        {
            get { return _length; }
            set { _length = value; }
        }
        public string Orientation
        {
            get { return _orientation; }
            set { _orientation = value; }
        }
        public MidPoint2DModel MidPoint
        {
            get { return _midPoint; }
            set { _midPoint = value; }
        }
        public double Depth
        {
            get { return _depth; }
            set { _depth = value; }
        }
        public double Thickness
        {
            get { return _thickness; }
            set { _thickness = value; }
        }
        public int Material
        {
            get { return _material; }
            set { _material = value; }
        }
        public List<ConnectionDataModel> ConnectionList
        {
            get { return _connectionList; }
            set { _connectionList = value; }
        }
        public bool ConnectionOnBegin
        {
            get { return _connectionOnBegin; }
            set { _connectionOnBegin = value; }
        }
        public bool ConnectionOnEnd
        {
            get { return _connectionOnEnd; }
            set { _connectionOnEnd = value; }
        }
        public void SetConnectionData(ConnectionDataModel data)
        {
            ConnectionList.Add(data);
        }
    }
}
