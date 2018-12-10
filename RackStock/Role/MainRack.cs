using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RackStock.Role
{
    class MainRack
    {
        private int rackId;
        private string name;
        private LinkedList<SubRack> subracks;

        public MainRack()
        {
            Subracks = new LinkedList<SubRack>();
        }

        public MainRack(int rackId)
        {
            RackId = rackId;
            Subracks = new LinkedList<SubRack>();
        }

        public MainRack(int rackId, LinkedList<SubRack> subracks)
        {
            RackId = rackId;
            Subracks = subracks;
        }

        public MainRack(int rackId, string name)
        {
            RackId = rackId;
            Name = name;
        }

        public MainRack(string name)
        {
            this.Name = name;
        }

        public int RackId { get => rackId; set => rackId = value; }
        public string Name { get => name; set => name = value; }
        internal LinkedList<SubRack> Subracks { get => subracks; set => subracks = value; }
    }
}
