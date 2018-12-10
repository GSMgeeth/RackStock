using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RackStock.Role
{
    class SubRack
    {
        private int subRackId;
        private string name;
        private LinkedList<Stock> stock;

        public SubRack(int subRackId)
        {
            SubRackId = subRackId;
            stock = new LinkedList<Stock>();
        }

        public SubRack(int subRackId, LinkedList<Stock> stock)
        {
            SubRackId = subRackId;
            Stock = stock;
        }

        public SubRack(string name)
        {
            Name = name;
        }

        public SubRack(int subRackId, string name)
        {
            SubRackId = subRackId;
            Name = name;
        }
        
        public int SubRackId { get => subRackId; set => subRackId = value; }
        public string Name { get => name; set => name = value; }
        internal LinkedList<Stock> Stock { get => stock; set => stock = value; }
    }
}
