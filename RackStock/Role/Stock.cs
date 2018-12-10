using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RackStock.Role
{
    class Stock
    {
        private int stockId;
        private string article;
        private string color;
        private string size;
        private DateTime date;
        private string desc;
        private int qty;

        public Stock(string article, string color, string size, DateTime date, string desc, int qty)
        {
            Article = article;
            Color = color;
            Size = size;
            Date = date;
            Desc = desc;
            Qty = qty;
        }

        public Stock(int stockId, string article, string color, string size, DateTime date, string desc, int qty)
        {
            StockId = stockId;
            Article = article;
            Color = color;
            Size = size;
            Date = date;
            Desc = desc;
            Qty = qty;
        }

        public int StockId { get => stockId; set => stockId = value; }
        public string Article { get => article; set => article = value; }
        public string Color { get => color; set => color = value; }
        public string Size { get => size; set => size = value; }
        public DateTime Date { get => date; set => date = value; }
        public string Desc { get => desc; set => desc = value; }
        public int Qty { get => qty; set => qty = value; }
    }
}
