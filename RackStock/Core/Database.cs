using MySql.Data.MySqlClient;
using RackStock.Role;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RackStock.Core
{
    class Database
    {
        public static void addRack(MainRack rack)
        {
            try
            {
                string qry = "insert into rack (name) values ('" + rack.Name + "')";

                DBConnection.updateDB(qry);
            }
            catch (Exception exc)
            {
                MessageBox.Show("Something went wrong!\n" + exc, "Add Main Rack", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void addSubRack(MainRack rack)
        {
            try
            {
                int rackId = rack.RackId;
                LinkedList<SubRack> subracks = rack.Subracks;

                foreach (SubRack sr in subracks)
                {
                    string qry = "insert into sub_rack (rack_id, name) values (" + rackId + ", '" + sr.Name + "')";

                    DBConnection.updateDB(qry);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show("Something went wrong!\n" + exc, "Add Sub Rack", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void addStock(MainRack rack)
        {
            try
            {
                int rackId = rack.RackId;
                int subRackId;
                
                foreach (SubRack sr in rack.Subracks)
                {
                    subRackId = sr.SubRackId;

                    foreach (Stock stock in sr.Stock)
                    {
                        string qry = "insert into stock (article, color, size, description, rack_id, sub_rack_id, date, qty) " +
                            "values ('" + stock.Article + "', '" + stock.Color + "', '" + stock.Size + "', '" + stock.Desc + "', " + rackId + ", " + subRackId + ", " +
                            "'" + stock.Date.ToString("yyyy/MM/d") + "', " + stock.Qty + ")";

                        DBConnection.updateDB(qry);
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show("Something went wrong!\n" + exc, "Add Stock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
