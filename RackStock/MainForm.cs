using RackStock.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RackStock.Role;
using System.Runtime.InteropServices;
using MySql.Data.MySqlClient;

namespace RackStock
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        LinkedList<string> lables = new LinkedList<string>();

        public MainForm()
        {
            //change color to red when stock < 5000

            InitializeComponent();
        }
        
        private void MainForm_Load(object sender, EventArgs e)
        {

        }
        
        private void metroButton2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Workbook|*.xlsx|Excel Workbook 2003|*.xls";
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                string name = openFileDialog1.SafeFileName;

                if (name.Contains(".xlsx") || name.Contains(".xls"))
                {
                    _Application excel = new _Excel.Application();
                    Workbook wb;
                    Worksheet ws;

                    string path = "D:/LabelStock/" + name;

                    wb = excel.Workbooks.Open(path);
                    ws = wb.Worksheets[1];
                    
                    for (int row = 5; row <= 26; row++)
                    {
                        string article;

                        if (ws.Cells[row, 1].Value2 is double)
                        {
                            article = "" + (int)ws.Cells[row, 1].Value2;
                        }
                        else
                        {
                            article = ws.Cells[row, 1].Value2;
                        }

                        string desc;

                        if (ws.Cells[row, 2].Value2 is double)
                        {
                            desc = "" + (int)ws.Cells[row, 2].Value2;
                        }
                        else
                        {
                            desc = ws.Cells[row, 2].Value2;
                        }

                        string color;

                        if (ws.Cells[row, 3].Value2 is double)
                        {
                            color = "" + (int)ws.Cells[row, 3].Value2;
                        }
                        else
                        {
                            color = ws.Cells[row, 3].Value2;
                        }

                        string size;

                        if (ws.Cells[row, 4].Value2 is double)
                        {
                            size = "" + (int)ws.Cells[row, 4].Value2;
                        }
                        else
                        {
                            size = ws.Cells[row, 4].Value2;
                        }

                        int rackId;
                        string rName;

                        if (ws.Cells[row, 5].Value2 is double)
                        {
                            rName = "" + (int)ws.Cells[row, 5].Value2;
                        }
                        else
                        {
                            rName = ws.Cells[row, 5].Value2;
                        }

                        MySqlDataReader reader = DBConnection.getData("select rack_id from rack where name='" + rName + "'");

                        if (reader.Read())
                        {
                            rackId = reader.GetInt32(0);

                            reader.Close();

                            int subRackId;
                            string srName;

                            if (ws.Cells[row, 6].Value2 is double)
                            {
                                srName = "" + (int)ws.Cells[row, 6].Value2;
                            }
                            else
                            {
                                srName = ws.Cells[row, 6].Value2;
                            }

                            reader = DBConnection.getData("select sub_rack_id from sub_rack where name='" + srName + "'");

                            if (reader.Read())
                            {
                                subRackId = reader.GetInt32(0);

                                reader.Close();

                                int qty = 0;

                                if (ws.Cells[row, 7].Value2 is double)
                                {
                                    qty = (int)ws.Cells[row, 7].Value2;
                                    qty *= (int)ws.Cells[row, 8].Value2;
                                }
                                else
                                {
                                    qty = ws.Cells[row, 7].Value2;
                                    qty *= ws.Cells[row, 8].Value2;
                                }

                                int tmpQty = 0;
                                row++;

                                while ((ws.Cells[row, 1].Value2 == null) && (ws.Cells[row, 7].Value2 != null))
                                {
                                    tmpQty = 0;

                                    if (ws.Cells[row, 7].Value2 is double)
                                    {
                                        tmpQty += (int)ws.Cells[row, 7].Value2;
                                        tmpQty *= (int)ws.Cells[row, 8].Value2;

                                        qty += tmpQty;
                                    }
                                    else
                                    {
                                        tmpQty += ws.Cells[row, 7].Value2;
                                        tmpQty *= ws.Cells[row, 8].Value2;

                                        qty += tmpQty;
                                    }

                                    row++;
                                }

                                row--;

                                MainRack mr = new MainRack(rackId);
                                SubRack sr = new SubRack(subRackId);
                                Stock s = new Stock(article, color, size, DateTime.Today, desc, qty);

                                LinkedList<Stock> stocks = new LinkedList<Stock>();
                                LinkedList<SubRack> sracks = new LinkedList<SubRack>();

                                stocks.AddFirst(s);

                                sr.Stock = stocks;

                                sracks.AddFirst(sr);

                                mr.Subracks = sracks;

                                try
                                {
                                    Database.addStock(mr);
                                }
                                catch (Exception exc)
                                {
                                    MessageBox.Show("Something wrong with the qty cell in excel file!\n" + exc, "File reader", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                if (!reader.IsClosed)
                                    reader.Close();

                                MessageBox.Show("No Main Sub Rack - " + srName, "File reader", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            if (!reader.IsClosed)
                                reader.Close();

                            MessageBox.Show("No Main Rack - " + rName, "File reader", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    wb.Close();
                    excel.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(excel);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Something wrong with the excel file!\n" + exception, "File reader", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void searchBtn_Click(object sender, EventArgs e)
        {
            String art = artTxtBox.Text;
            String color = colorTxtBox.Text;
            String size = sizeTxtBox.Text;

            foreach (string s in lables)
            {
                this.Controls.Find(s, true)[0].ForeColor = Color.White;
            }

            lables = new LinkedList<string>();

            String qry = "";

            if (!art.Equals("") && !color.Equals("") && !size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.article='" + art + "' and s.color='" + color + "' and s.size='" + size + "'";
            }
            else if (!art.Equals("") && !color.Equals("") && size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.article='" + art + "' and s.color='" + color + "'";
            }
            else if (!art.Equals("") && color.Equals("") && !size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.article='" + art + "' and s.size='" + size + "'";
            }
            else if (art.Equals("") && !color.Equals("") && !size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.color='" + color + "' and s.size='" + size + "'";
            }
            else if (!art.Equals("") && color.Equals("") && size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.article='" + art + "'";
            }
            else if (art.Equals("") && !color.Equals("") && size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.color='" + color + "'";
            }
            else if (art.Equals("") && color.Equals("") && !size.Equals(""))
            {
                qry = "select sr.name from stock s inner join sub_rack sr on s.rack_id=sr.rack_id and s.sub_rack_id=sr.sub_rack_id " +
                        "where s.size='" + size + "'";
            }
            else if (art.Equals("") && color.Equals("") && size.Equals(""))
            {
                MessageBox.Show("Must add at least one factor to filter!", "Search Rack", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            if (!qry.Equals(""))
            {
                MySqlDataReader reader = DBConnection.getData(qry);

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string subRackName = reader.GetString(0);

                        string letter = subRackName.Substring(subRackName.Length - 1).ToLower();
                        string no = subRackName.Substring(0, subRackName.Length - 1);

                        string lblName = letter + no + "Lbl";

                        this.Controls.Find(lblName, true)[0].ForeColor = Color.Red;

                        lables.AddLast(lblName);
                    }
                }

                reader.Close();
            }
        }
    }
}
