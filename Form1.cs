// Form1.cs
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using ClosedXML.Excel;
using Newtonsoft.Json;
using AddressFilter.Models;

namespace AddressFilter
{
    public partial class Form1 : Form
    {
        private DataTable _table;
        private Dictionary<string, string> _postcodeMap;
        public Form1()
        {
            InitializeComponent();

            // Load and flatten Malaysia postcode→city map
            try
            {
                var rawJson = File.ReadAllText("postcode_city.json");
                var states = JsonConvert.DeserializeObject<List<State>>(rawJson)
                             ?? new List<State>();

                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var state in states)
                    foreach (var city in state.city)
                        foreach (var pc in city.postcode)
                        {
                            var key = pc.Trim();
                            if (!map.ContainsKey(key))
                                map[key] = city.name;
                        }
                _postcodeMap = map;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading or parsing JSON map: {ex.Message}");
                _postcodeMap = new Dictionary<string, string>();
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel Workbook|*.xlsx;*.xls",
                Title = "Select Excel File"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook(dlg.FileName);
                var ws = wb.Worksheets.Worksheet(1);
                var range = ws.RangeUsed();
                // Create a new DataTable and preset headers
                var dt = new DataTable();
                dt.Columns.Add("custID");   
                dt.Columns.Add("CustName"); 
                dt.Columns.Add("UnitNumber");
                dt.Columns.Add("Address1");
                dt.Columns.Add("Address2");
                dt.Columns.Add("City");     
                dt.Columns.Add("Postcode"); 
                dt.Columns.Add("State");    
                dt.Columns.Add("Col9");
                dt.Columns.Add("Contact");  
                dt.Columns.Add("Col11");
                dt.Columns.Add("Col12");
                dt.Columns.Add("Col13");

                // Load every row (including header row) as data
                foreach (var excelRow in range.Rows())
                {
                    var values = new object[13];
                    values[0] = excelRow.Cell(1).GetString();  // custID
                    values[1] = excelRow.Cell(2).GetString();  // CustName
                    values[2] = excelRow.Cell(3).GetString();  // UnitNumber
                    values[3] = excelRow.Cell(4).GetString();  // Address1
                    values[4] = excelRow.Cell(5).GetString();  // Address2
                    values[5] = string.Empty;                  // City (to fill)
                    values[6] = excelRow.Cell(7).GetString();  // Postcode
                    values[7] = string.Empty;                  // State (to fill)
                    values[8] = string.Empty;                  // Col9
                    values[9] = excelRow.Cell(10).GetString(); // Contact
                    values[10] = string.Empty;
                    values[11] = string.Empty;
                    values[12] = string.Empty;
                    dt.Rows.Add(values);
                }

                _table = dt;
                dataGridView1.DataSource = _table;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing Excel: {ex.Message}");
            }
        }

        private void btnMap_Click(object sender, EventArgs e)
        {
            if (_table == null)
            {
                MessageBox.Show("Import first!");
                return;
            }

            try
            {
                foreach (DataRow row in _table.Rows)
                {
                   var pc = row["Postcode"]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(pc))
                {
                    if (_postcodeMap.TryGetValue(pc, out var city))
                        row["City"] = city;
                    //if (_postcodeMap.TryGetValue(pc, out var st))
                    //    row["State"] = st;
                }
                }

                dataGridView1.Refresh();
                MessageBox.Show("Mapping complete!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error mapping cities: {ex.Message}");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (_table == null)
            {
                MessageBox.Show("Nothing to export.");
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "Save Filtered Excel"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();
                wb.Worksheets.Add(_table, "Filtered");
                wb.SaveAs(dlg.FileName);
                MessageBox.Show("Exported to " + dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting Excel: {ex.Message}");
            }
        }
    }
}
