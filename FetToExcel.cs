namespace FetToExcel
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Reflection;
    using System.Windows.Forms;
    using System.Xml;
    using System.Linq;
    using System.Data;
    using System.Diagnostics;

    public partial class FetToExcel : Form
    {
        public FetToExcel()
        {
            InitializeComponent();
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            this.Text = $"{fileVersion.ProductName} - Versione {fileVersion.ProductVersion}";
        }

        private void btnFileFet_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openFD.ShowDialog();
            if (result == DialogResult.OK) 
            {
                this.txtFileFet.Text = this.openFD.FileName;
            }
        }

        private void btnFileExcel_Click(object sender, EventArgs e)
        {
            this.saveFD.ShowDialog();
        }

        private void saveFD_FileOk(object sender, System.ComponentModel.CancelEventArgs e) => this.txtFileExcel.Text = this.saveFD.FileName;

        private void btnImporta_Click(object sender, EventArgs e)
        {
            try
            {
                //check file
                if (string.IsNullOrWhiteSpace(this.txtFileFet.Text))
                {
                    MessageBox.Show("Indicare il file Fet da importare in Excel!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.btnFileFet.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(this.txtFileExcel.Text))
                {
                    MessageBox.Show("Indicare il file Excel!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.btnFileExcel.Focus();
                    return;
                }

                XmlDocument document = new XmlDocument();
                document.Load(this.txtFileFet.Text);
                XmlNodeList teachersTimeTable = document.GetElementsByTagName("Teachers_Timetable");
                if (teachersTimeTable.Count != 1)
                {
                    throw new Exception("Nodo Teachers_Timetable errato!");
                }

                string pathFileTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TemplateOrario.xlsx");

                if (!File.Exists(pathFileTemplate))
                {
                    MessageBox.Show($"Manca il file {pathFileTemplate}!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }


                XmlNodeList teachers = teachersTimeTable[0].ChildNodes;

                
                
                List<int> giorni = new List<int>();
                List<int> ore = new List<int>();

                //check e controllo limiti
                foreach (XmlNode teacher in teachers)
                {
                    string teacherName = teacher.Attributes["name"].Value;
                    if (string.IsNullOrWhiteSpace(teacherName))
                    {
                        throw new Exception("Controllare xml perchè non è stata trovato il nome di una teacher!");
                    }

                    foreach (XmlNode day in teacher.ChildNodes)
                    {
                        string giorno = day.Attributes["name"].Value;

                        int? k = IndexGiorno(giorno);
                        if (k.HasValue)
                        { 
                            giorni.Add(k.Value);
                        }
                        else
                        {
                            throw new Exception($"Controllare il giorno nell'xml perchè '{teacherName}-{giorno}' non è gestito!");
                        }

                        foreach (XmlNode hour in day.ChildNodes)
                        {
                            if (!hour.HasChildNodes)
                            {
                                continue;
                            }

                            string ora = hour.Attributes["name"].Value;

                            ore.Add(this.OffsetOra(ora));

                            XmlNodeList students = hour.SelectNodes("Students");

                            // teacher ha solo un students (classe) nella stessa ora
                            if (students.Count != 1)
                            {
                                throw new Exception($"Controllare la classe nell'xml perchè '{teacherName}-{giorno}-{ora}' ha più classi e può essercene una sola!");
                            }
                        }
                    }
                }

                //indice del primo giorno
                int idxFirstDay = giorni.Min();
                //indice dell'ultimo giorno
                int idxMaxGiorno = giorni.Max();
                //numero di giorni
                int numeroGiorni = idxMaxGiorno - idxFirstDay + 1;

                int idxMaxOre = ore.Max();
                int idxMinOre = ore.Min();
                int numeroOre = idxMaxOre - idxMinOre + 1;

                int numeroColonne = 1 + numeroGiorni * numeroOre;
                DataTable dataSource = new DataTable();
                foreach (int i in Enumerable.Range(1, numeroColonne))
                {
                    dataSource.Columns.Add($"Col{i}");
                }

                int? idxOra = null;
                int? idxGiorno = null;
                int? numeroColonna = null;

                foreach (XmlNode teacher in teachers)
                {
                    string teacherName = teacher.Attributes["name"].Value;

                    DataRow row = dataSource.NewRow();

                    row["Col1"] = teacherName;

                    foreach (XmlNode day in teacher.ChildNodes)
                    {
                        string giorno = day.Attributes["name"].Value;

                        idxGiorno = this.IndexGiorno(giorno, idxFirstDay);
                        

                        foreach (XmlNode hour in day.ChildNodes)
                        {
                            if (!hour.HasChildNodes)
                            {
                                continue;
                            }

                            string ora = hour.Attributes["name"].Value;
                            idxOra = this.OffsetOra(ora, idxMinOre);
                            
                            XmlNodeList students = hour.SelectNodes("Students");
                           
                            string studenti = students[0].Attributes["name"].Value;
                            
                            //numeroColonna = posizione nome insegnante + posizione giorno + posizione ora
                            numeroColonna = 1 + idxGiorno.Value * numeroOre + idxOra + 1;
                            
                            row[$"Col{numeroColonna.Value}"] = studenti;
                        }
                    }

                    dataSource.Rows.Add(row);
                }

                var fileInfo = new FileInfo(this.txtFileExcel.Text);
                var fileInfoTemplate = new FileInfo(pathFileTemplate);
                using (var excelPackage = new ExcelPackage(fileInfo, fileInfoTemplate))
                {

                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[1]; // 1° foglio

                    string startCell = "A5";
                    if (!string.IsNullOrWhiteSpace(txtStartCell.Text))
                    {
                        startCell = txtStartCell.Text;
                    }
                    
                    excelWorksheet.Cells[startCell.Trim()].LoadFromDataTable(dataSource, false);
                    excelPackage.Save();
                }

                MessageBox.Show("Esportazione effettuata correttamente!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }


        private int? IndexGiorno(string giorno, int firstDay = 0)
        {
            DayOfWeek? dayOfWeek = null;

            if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Monday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Monday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Tuesday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Tuesday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Wednesday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Wednesday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Thursday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Thursday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Friday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Friday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Saturday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Saturday;
            }
            else if (string.Compare(DateTimeFormatInfo.CurrentInfo.GetDayName(DayOfWeek.Sunday), giorno, true) == 0)
            {
                dayOfWeek = DayOfWeek.Sunday;
            }
            else
            {
                return null;
            }

            return (int)dayOfWeek - firstDay;
        }

        private int OffsetOra(string ora, int inizioOra = 0)
        {
            TimeSpan ts = new TimeSpan(Convert.ToInt32(ora.Split(':')[0]), 0, 0);
            int idx = ts.Hours - inizioOra;
            return idx;
        }

        private void llinkFetToExcel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.llinkFetToExcel.LinkVisited = true;
            System.Diagnostics.Process.Start("https://github.com/nicogis/FetToExcel/releases");
        }
    }
}
