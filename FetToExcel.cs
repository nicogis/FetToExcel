//-----------------------------------------------------------------------
// <copyright file="FetToExcel.cs" company="Studio A&T s.r.l.">
//     Author: nicogis
//     Copyright (c) Studio A&T s.r.l. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
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
    using System.Web.Script.Serialization;

    public partial class FetToExcel : Form
    {       
        private readonly string pathFileTemplate = null;
        private readonly string pathFileOuputExcel = null;
        private readonly string cellStartTeachers = null;
        private readonly bool openExcel = false;

        public FetToExcel()
        {
            InitializeComponent();
            FileVersionInfo fileVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            this.Text = $"{fileVersion.ProductName} - Versione {fileVersion.ProductVersion}";

            this.pathFileTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TemplateOrario.xlsx");
            this.cellStartTeachers = "A5";

            // controllo se esiste il file di configurazione
            string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
            if (File.Exists(configFile))
            {
                try
                {
                    string jsonConfig = File.ReadAllText(configFile);
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    Config config = js.Deserialize<Config>(jsonConfig);

                    string pathTeachersXml = config.PathTeachersXml;
                    string pathTemplateExcel = config.PathTemplateExcel;
                    string pathOuputExcel = config.PathOuputExcel;
                    openExcel = config.OpenExcel;

                    if (!string.IsNullOrWhiteSpace(pathTeachersXml))
                    {
                        if (Directory.Exists(pathTeachersXml))
                        {
                            this.openFD.InitialDirectory = pathTeachersXml;
                        }
                        else
                        {
                            if (File.Exists(pathTeachersXml))
                            {
                                this.txtFileFet.Text = pathTeachersXml;
                                this.btnFileFet.Enabled = false;
                            }
                        }

                    }

                    if (!string.IsNullOrWhiteSpace(pathOuputExcel))
                    {
                        if (Directory.Exists(pathOuputExcel))
                        {
                            this.saveFD.InitialDirectory = pathOuputExcel;
                        }
                        else
                        {
                            this.pathFileOuputExcel = Path.Combine(Path.GetDirectoryName(pathOuputExcel), Path.GetFileNameWithoutExtension(pathOuputExcel) + "_{0}.xlsx");
                            this.btnFileExcel.Enabled = false;                           
                        }

                    }

                    if (File.Exists(pathTemplateExcel))
                    {
                        pathFileTemplate = pathTemplateExcel;
                    }

                    if (!string.IsNullOrWhiteSpace(config.CellStartTeachers))
                    {
                        cellStartTeachers = config.CellStartTeachers;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Errore nel file di configurazione '{configFile}' pertanto verrà stato ignorato.{Environment.NewLine}Errore: {ex.Message}");
                }
            }

            this.txtStartCell.Text = cellStartTeachers;
        }

        private void BtnFileFet_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openFD.ShowDialog();
            if (result == DialogResult.OK) 
            {
                this.txtFileFet.Text = this.openFD.FileName;
            }
        }

        private void BtnFileExcel_Click(object sender, EventArgs e)
        {
            this.saveFD.ShowDialog();
        }

        private void SaveFD_FileOk(object sender, System.ComponentModel.CancelEventArgs e) => this.txtFileExcel.Text = this.saveFD.FileName;

        private void BtnImporta_Click(object sender, EventArgs e)
        {
            try
            {                
                //check file
                if (string.IsNullOrWhiteSpace(this.txtFileFet.Text))
                {
                    MessageBox.Show("Indicare il file Fet ('*_teachers.xml') da importare in Excel!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.btnFileFet.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(pathFileOuputExcel))
                {
                    if (string.IsNullOrWhiteSpace(this.txtFileExcel.Text))
                    {
                        MessageBox.Show("Indicare il file Excel di output!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.btnFileExcel.Focus();
                        return;
                    }
                }
                else
                {
                    this.txtFileExcel.Text = string.Format(pathFileOuputExcel, DateTime.Now.ToString("yyyyMMddHHmmss"));
                }

                XmlDocument document = new XmlDocument();
                document.Load(this.txtFileFet.Text);
                XmlNodeList teachersTimeTable = document.GetElementsByTagName("Teachers_Timetable");
                if (teachersTimeTable.Count != 1)
                {
                    throw new Exception("Nodo Teachers_Timetable errato!");
                }

                

                if (!File.Exists(pathFileTemplate))
                {
                    MessageBox.Show($"Manca il file template Excel '{pathFileTemplate}'!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

                    string startCell = cellStartTeachers;
                    if (!string.IsNullOrWhiteSpace(txtStartCell.Text))
                    {
                        startCell = txtStartCell.Text;
                    }
                    
                    excelWorksheet.Cells[startCell.Trim()].LoadFromDataTable(dataSource, false);
                    excelPackage.Save();
                }

                MessageBox.Show("Esportazione effettuata correttamente!", "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (openExcel)
                {
                    Process.Start(fileInfo.FullName);
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Fet to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private int? IndexGiorno(string giorno, int firstDay = 0)
        {
            DayOfWeek? dayOfWeek;

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

        private void LlinkFetToExcel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ((LinkLabel)sender).LinkVisited = true;
            Process.Start("https://github.com/nicogis/FetToExcel/releases");
        }

        private void llinkFetToExcelHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ((LinkLabel)sender).LinkVisited = true;
            Process.Start("https://github.com/nicogis/FetToExcel/blob/master/README.md");
        }
    }
}
