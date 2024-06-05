using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using E3Lib;
using EDSommaireByINDEX;

namespace EDSommaireByINDEX
{
    public partial class Form1 : Form
    {

        bool revisionFound = false;
        public List<SummaryRow> rows = new List<SummaryRow>();
        public Dictionary<string, List<string>> Revisions = new Dictionary<string, List<string>>();

        public CancellationTokenSource cancellationTokenSource;

        public Form1()
        {

            // checkbox1 = mode auto
            //checkbox 2 = groupé par projet


            InitializeComponent();
            if (!E3Application.Instances.SelectMany(i => i.JobIds).Any())
            {
                MessageBox.Show("Please Open an E3 Project.");
                this.Close();
                return;
            }
            Load();
            using (E3Job job = E3Job.Instances.First())
            {
                var distinctPrjNumeros = job.Sheets.Select(s => s.GetAttributeValue(textBox18.Text)).Distinct().ToList();




                if (checkBox1.Checked)
                {
                    if (checkBox2.Checked == false)
                    {
                        OnePrjNumCreateSheet();
                        this.Close();
                        return;
                    }
                    else
                    {
                        CreateMultipleProjectSheet();
                        this.Close();
                        return;
                    }



                }


            }

            //A_RevAct
            // StartCheckingApplicationRunning();
        }


        private void StartCheckingApplicationRunning()
        {
            // Initialize cancellation token source
            cancellationTokenSource = new CancellationTokenSource();

            // Start a background thread to check the application state
            var thread = new Thread(() => CheckApplicationRunning(cancellationTokenSource.Token));
            thread.Start();
        }

        private void StopCheckingApplicationRunning()
        {
            // Request cancellation of the background thread
            cancellationTokenSource?.Cancel();
        }

        private void CheckApplicationRunning(CancellationToken cancellationToken)
        {
            try
            {
                // Run until cancellation is requested
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        // Attempt to get an instance of E3Application
                        var e3Application = E3Application.Instances.First();

                        // Application is running, you can perform additional actions here if needed
                    }
                    catch (InvalidOperationException)
                    {
                        // No instance of E3Application is found, handle accordingly
                        // For example, close the form or perform any cleanup
                        this.Invoke((Action)delegate { this.Close(); });
                    }
                    catch (Exception ex)
                    {
                        // Handle other exceptions
                        // Log or perform any necessary actions based on your requirements
                        Console.WriteLine($"Exception: {ex.Message}");
                    }

                    // Sleep for 1 second before the next check
                    Thread.Sleep(1000);
                }
            }
            catch (OperationCanceledException)
            {
                // Cancellation is requested, exit the thread
            }
        }
        public void Save()
        {// PROBLEME PATH
            string path = "./save.txt";
            File.WriteAllText(path, "");
            using (StreamWriter w = new StreamWriter(path))
            {
                w.WriteLine("YMAX :");
                w.WriteLine(numericUpDown1.Value);
                w.WriteLine("YMIN :");
                w.WriteLine(numericUpDown2.Value);
                w.WriteLine("X_SHEETNAME :");
                w.WriteLine(numericUpDown3.Value);
                w.WriteLine("X_TITLE1 :");
                w.WriteLine(numericUpDown4.Value);
                w.WriteLine("X_TITLE2 :");
                w.WriteLine(numericUpDown5.Value);
                w.WriteLine("X_FUNCTION :");
                w.WriteLine(numericUpDown6.Value);
                w.WriteLine("X_REVISION :");
                w.WriteLine(numericUpDown7.Value);
                w.WriteLine("X_REVISION_DELTA :");
                w.WriteLine(numericUpDown8.Value);
                w.WriteLine("Y_REVISION_HEADER :");
                w.WriteLine(numericUpDown9.Value);
                w.WriteLine("Y_SPACING :");
                w.WriteLine(numericUpDown10.Value);
                w.WriteLine("ROWS_PER_PAGE :");
                w.WriteLine(numericUpDown11.Value);
                w.WriteLine("TEXT_HEIGHT :");
                w.WriteLine(numericUpDown12.Value);
                w.WriteLine("SHEET_FORMAT :");
                w.WriteLine(textBox13.Text);
                w.WriteLine("SHEET_FORMAT_FIND :");
                w.WriteLine(textBox14.Text);
                w.WriteLine("SHEET_PREFIX :");
                w.WriteLine(textBox15.Text);
                w.WriteLine("FIRSTSHEET_REVISION_INDEX :");
                w.WriteLine(numericUpDown14.Value);
                w.WriteLine("SHEET_TITLE :");
                w.WriteLine(textBox17.Text);
                w.WriteLine("AttributBySheet :");
                w.WriteLine(textBox18.Text);
                w.WriteLine("SheetInfos1 :");
                w.WriteLine(textBox19.Text);
                w.WriteLine("SheetInfos2 :");
                w.WriteLine(textBox20.Text);
                w.WriteLine("SheetInfosRev :");
                w.WriteLine(textBox21.Text);
                w.WriteLine("SheetIndiceRev :");
                w.WriteLine(textBox22.Text);
                w.WriteLine("Mode Auto :");
                if (checkBox1.Checked)
                    w.WriteLine("true");
                else
                    w.WriteLine("false");
                w.WriteLine("MultiProjet :");
                if (checkBox2.Checked)
                    w.WriteLine("true");
                else
                    w.WriteLine("false");

            }
        }

        public void Load()
        {
            try
            {
                using (StreamReader r = new StreamReader("./save.txt"))
                {
                    if (!File.Exists("./save.txt"))
                    {
                        return;
                    }
                    try
                    {
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown1);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown2);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown3);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown4);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown5);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown6);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown7);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown8);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown9);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown10);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown11);
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown12);
                        r.ReadLine();
                        textBox13.Text = r.ReadLine();
                        r.ReadLine();
                        textBox14.Text = r.ReadLine();
                        r.ReadLine();
                        textBox15.Text = r.ReadLine();
                        r.ReadLine();
                        AssignValue(r.ReadLine(), numericUpDown14);
                        r.ReadLine();
                        textBox17.Text = r.ReadLine();
                        r.ReadLine();
                        textBox18.Text = r.ReadLine();
                        r.ReadLine();
                        textBox19.Text = r.ReadLine();
                        r.ReadLine();
                        textBox20.Text = r.ReadLine();
                        r.ReadLine();
                        textBox21.Text = r.ReadLine();
                        r.ReadLine();
                        textBox22.Text = r.ReadLine();
                        r.ReadLine();
                        if (r.ReadLine().Equals("true"))
                            checkBox1.Checked = true;
                        r.ReadLine();
                        if (r.ReadLine().Equals("true"))
                            checkBox2.Checked = true;




                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (FileNotFoundException ex)
            {
            }

        }
        public void AssignValue(string value, Control control)
        {
            // Vérifier si la valeur peut être convertie en décimal
            if (decimal.TryParse(value, out decimal numericValue))
            {
                // Si oui, assigner à NumericUpDown
                ((NumericUpDown)control).Value = numericValue;
            }
            else
            {
                // Sinon, assigner à TextBox
                control.Text = value;
            }
        }
        public void LoadData(string prjNum)
        {


            Revisions.Clear();
            rows.Clear();

            using (E3Job job = E3Job.Instances.First())
            {

                revisionFound = false;

                foreach (var sheet in job.Sheets)
                {
                    if (sheet.Format == textBox14.Text)
                    {
                        revisionFound = true;
                        break;
                    }
                }
                foreach (var sheet in job.Sheets)
                {
                    List<string> rowRevisions = new List<string>();

                    if (prjNum != null && sheet.GetAttributeValue(textBox18.Text) != prjNum)
                    {
                        continue;
                    }

                    if (sheet.Format == textBox14.Text)
                    {
                        foreach (var text in sheet.Texts)
                        {
                            if (text.Type == 600)
                            {

                                string key = sheet.GetAttributeValue(textBox18.Text);
                                string value = text.InternalText;

                                if (Revisions.TryGetValue(key, out List<string> list))
                                {
                                    if (!list.Contains(value))
                                    {
                                        if (value != "")
                                            list.Add(value);
                                    }
                                }
                                else
                                {
                                    Revisions.Add(key, new List<string> { value });
                                }
                            }
                        }
                    }
                    
                    SummaryRow row = new SummaryRow
                    {
                        sheet = sheet.Name,
                        prjNum = sheet.GetAttributeValue(textBox18.Text),
                        Description1 = sheet.GetAttributeValue(textBox19.Text),
                        Description2 = sheet.GetAttributeValue(textBox20.Text),
                        function = sheet.Assignment + " " + sheet.Location
                    };

                    foreach (var attribute in sheet.Attributes)
                    {
                        
                        if (attribute.Value != "" && attribute.InternalName.Contains("A_Revision") && !rowRevisions.Contains(attribute.Value))
                        {
                            rowRevisions.Add(attribute.Value);
                        }
                            
                    }
                    row.Revisions = rowRevisions;

                    rows.Add(row);

                    
                }
            }
        }
        public void createOneSheet(string prjNum, E3Job job, string sheetName)
        {

            E3Sheet summarySheet;

            double Y_Spacing = (double)numericUpDown10.Value;
            double X_SpacingTitle = 8;

            //create the sheets and check what E3 returned
            summarySheet = job.CreateSheet();
            summarySheet.Create(sheetName, textBox13.Text);
            summarySheet.SetAttributeValue(textBox18.Text, prjNum);
            summarySheet.SetAttributeValue(textBox22.Text, Revisions.Values.Last().Last());
            //add them to rows
            if (summarySheet.Id == 0)
            {
                MessageBox.Show("Sheet could not be created.");
            }
            job.Application.PutInfo(0, sheetName + " created.");
            int maxRevisions = rows.Max(r => r.Revisions.Count);
            SummaryRow row = new SummaryRow();

            row.sheet = summarySheet.Name;
            row.prjNum = prjNum;
            row.Description1 = summarySheet.GetAttributeValue(textBox19.Text);
            row.Description2 = summarySheet.GetAttributeValue(textBox20.Text);
            row.function = summarySheet.Assignment + " " + summarySheet.Location;
            row.Revisions = Enumerable.Repeat("X", maxRevisions).ToList();

            rows.Insert(0, row);
            if (!rows.Contains(row))
                rows.Add(row);
            //place headers

            foreach (var revisions in Revisions.Values)
            {
                revisions.Sort();
                foreach (var revision in revisions)
                {
                    E3Graph graph5 = job.CreateGraph();
                    graph5.CreateText(summarySheet.Id, revision, (double)numericUpDown9.Value + X_SpacingTitle, (double)numericUpDown1.Value);
                    graph5.SetTextHeight((double)numericUpDown12.Value);
                    graph5.SetTextStyle(1);
                    X_SpacingTitle += 8;
                }
            }
            var sortedRows = rows;//.OrderBy(r => r.sheet).ToList();

              for (int j = 0; j < sortedRows.Count(); j++)
            
            {
                int rowCounter = 0;

                var _row = sortedRows[j];
               
                rowCounter++;
                double X_Spacing = 0;
                E3Graph graph = job.CreateGraph();

                graph.CreateText(summarySheet.Id, _row.sheet, (double)numericUpDown3.Value, (double)numericUpDown1.Value - Y_Spacing);
                graph.SetTextStyle(1);
                graph.SetTextHeight((double)numericUpDown12.Value);
                E3Graph graph2 = job.CreateGraph();

                graph2.CreateText(summarySheet.Id, _row.Description1, (double)numericUpDown4.Value, (double)numericUpDown1.Value - Y_Spacing);
                graph2.SetTextStyle(1);
                graph2.SetTextHeight((double)numericUpDown12.Value);
                E3Graph graph3 = job.CreateGraph();
                graph3.CreateText(summarySheet.Id, _row.Description2, (double)numericUpDown5.Value + 50, (double)numericUpDown1.Value - Y_Spacing);
                graph3.SetTextStyle(1);
                graph3.SetTextHeight((double)numericUpDown12.Value);

                E3Graph graph4 = job.CreateGraph();
                graph4.CreateText(summarySheet.Id, _row.function, (double)numericUpDown6.Value, (double)numericUpDown1.Value - Y_Spacing);
                graph4.SetTextHeight((double)numericUpDown12.Value);
                graph4.SetTextStyle(1);


                foreach (var revision in _row.Revisions)
                {
                    if (revision == null || revision == "")
                    {
                        continue;
                    }

                    else
                    {
                        E3Graph graph5 = job.CreateGraph();
                        graph5.CreateText(summarySheet.Id, "X", (double)numericUpDown7.Value + X_Spacing, (double)numericUpDown1.Value - Y_Spacing);
                        graph5.SetTextHeight((double)numericUpDown12.Value);
                        graph5.SetTextStyle(1);

                    }
                    X_Spacing += 8;
                }
                Y_Spacing += (double)numericUpDown10.Value;
                // Créer une nouvelle feuille toutes les 32 lignes
                decimal index = numericUpDown14.Value;
            }

        }
        private void OnePrjNumCreateSheet()
        {
            List<int> sheetIds = new List<int>();

            // Iterate over prjNum
            using (E3Job job = E3Job.Instances.First())
            using (E3Sheet summarySheet = job.CreateSheet())
            {
                var sheetNumber = (int)numericUpDown14.Value;
                string baseSheetName = textBox15.Text;
                double Y_Spacing = (double)numericUpDown10.Value;

                // Load the Data (fill all lists)
                LoadData(null);

                if (!revisionFound)
                {
                    job.Application.PutInfo(1, "No Revisions found in this project.");
                    return;
                }

                if (rows.Count < numericUpDown11.Value)
                {
                    string sheetName = baseSheetName + sheetNumber;
                    createOneSheet(null, job, sheetName);
                    sheetNumber++;  // Increment sheetNumber after creating the sheet
                    return;
                }

                double NumberOfSheetsToCreate = Math.Ceiling(rows.Count() / ((double)numericUpDown11.Value - 1));

                for (int i = 0; i < NumberOfSheetsToCreate; i++)
                {
                    int maxRevisions = rows.Max(row => row.Revisions.Count);

                    // Create the first summary row with the maximum number of "X"
                    SummaryRow firstSummaryRow = new SummaryRow
                    {
                        sheet = baseSheetName + sheetNumber,
                        Description1 = summarySheet.GetAttributeValue(textBox19.Text),
                        Revisions = Enumerable.Repeat("X", maxRevisions).ToList()
                    };

                    // Add the first summary row to the rows list
                    rows.Insert(0, firstSummaryRow);

                    if (!rows.Contains(firstSummaryRow))
                        rows.Add(firstSummaryRow);

                    sheetNumber++;  // Increment sheetNumber after adding the first summary row
                }

                var sortedRows = rows.OrderBy(row => ExtractNumber(row.sheet)).ToList();

                double X_Spacing = 0;

                for (int j = 0; j < sortedRows.Count; j++)
                {
                    var _row = sortedRows[j];

                    if (j % numericUpDown11.Value == 0)
                    {
                        string sheetName = baseSheetName + (sheetNumber - 2);
                        Y_Spacing = (double)numericUpDown10.Value;
                        CreateSummarySheet(job, summarySheet, sheetName, "");
                        sheetNumber++;  // Increment sheetNumber after creating the summary sheet
                    }

                    CreateGraph(job, summarySheet, (double)numericUpDown3.Value, Y_Spacing, _row.sheet);
                    CreateGraph(job, summarySheet, (double)numericUpDown4.Value, Y_Spacing, _row.Description1);
                    CreateGraph(job, summarySheet, (double)numericUpDown5.Value, Y_Spacing, _row.Description2);
                    CreateGraph(job, summarySheet, (double)numericUpDown6.Value, Y_Spacing, _row.function);

                    double spaceBetweenX = (double)numericUpDown7.Value;
                    foreach (var revision in _row.Revisions)
                    {
                        if (string.IsNullOrEmpty(revision))
                            continue;

                        CreateGraph(job, summarySheet, spaceBetweenX, Y_Spacing, "X");
                        X_Spacing += 8;
                        spaceBetweenX += 8;
                    }

                    Y_Spacing += (double)numericUpDown10.Value;
                }

                job.Application.PutInfo(0, "<EDSummaryIndex ended>");
            }
        }


        public void CreateMultipleProjectSheet()
        {
            List<int> sheetIds = new List<int>();

            // Iterate over prjNum
            using (E3Job job = E3Job.Instances.First())
            using (E3Sheet summarySheet = job.CreateSheet())
            {
                var distinctPrjNums = job.Sheets.Select(s => s.GetAttributeValue(textBox18.Text)).Distinct().ToList();

                foreach (var prjNum in distinctPrjNums)
                {
                    var sheetNumber = (int)numericUpDown14.Value;
                    string baseSheetName = textBox15.Text;
                    double Y_Spacing = (double)numericUpDown10.Value;

                    // Load the Data (fill all lists)
                    LoadData(prjNum);

                    if (!revisionFound)
                    {
                        job.Application.PutInfo(1, "No Revisions found in this project.");
                        return;
                    }

                    if (rows.Count < numericUpDown11.Value)
                    {
                        string sheetName = baseSheetName + sheetNumber;
                        createOneSheet(prjNum, job, sheetName);
                        sheetNumber++;  // Increment sheetNumber after creating the sheet
                        continue;
                    }

                    double NumberOfSheetsToCreate = Math.Ceiling(rows.Count() / ((double)numericUpDown11.Value - 1));

                    for (int i = 0; i < NumberOfSheetsToCreate; i++)
                    {
                        int maxRevisions = rows.Max(row => row.Revisions.Count);

                        // Create the first summary row with the maximum number of "X"
                        SummaryRow firstSummaryRow = new SummaryRow
                        {
                            sheet = baseSheetName + sheetNumber,
                            Description1 = "sommaire",
                            prjNum = prjNum,
                            Revisions = Enumerable.Repeat("X", maxRevisions).ToList()
                        };

                        // Add the first summary row to the rows list
                        rows.Insert(0, firstSummaryRow);

                        if (!rows.Contains(firstSummaryRow))
                            rows.Add(firstSummaryRow);

                        sheetNumber++;  // Increment sheetNumber after adding the first summary row
                    }

                    var sortedRows = rows.OrderBy(row => ExtractNumber(row.sheet)).ToList();

                    double X_Spacing = 0;

                    for (int j = 0; j < sortedRows.Count; j++)
                    {
                        var _row = sortedRows[j];

                        if (j % numericUpDown11.Value == 0)
                        {
                            string sheetName = baseSheetName +( sheetNumber -2);
                            Y_Spacing = (double)numericUpDown10.Value;
                            CreateSummarySheet(job, summarySheet, sheetName, prjNum);
                            sheetNumber++;  // Increment sheetNumber after creating the summary sheet
                        }

                        CreateGraph(job, summarySheet, (double)numericUpDown3.Value, Y_Spacing, _row.sheet);
                        CreateGraph(job, summarySheet, (double)numericUpDown4.Value, Y_Spacing, _row.Description1);
                        CreateGraph(job, summarySheet, (double)numericUpDown5.Value, Y_Spacing, _row.Description2);
                        CreateGraph(job, summarySheet, (double)numericUpDown6.Value, Y_Spacing, _row.function);

                        double spaceBetweenX = (double)numericUpDown7.Value;
                        foreach (var revision in _row.Revisions)
                        {
                            if (string.IsNullOrEmpty(revision))
                                continue;

                            CreateGraph(job, summarySheet, spaceBetweenX, Y_Spacing, "X");
                            X_Spacing += 8;
                            spaceBetweenX += 8;
                        }

                        Y_Spacing += (double)numericUpDown10.Value;
                    }

                    job.Application.PutInfo(0, "<EDSummaryIndex ended>");
                }
            }
        }

        private static int ExtractNumber(string sheet)
        {
            var match = Regex.Match(sheet, @"\d+");
            return match.Success ? int.Parse(match.Value) : 0;
        }


        private void CreateGraph(E3Job job, E3Sheet summarySheet, double X, double Y, string text)
        {
            using (E3Graph graph = job.CreateGraph())
            {

                graph.CreateText(summarySheet.Id, text, X, (double)numericUpDown1.Value - Y);
                graph.SetTextStyle(1);
                graph.SetTextHeight((double)numericUpDown12.Value);
            }

        }

        private void CreateSummarySheet(E3Job job, E3Sheet summarySheet, string sheetName, string prjNum)
        {
            double X_SpacingTitle = 8;
            summarySheet.Create(sheetName, textBox13.Text);
            summarySheet.SetAttributeValue(textBox18.Text, prjNum);
            summarySheet.SetAttributeValue(textBox22.Text, Revisions.Values.Last().Last());

            if (summarySheet.Id == 0)
            {
                MessageBox.Show("Sheet could not be created.");
                return;
            }
            //add them to rows
            job.Application.PutInfo(0, sheetName + " created.");
            foreach (var revisions in Revisions.Values)
            {
               // revisions.Sort();
                foreach (var revision in revisions)
                {
                    E3Graph graph5 = job.CreateGraph();
                    graph5.CreateText(summarySheet.Id, revision, (double)numericUpDown9.Value + X_SpacingTitle, (double)numericUpDown1.Value);
                    graph5.SetTextHeight((double)numericUpDown12.Value);
                    graph5.SetTextStyle(1);
                    X_SpacingTitle += 8;
                }
            }
        }

        private void textBox13_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox14_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox15_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox16_Leave(object sender, EventArgs e)
        {
            Save();

        }
        private void textBox17_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox18_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox19_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox20_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void textBox21_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown1_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown2_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown3_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown4_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown5_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown6_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown7_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown8_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown9_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown10_Leave(object sender, EventArgs e)
        {
            Save();

        }
        private void numericUpDown11_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void numericUpDown12_Leave(object sender, EventArgs e)
        {
            Save();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Save();
            DialogResult result = MessageBox.Show("Old Summary sheets will be replaced, do you want to continue?", "Continue", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (string.IsNullOrEmpty(textBox13.Text) || string.IsNullOrEmpty(textBox14.Text) || string.IsNullOrEmpty(textBox19.Text) || string.IsNullOrEmpty(textBox20.Text) || string.IsNullOrEmpty(textBox21.Text))
            {
                MessageBox.Show("You must fill every input box.");
            }
            // Vérifier la réponse de l'utilisateur
            if (result == DialogResult.Yes)
            {
                if (checkBox2.Checked == false)
                {
                    using (E3Job job = E3Job.Instances.First())
                    {
                        job.Application.PutInfo(0, "<Starting EDSummaryByIndex>");
                        foreach (var sheet in job.Sheets)
                        {
                            if (sheet.GetAttributeValue(textBox19.Text) == "sommaire")
                            {
                                sheet.Delete();
                            }
                        }
                    }
                    OnePrjNumCreateSheet();

                }
                else
                {
                    using (E3Job job = E3Job.Instances.First())
                    {
                        job.Application.PutInfo(0, "<Starting EDSummaryByIndex>");
                        foreach (var sheet in job.Sheets)
                        {
                            if (sheet.GetAttributeValue(textBox19.Text) == "sommaire")
                            {
                                sheet.Delete();
                            }
                        }
                    }
                    CreateMultipleProjectSheet();
                }
            }
            else
            {
                return;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Save();
        }

        private void numericUpDown13_Leave(object sender, EventArgs e)
        {
            Save();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Save();

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                textBox18.Enabled = true;
            }
            else
            {
                textBox18.Enabled = false;
            }
        }
    }
}
public class RowComparer : IComparer<SummaryRow>
{
    public int Compare(SummaryRow x, SummaryRow y)
    {
        // Assurez-vous de gérer les cas où x ou y pourraient être null, si nécessaire.

        // Comparez les noms de feuille (sheet.name)
        return string.Compare(x.sheet, y.sheet, StringComparison.Ordinal);
    }
}