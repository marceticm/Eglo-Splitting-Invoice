using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Stavljanje_na_stanje
{
    public class Radnja
    {
        public string Naziv { get; set; }
        public int CustomerNumber { get; set; }
    }
    public partial class MainWindow : System.Windows.Window
    {
        IDictionary<int, string> dict = new Dictionary<int, string>();
        private List<string> sifre = new List<string>();
        private List<string> kolicine = new List<string>();
        private List<string> opisi = new List<string>();
        private List<string> porudzbenice = new List<string>();
        private string fileName;
        private bool postojiRadnja;
        string ime = "";

        public MainWindow()
        {
            InitializeComponent();
            btnSacuvajFajl.Visibility = Visibility.Hidden;
            loadingAnimation.Visibility = Visibility.Hidden;
            lblObrada.Visibility = Visibility.Hidden;

            string[] tekst = new string[] { };
            string txtFileName = "KupciZaProgram.csv";
            if (File.Exists(txtFileName))
            {
                tekst = File.ReadAllLines(txtFileName);
            }
            else
            {
                MessageBox.Show("Datoteka sa brojevima za komisioniranje ne postoji!");
            }
            for (int i = 0; i < tekst.Length; i++)
            {
                var linija = tekst[i];
                var splitovanaLinija = linija.Split(';');
                dict.Add(Convert.ToInt32(splitovanaLinija[0]), splitovanaLinija[1]);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() ?? false)
            {
                fileName = openFileDialog1.FileName;

                if (fileName.Substring(fileName.Length - 4) != ".csv")
                {
                    MessageBox.Show("Greska: Niste odabrali .CSV fajl");
                    return;
                }
            }
            btnUvuciFakturu.Visibility = Visibility.Hidden;
            btnSacuvajFajl.Visibility = Visibility.Visible;
            rdbBezOpisa.Visibility = Visibility.Visible;
            rdbOpis.Visibility = Visibility.Visible;
        }

        public void Splitting(string fileName)
        {
            foreach (var par in dict)
            {
                sifre = new List<string>();
                kolicine = new List<string>();
                opisi = new List<string>();
                porudzbenice = new List<string>();
                // ovo ubaciti dole u using ili while
                postojiRadnja = false;

                try
                {
                    using (var reader = new StreamReader(fileName, Encoding.UTF8))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(new char[] { ';' });

                            var radnja = new Radnja();

                            ime = par.Value;
                            radnja.Naziv = ime;
                            radnja.CustomerNumber = par.Key;

                            var broj = values[12];
                            var orderName = values[9];

                            bool sadrziFZ(string value)
                            {
                                if (value.Contains("fz") || value.Contains("FZ") || value.Contains("fZ") || value.Contains("Fz"))
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }

                            if (broj.Equals(radnja.CustomerNumber.ToString()))
                            {
                                postojiRadnja = true;
                                if (sadrziFZ(orderName))
                                {
                                    continue;
                                }
                                else
                                {
                                    sifre.Add(values[2]);
                                    kolicine.Add(values[14]);
                                    opisi.Add(values[3]);
                                    porudzbenice.Add(values[9]);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Doslo je do greske:\n\n" + ex.Message +
                        "\n\nAplikacija ce se restartovati.", "UPOZORENJE!");
                    System.Diagnostics.Process.Start(System.Windows.Application.ResourceAssembly.Location);
                    this.Dispatcher.Invoke(() =>
                    {
                        System.Windows.Application.Current.Shutdown();
                    });
                }

                if (!postojiRadnja)
                {
                    continue;
                }
                else
                {
                    var excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = false;
                    Workbook wb = excel.Workbooks.Add();
                    // var sheets = wb.Worksheets;
                    sifre = sifre.Select(c => c.Replace("-", string.Empty)).ToList();
                    sifre = sifre.Select(c => ubacujeRazmak(c)).ToList();

                    string ubacujeRazmak(string ulaz)
                    {
                        for (int i = 0; i < ulaz.Length; i++)
                        {
                            if (char.IsDigit(ulaz[i]))
                            {
                                ulaz = ulaz.Insert(i, " ");
                                return ulaz;
                            }
                        }
                        return ulaz;
                    }
                    //sheets[1].Cells.ClearContents();

                    this.Dispatcher.Invoke(() =>
                    {
                        for (int i = 1; i <= sifre.Count; i++)
                        {
                            if (rdbBezOpisa.IsChecked == true)
                            {
                                excel.Cells[i, 1].Value2 = sifre[i - 1];
                                excel.Cells[i, 2].Value2 = kolicine[i - 1];
                                excel.Cells[i, 3].Value2 = porudzbenice[i - 1];
                            }
                            else
                            {
                                excel.Cells[i, 1].Value2 = sifre[i - 1];
                                excel.Cells[i, 2].Value2 = opisi[i - 1];
                                excel.Cells[i, 3].Value2 = kolicine[i - 1];
                            }
                        }
                        excel.Columns.AutoFit();

                        if (rdbBezOpisa.IsChecked == true)
                        {
                            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                           + @"\Dokumenta " + DateTime.Now.ToShortDateString() + @"\");

                            wb.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                + @"\Dokumenta " + DateTime.Now.ToShortDateString() + @"\"
                                + ime + " " + DateTime.Now.ToShortDateString() + ".xlsx");
                        }
                        else
                        {
                            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                           + @"\Dokumenta Sa Opisom " + DateTime.Now.ToShortDateString() + @"\");

                            wb.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                + @"\Dokumenta Sa Opisom " + DateTime.Now.ToShortDateString() + @"\"
                                + ime + " " + DateTime.Now.ToShortDateString() + ".xlsx");
                        }
                    });

                    wb.Close();
                    excel.Quit();
                    //wb.Close();
                }
            }
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var encoding = Encoding.UTF8;
            btnSacuvajFajl.Visibility = Visibility.Hidden;
            rdbBezOpisa.Visibility = Visibility.Hidden;
            rdbOpis.Visibility = Visibility.Hidden;
            // MessageBox.Show("Operacija pocinje, molim sacekajte...");

            Task task = new Task(() => Splitting(fileName));
            task.Start();
            loadingAnimation.Visibility = Visibility.Visible;
            lblObrada.Visibility = Visibility.Visible;
            await task;
            lblObrada.Visibility = Visibility.Hidden;
            loadingAnimation.Visibility = Visibility.Hidden;
            MessageBox.Show("Operacija uspela!\nSpiskovi su sacuvani.");
        }


        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Windows.Application.ResourceAssembly.Location);
            System.Windows.Application.Current.Shutdown();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
