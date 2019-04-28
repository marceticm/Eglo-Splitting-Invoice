using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace Stavljanje_na_stanje
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public class Radnja
    {
        public string Naziv { get; set; }
        public int CustomerNumber { get; set; }

    }
    // kako bi bilo da korisnik moze da doda kupca i broj za komisioniranje
    public partial class MainWindow : System.Windows.Window
    {
        private List<string> sifre = new List<string>();
        private List<string> kolicine = new List<string>();
        private List<string> opisi = new List<string>();
        private List<string> porudzbenice = new List<string>();
        private string fileName;
        private bool postojiRadnja;
        private const int dusanovacBroj = 551272;
        private const int nisBroj = 551157;
        private const int bulevarBroj = 551136;
        private const int zmajBroj = 551129;
        private const int sabacBroj = 551236;
        private const int pazarBroj = 551062;
        private const int karaburmaBroj = 551275;
        private const int attriumBroj = 551226;
        private const int halogenBroj = 550897;
        private const int halogenKisackaBroj = 551152;
        private const int emmezetaBroj = 551221;
        private const int eltonBroj = 551147;
        private const int nortecBroj = 551205;
        private const int elektrostarBroj = 551351;
        private const int homeLightDecorBroj = 551127;
        private const int elektroMirkoBroj = 551355;
        private const int rgbZrenjaninBroj = 551500;
        private const int joluxBroj = 551162;
        private const int nexalBroj = 551607;
        private const int rgbSpensBroj = 551113;
        private const int iluminaBroj = 551670;
        private const int teaLightBroj = 551550;
        private const int nesaElektroBroj = 551120;
        private const int vistaBroj = 551190;
        private const int greenDesignBroj = 551151;
        private const int xlBroj = 551211;
        private const int lazarBroj = 551517;
        private const int moneroBroj = 551142;
        private const int isterBroj = 551160;
        private const int staisBroj = 551196;
        private const int lutzDecoBroj = 551530;
        private const int lutzShopBroj = 551126;
        private const int lutzLagerBroj = 551133;

        private string[] imena = new string[]
        {
                 "Dusanovac", "Nis", "Bulevar", "Zmaj", "Sabac", "Novi Pazar", "Karaburma",
                  "Attrium", "Halogen", "Halogen Kisacka", "Emmezeta", "Elton",
                  "Nortec", "Elektrostar", "Home Light Decor", "Elektro Mirko",
                  "RGB Zrenjanin", "Jolux", "Nexal", "RGB Spens", "Ilumina", "Tea Light",
                  "Nesaelektro", "Vista", "Green Design", "XL Prostor", "Lazar Group",
                  "Monero", "Ister", "Stais", "LUTZ DECO (AE8)", "LUTZ SHOP (NS)", "LUTZ LAGER (ACZ)"
        };

        public MainWindow()
        {
            InitializeComponent();
            btnSacuvajFajl.Visibility = Visibility.Hidden;
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var encoding = Encoding.UTF8;
            btnSacuvajFajl.Visibility = Visibility.Hidden;
            rdbBezOpisa.Visibility = Visibility.Hidden;
            rdbOpis.Visibility = Visibility.Hidden;
            MessageBox.Show("Operacija pocinje, molim sacekajte...");

            foreach (var ime in imena)
            {
                sifre = new List<string>();
                kolicine = new List<string>();
                opisi = new List<string>();
                porudzbenice = new List<string>();
                // ovo ubaciti dole u using ili while
                postojiRadnja = false;
                using (var reader = new StreamReader(fileName, encoding))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(new char[] { ';' });

                        var radnja = new Radnja();

                        radnja.Naziv = ime;
                        // zameniti nekako ovu grdosiju, da cita iz nekog niza
                        switch (radnja.Naziv)
                        {
                            case "Dusanovac": radnja.CustomerNumber = dusanovacBroj; break;
                            case "Nis": radnja.CustomerNumber = nisBroj; break;
                            case "Bulevar": radnja.CustomerNumber = bulevarBroj; break;
                            case "Zmaj": radnja.CustomerNumber = zmajBroj; break;
                            case "Sabac": radnja.CustomerNumber = sabacBroj; break;
                            case "Novi Pazar": radnja.CustomerNumber = pazarBroj; break;
                            case "Karaburma": radnja.CustomerNumber = karaburmaBroj; break;
                            case "Attrium": radnja.CustomerNumber = attriumBroj; break;
                            case "Halogen": radnja.CustomerNumber = halogenBroj; break;
                            case "Halogen Kisacka": radnja.CustomerNumber = halogenKisackaBroj; break;
                            case "Emmezeta": radnja.CustomerNumber = emmezetaBroj; break;
                            case "Elton": radnja.CustomerNumber = eltonBroj; break;
                            case "Nortec": radnja.CustomerNumber = nortecBroj; break;
                            case "Elektrostar": radnja.CustomerNumber = elektrostarBroj; break;
                            case "Home Light Decor": radnja.CustomerNumber = homeLightDecorBroj; break;
                            case "Elektro Mirko": radnja.CustomerNumber = elektroMirkoBroj; break;
                            case "RGB Zrenjanin": radnja.CustomerNumber = rgbZrenjaninBroj; break;
                            case "Jolux": radnja.CustomerNumber = joluxBroj; break;
                            case "Nexal": radnja.CustomerNumber = nexalBroj; break;
                            case "RGB Spens": radnja.CustomerNumber = rgbSpensBroj; break;
                            case "Ilumina": radnja.CustomerNumber = iluminaBroj; break;
                            case "Tea Light": radnja.CustomerNumber = teaLightBroj; break;
                            case "Nesaelektro": radnja.CustomerNumber = nesaElektroBroj; break;
                            case "Vista": radnja.CustomerNumber = vistaBroj; break;
                            case "Green Design": radnja.CustomerNumber = greenDesignBroj; break;
                            case "XL Prostor": radnja.CustomerNumber = xlBroj; break;
                            case "Lazar Group": radnja.CustomerNumber = lazarBroj; break;
                            case "Monero": radnja.CustomerNumber = moneroBroj; break;
                            case "Ister": radnja.CustomerNumber = isterBroj; break;
                            case "Stais": radnja.CustomerNumber = staisBroj; break;
                            case "LUTZ DECO (AE8)": radnja.CustomerNumber = lutzDecoBroj; break;
                            case "LUTZ SHOP (NS)": radnja.CustomerNumber = lutzShopBroj; break;
                            case "LUTZ LAGER (ACZ)": radnja.CustomerNumber = lutzLagerBroj; break;
                            default:
                                break;
                        }

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
                // da ne izbacuje prazne eksele ili break?
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

                    wb.Close();
                    excel.Quit();
                    //wb.Close();
                }
            }

            // btnSacuvajFajl.Visibility = Visibility.Hidden;
            // rdbBezOpisa.Visibility = Visibility.Hidden;
            // rdbOpis.Visibility = Visibility.Hidden;
            MessageBox.Show("Operacija uspela!\nSpiskovi su sacuvani.");
            // System.Diagnostics.Process.Start("explorer.exe", string.Format("/select,\"{0}\"", saveFileDialog1.FileName));
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
