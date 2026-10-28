using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CDNBase;
using ExcelDataReader;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using CDNHlmn;

namespace ONE032
{
    public partial class ONE032WIN : Form
    {
       // public static string sKnt_Kod = "OSTREWKA";
        public static string Operator = "ADMIN";
        public static string Haslo = "Platforma2020!";
        public static string Firma = "PPBW_PPBW";
        public static string Baza = "[CDN_PPBW_PPBW].[CDN].";

       /* public static string sKnt_Kod = "!NIEOKREŚLONY!";
        public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
        public static string Firma = "PREZENTACJA_KH";
        public static string Baza = "[CDN_PREZENTACJA_KH_1].[CDN].";*/
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;
        public ONE032WIN()
        {
            InitializeComponent();
        }

        #region funkcje optimy

       

        static protected void LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,   0,  1, 1 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            Application.HASPServerCommandLine = "41941-FE3B2-002::5000139271";
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }
        #endregion

        void WyjscieZApki(string a)
        {
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                MessageBox.Show("Zakończono przetwarzanie pliku.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Generowanie dokumentów zakończyło się niepowodzeniem.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            #region excel5
            Cursor.Current = Cursors.WaitCursor;
            if (tB1.Text == "nie wybrano pliku")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                dataTable = dataSet.Tables[0];
            }

            //Przypisywanie danych z datasetu do listy przyszłych pozycji.
            //List<ONE032Data> listaPozycjiPz = new List<ONE032Data>();
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Pusty arkusz", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return;
            }




          
            #endregion

            #region dzialanie na optimie
            #region logowanie
            //Logowanie do O! bez wyświetlania okienka logowania
            Cursor.Current = Cursors.WaitCursor;

            //JakaBaza("klient");

            try
            {
                LogowanieAutomatyczne();
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                    MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                    MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //MessageBox.Show("Zalogowano ", "Komunikat",
            // MessageBoxButtons.OK, MessageBoxIcon.Error);
            #endregion



            CDNBase.AdoSession Sesja = Login.CreateSession();
            #endregion

            dodawanieDekretu(dataTable, Sesja);


            //zapisujemy
            //Sesja.Save();

            WyjscieZApki("done");


        }

        #region wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion

        #region dekret

        private void dodawanieDekretu(DataTable dataTable, AdoSession sesja)
        {

            try
            {
                var rOtwartyOkres = sesja.CreateObject("CDN.Okresy").Item("OOb_Symbol='"+ Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][0]).Year+"'");

                int ID_OkresuObr = rOtwartyOkres.ID;

                List<string> miesiace = new List<string>();

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    string miesiac = dataTable.Rows[i][11].ToString().Substring(dataTable.Rows[i][11].ToString().IndexOf("/") - 2, 7);
                    if (!miesiace.Contains(miesiac))
                        miesiace.Add(miesiac);
                }

                int k = 0;
                pBar1.Visible = true;
                // Set Minimum to 1 to represent the first file being copied.
                pBar1.Minimum = 1;
                // Set Maximum to the total number of files to copy.
                pBar1.Maximum = miesiace.Count+1;
                // Set the initial value of the ProgressBar.
                pBar1.Value = 1;
                // Set the Step property to a value of 1 to represent each file being copied.
                pBar1.Step = 1;

                foreach (string dat in miesiace)
                {
                    try
                    {
                        pBar1.PerformStep();
                        CDNBase.AdoSession Sesja = Login.CreateSession();
                        var rDekret = Sesja.CreateObject("CDN.DEKRETY").AddNew(null);

                        // Id dziennika 
                        rDekret.OObID = ID_OkresuObr;

                        string IdDziennik_Query = "SELECT Dzi_DziId FROM cdn.Dzienniki WHERE Dzi_OObId = " + ID_OkresuObr + " and Dzi_Nazwa = 'INNE'";

                        object dummy = Type.Missing;
                        var wynikQuery = Sesja.Connection.Execute(IdDziennik_Query, out dummy, 0);

                        int ID_Dziennik = 0;

                        if (!wynikQuery.EOF)
                            ID_Dziennik = Convert.ToInt32(wynikQuery.GetString());
                        else
                        {
                            MessageBox.Show("Brak dziennika inne dla okresu obrachunkowego z pliku", "Komunikat",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                       int dzien= DateTime.DaysInMonth(Convert.ToInt32(dat.Substring(3, 4)), Convert.ToInt32(dat.Substring(0, 2)));
                       string dataDok= dat.Substring(3, 4)+"-"+dat.Substring(0, 2)+"-"+dzien.ToString();

                        rDekret.DziID = ID_Dziennik;

                        rDekret.Bufor = 1;
                        rDekret.DataDok = Convert.ToDateTime(dataDok);//Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][0]);
                        rDekret.Dokument = "Rozliczenie projektów (konta pozabilansowe) " + dataDok;//rDekret.DataDok.ToString("yyyy-MM-dd");
                        rDekret.Opis = "Rozliczenie projektów (konta pozabilansowe) " + dataDok;//rDekret.DataDok.ToString("yyyy-MM-dd");




                        

                        var rPozycje = rDekret.Elementy;

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            if (dataTable.Rows[i][11].ToString().Contains(dat))
                            {
                                var rPozycja = rPozycje.AddNew();

                                var rKontaWN = Sesja.CreateObject("CDN.Konta").Item("ACC_OOBid =" + ID_OkresuObr + "and  ACC_Numer='901'");

                                rPozycja.KontoWN = rKontaWN;

                                var rKontaMA = Sesja.CreateObject("CDN.Konta").Item("ACC_OOBid =" + ID_OkresuObr + "and ACC_Numer='902'");

                                rPozycja.KontoMA = rKontaMA;
                                rPozycja.Kwota = Math.Round(Convert.ToDouble(dataTable.Rows[i][6]), 2, MidpointRounding.AwayFromZero);


                                rPozycja.Opis = dataTable.Rows[i][7].ToString() + " / " + dataTable.Rows[i][10].ToString() + " / " + dataTable.Rows[i][11].ToString();
                            }

                        }

                        Sesja.Save();
                        Sesja.Dispose();
                       
                    }

                    catch (Exception es)
                    {
                        MessageBox.Show("Błąd w dodawaniu dekretu " +dat+": "+ Convert.ToString(es), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                pBar1.Visible = false;
            }
            catch(Exception e)
            {
               /* if(e.ToString().Contains("current financial period"))
                {
                    MessageBox.Show("Zmień okres obrachunkowy na: " + Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][0]).Year, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }*/
                MessageBox.Show("Błąd w dodawaniu dekretu: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion

        #region WZ
        void dodawanieWZ(DataTable dataTable, CDNBase.AdoSession Sesja)
        {
            List<string> akronimy = new List<string>();
            List<DateTime> data = new List<DateTime>();
            string akronim = "";
            for (int i = 2; i < dataTable.Rows.Count; i++)
            {
                if (Convert.ToString(dataTable.Rows[i][7]) == "")
                    akronim = Convert.ToString(dataTable.Rows[i][9]);
                if (!akronimy.Contains(akronim) && akronim != "")
                {
                    akronimy.Add(akronim);
                    data.Add(Convert.ToDateTime(dataTable.Rows[i][8]));
                }

            }

            string akronimRow = Convert.ToString(dataTable.Rows[2][12]);
            for (int i = 3; i < dataTable.Rows.Count; i++)
            {
                if (Convert.ToString(dataTable.Rows[i][12])!= akronimRow)
                {
                    MessageBox.Show("W pliku jest więcej niż jeden akronim", "Komunikat",
                                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }


            int k = 0;
            pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar1.Maximum = akronimy.Count + 1;
            // Set the initial value of the ProgressBar.
            pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar1.Step = 1;

            foreach (string aAkronimy in akronimy)
            {
                try
                {
                    pBar1.PerformStep();
                    Sesja = Login.CreateSession();

                    //MessageBox.Show(aAkronimy, "Komunikat",
                    //MessageBoxButtons.OK, MessageBoxIcon.Error);

                    string atrWartosc = aAkronimy;
                    object dummy = Type.Missing;
                    var sqlQuery1 = "SELECT TOP(1) [DAt_WartoscTxt]  FROM " + Baza + "[DokAtrybuty] where [DAt_WartoscTxt] ='" + atrWartosc + "'";
                    var wynikQuery = Sesja.Connection.Execute(sqlQuery1, out dummy, 0);

                    List<string> kody = new List<string>();
                    string kod = "";
                    for (int i = 2; i < dataTable.Rows.Count; i++)
                    {
                        if (Convert.ToString(dataTable.Rows[i][9]) == aAkronimy)
                        {
                            kod = Convert.ToString(dataTable.Rows[i][0]);
                            if (!kody.Contains(kod))
                            {
                                kody.Add(kod);
                                //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                                // MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }

                    string zly_kod = "";
                    bool dobre_kody = true;
                    foreach (string aKody in kody)
                    {
                        var sqlQuery2 = "SELECT TOP(1) [Twr_Kod] FROM " + Baza + "[Towary] where [Twr_Kod] ='" + aKody + "'";
                        var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                        //MessageBox.Show("Kod: " + aKody , "Komunikat",
                        //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (wynikQuery2.EOF)
                        {
                            dobre_kody = false;
                            zly_kod = aKody;
                            //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                            //MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        }
                    }

                    //var sqlQuery3 = "SELECT TOP(1) [Knt_Kod]  FROM " + Baza + "[Kontrahenci] where [Knt_Kod] ='" + sKnt_Kod + "'";
                    //var wynikQuery3 = Sesja.Connection.Execute(sqlQuery3, out dummy, 0);

                    try
                    {

                        if (wynikQuery.EOF && dobre_kody == true )
                        {


                            CDNHlmn.DokumentyHaMag PeZety = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
                            CDNHlmn.IDokumentHaMag PeZet = (CDNHlmn.IDokumentHaMag)PeZety.AddNew(null);

                            CDNTwrb1.Towary zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));


                            var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
                            var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='" + akronimRow + "'"];   //Tutaj kod ostrawki

                            //dodac typ WZ
                            PeZet.Rodzaj = 306000;
                            PeZet.TypDokumentu = 306;
                            PeZet.Bufor = 1;
                            //PeZet.NumerObcy ="";

                            //Ustawiamy date
                            PeZet.DataDok = data[k];
                            PeZet.DataWys = data[k];
                            PeZet.Podmiot = (CDNHeal.IPodmiot)kontrahent;


                            //atrybut dokumentu
                            string dea_kod_nr_dok = "NR DOKUMENTU FORCOM";
                            var AtrDokument = PeZet.Atrybuty;
                            var rAtrybut = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);
                            var Atr = AtrDokument.AddNew(null);
                            Atr.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybut["dea_Kod = '" + dea_kod_nr_dok + "'"];
                            Atr.Wartosc = atrWartosc;

                            

                            //List<string> kody = new List<string>();
                            //string kod = "";


                            Sesja.Save();
                            var id_PZ = PeZet.ID;
                            int c = 0;

                            foreach (string aKody in kody)
                            {

                                List<string> ceny = new List<string>();
                                string cena = "";
                                for (int i = 2; i < dataTable.Rows.Count; i++)
                                {
                                    if (Convert.ToString(dataTable.Rows[i][9]) == aAkronimy && Convert.ToString(dataTable.Rows[i][0])==aKody )
                                    {
                                        cena = Convert.ToString(dataTable.Rows[i][11]);
                                        if (!ceny.Contains(cena))
                                        {
                                            ceny.Add(cena);
                                            //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                                            // MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        }
                                    }
                                }

                                foreach (string aCeny in ceny)
                                {

                                    try
                                    {

                                        Sesja = Login.CreateSession();
                                        PeZet = Sesja.CreateObject("CDN.DokumentyHaMag").Item("TrN_TrNId=" + id_PZ);
                                        CDNBase.ICollection pozycje = PeZet.Elementy;
                                        var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);
                                        //double count = 0.0;
                                        /*for (int a = 2; a < dataTable.Rows.Count; a++)
                                        {
                                            if (Convert.ToString(dataTable.Rows[a][0]) == aKody && Convert.ToString(dataTable.Rows[a][9]) == aAkronimy)
                                                count = count + Convert.ToDouble(dataTable.Rows[a][2]);
                                        }*/
                                        int count = 0;
                                        for (int a = 2; a < dataTable.Rows.Count; a++)
                                        {
                                            if (Convert.ToString(dataTable.Rows[a][0]) == aKody && Convert.ToString(dataTable.Rows[a][9]) == aAkronimy && Convert.ToString(dataTable.Rows[a][11])==aCeny)
                                            {
                                                count++;
                                            }
                                        }
                                                //przypisywanie wartości pozycji
                                          CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_Kod='" + aKody + "'"];
                                        //MessageBox.Show(zTowar.Kod, "Komunikat",
                                        //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        pozycja.TowarKod = zTowar.Kod;
                                        pozycja.Ilosc = count;
                                        // pozycja.Cena0 = 0;
                                        // pozycja.Cena0WD = 0;
                                        // pozycja.WartoscNetto = 0;
                                        // pozycja.WartoscBrutto = 0;
                                        pozycja.Cena0 = Convert.ToDecimal(aCeny);
                                        

                                        //var zasob = pozycja.Zasoby;

                                        var rDostawy = pozycja.Dostawy;
                                        string dea_Kod_NR_SERYJNY = "NUMER SERYJNY";
                                        for (int a = 2; a < dataTable.Rows.Count; a++)
                                        {
                                            if (String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[a][3])) == false && Convert.ToString(dataTable.Rows[a][0]) == aKody && Convert.ToString(dataTable.Rows[a][9]) == aAkronimy
                                                && Convert.ToString(dataTable.Rows[a][11]) == aCeny)
                                            {
                                                try
                                                {
                                                    pozycja.UstawDostawy();
                                                    foreach (IDostawaElem dostawa in pozycja.Dostawy)
                                                    {
                                                        if (dostawa.Cecha1_Wartosc == Convert.ToString(dataTable.Rows[a][3]) && dostawa.IloscDostepna >= 1)
                                                        {
                                                            dostawa.Ilosc = 1.0;
                                                            break;
                                                        }

                                                    }
                                                    pozycja.UsunDostawy();
                                                    /* var rDostawa = rDostawy.AddNew();
                                                     pozycja.UstawDostawy();
                                                     rDostawa.Cecha1_Kod = dea_Kod_NR_SERYJNY;
                                                     rDostawa.Cecha1_Wartosc = Convert.ToString(dataTable.Rows[a][3]);
                                                     rDostawa.Ilosc = Convert.ToString(dataTable.Rows[a][2]);

                                                     if (dataTable.Columns.Count == 12 && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[a][11])) == false && Convert.ToString(dataTable.Rows[a][11]).Substring(0, 1) == "T")
                                                     {
                                                         rDostawa.Cecha2_Kod = "LEASING";
                                                         rDostawa.Cecha2_Wartosc = "Tak";
                                                         //MessageBox.Show("Dodawanie Leasingu", "Komunikat",
                                                         //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                     }

                                                     pozycja.UsunDostawy();*/
                                                }
                                                catch (Exception empty)
                                                {
                                                    MessageBox.Show("Błąd w dodawaniu zasobu: " + Convert.ToString(empty), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                }
                                            }
                                        }


                                        // pozycja.UsunDostawy();


                                        Sesja.Save();
                                        var ZasPozycji = pozycja.Atrybuty;
                                        var rAtrybuty = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);

                                        //usuwanie istniejących atrybutów
                                        string TrEId = "SELECT TOP(1) [TrA_TrEId] FROM " + Baza + "[TraElemAtr] WHERE [TrA_TrEId] =" + pozycja.ID; //( SELECT [TrE_TrEID] FROM "+Baza+"[TraElem] WHERE [TrE_TrNId] = ("
                                                                                                                                                   //+ "SELECT [DAt_TrNId] FROM "+Baza+"[DokAtrybuty] WHERE [DAt_WartoscTxt] = '"+aAkronimy+"') and [TrE_TwrKod] = '"+aKody+"')";
                                        var wynikExistQuery = Sesja.Connection.Execute(TrEId, out dummy, 0);
                                        while (!wynikExistQuery.EOF)
                                        {
                                            ZasPozycji.Delete(0);
                                            Sesja.Save();
                                            wynikExistQuery = Sesja.Connection.Execute(TrEId, out dummy, 0);
                                        }


                                        string dea_Kod_Gwarancja_MC = "GWARACJA MC";
                                        string AtrQuery2 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_Gwarancja_MC + "') and [TwA_TwrId] = (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        string dea_Kod_Gwarancja_Rodzaj = "GWARANCJA RODZAJ";
                                        string AtrQuery3 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_Gwarancja_Rodzaj + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        string dea_Kod_Stawka_Amorty = "STAWKA AMORTYZACJI";
                                        string AtrQuery4 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_Stawka_Amorty + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        string dea_Kod_Typ_Indeksu = "TYP_INDEKSU";
                                        string AtrQuery5 = "SELECT  TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_Typ_Indeksu + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        string dea_Kod_ST_WNIP = "ST/WNIP";
                                        string AtrQuery6 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_ST_WNIP + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        string dea_Kod_KST = "KST";
                                        string AtrQuery7 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod_KST + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + aKody + "')";

                                        // string connectionString=Sesja.ConfigConnection.ConnectionString;
                                        //string wynikQueryAtr = ConnectToSQL(connectionString, AtrQuery);




                                        string wynik = "";

                                        var wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                        //MessageBox.Show(wynikQueryAtr.GetString(), "Komunikat",
                                        //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr2 = ZasPozycji.AddNew(null);
                                        Atr2.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_Gwarancja_MC + "'"];
                                        Atr2.Wartosc = wynik;




                                        wynikQueryAtr = Sesja.Connection.Execute(AtrQuery3, out dummy, 0);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr3 = ZasPozycji.AddNew(null);
                                        Atr3.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_Gwarancja_Rodzaj + "'"];
                                        Atr3.Wartosc = wynik;



                                        wynikQueryAtr = Sesja.Connection.Execute(AtrQuery4, out dummy, 0);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr4 = ZasPozycji.AddNew(null);
                                        Atr4.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_Stawka_Amorty + "'"];
                                        Atr4.Wartosc = wynik;



                                        wynikQueryAtr = Sesja.Connection.Execute(AtrQuery5, out dummy, 0);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr5 = ZasPozycji.AddNew(null);
                                        Atr5.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_Typ_Indeksu + "'"];
                                        Atr5.Wartosc = wynik;



                                        wynikQueryAtr = Sesja.Connection.Execute(AtrQuery7, out dummy, 0);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr7 = ZasPozycji.AddNew(null);
                                        Atr7.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_KST + "'"];
                                        Atr7.Wartosc = wynik;




                                        wynikQueryAtr = Sesja.Connection.Execute(AtrQuery6, out dummy, 0);
                                        if (wynikQueryAtr.EOF)
                                            wynik = "";
                                        else
                                            wynik = wynikQueryAtr.GetString().Trim();
                                        var Atr6 = ZasPozycji.AddNew(null);
                                        Atr6.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_ST_WNIP + "'"];
                                        Atr6.Wartosc = wynik;

                                        Sesja.Save();
                                        Sesja.Dispose();
                                    }
                                    catch (Exception ex)
                                    {

                                        MessageBox.Show("Błąd przy dodawaniu pozycji: " + Convert.ToString(ex.StackTrace) + Convert.ToString(ex), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }

                                }

                                ceny.Clear();

                            }

                        }
                        else
                        {
                            if (zly_kod != "")
                            {
                                MessageBox.Show("Towar " + zly_kod + " nie istnieje w bazie. Nie importowano dokumentu " + atrWartosc, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                           
                            else
                            {
                                MessageBox.Show("Dla " + atrWartosc + " istnieje już dokument", "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        kody.Clear();
                        k++;
                        Sesja.Save();
                        Sesja.Dispose();
                        
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Wywala: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd: " + e.Message + "\nNie dodano WZ dla akronimu: " + aAkronimy, "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            pBar1.Value = 1;
            pBar1.Visible = false;
            MessageBox.Show("Zakończono generowanie WZ.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        #endregion

        #region PW
        void dodawaniePW(DataTable dataTable, CDNBase.AdoSession Sesja)
        {
            List<string> akronimy = new List<string>();
            List<DateTime> data = new List<DateTime>();
            List<string> knt_kod = new List<string>();
            string akronim = "";
            for (int i = 2; i < dataTable.Rows.Count; i++)
            {
                /* if (String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][7])) == false)
                 MessageBox.Show(Convert.ToString(dataTable.Rows[i][5])+" Długość: "+Convert.ToString(Convert.ToString(dataTable.Rows[i][5]).Length), "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);*/

                if (String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[i][7])) == false && Convert.ToString(dataTable.Rows[i][5]).Substring(0, 1) == "P")
                    akronim = Convert.ToString(dataTable.Rows[i][9]);
                if (!akronimy.Contains(akronim) && akronim != "")
                {
                    akronimy.Add(akronim);
                    data.Add(Convert.ToDateTime(dataTable.Rows[i][8]));
                    knt_kod.Add("PDV" + Convert.ToString(dataTable.Rows[i][7]));
                }

            }

            int k = 0;
            pBar1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar1.Maximum = akronimy.Count + 1;
            // Set the initial value of the ProgressBar.
            pBar1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar1.Step = 1;

           
                foreach (string aAkronimy in akronimy)
                {
                    try
                    {
                        pBar1.PerformStep();
                        Sesja = Login.CreateSession();
                    //MessageBox.Show(aAkronimy, "Komunikat",
                    //MessageBoxButtons.OK, MessageBoxIcon.Error);
                    string atrWartosc = aAkronimy;
                    List<string> kody = new List<string>();
                        object dummy = Type.Missing;

                   
                        
                        var sqlQuery1 = "SELECT TOP(1)[DAt_WartoscTxt]  FROM " + Baza + "[DokAtrybuty] where [DAt_WartoscTxt] ='" + atrWartosc + "'";
                         var wynikQuery = Sesja.Connection.Execute(sqlQuery1, out dummy, 0);


                        string kod = "";
                        for (int i = 2; i < dataTable.Rows.Count; i++)
                        {
                            if (Convert.ToString(dataTable.Rows[i][9]) == aAkronimy)
                            {
                                kod = Convert.ToString(dataTable.Rows[i][0]);
                                if (!kody.Contains(kod))
                                {
                                    kody.Add(kod);
                                    //MessageBox.Show("Dodawanie kodu-> Kod: " + kod, "Komunikat",
                                    // MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                   

                        string zly_kod = "";
                        bool dobre_kody = true;
                        foreach (string aKody in kody)
                        {
                       
                            var sqlQuery2 = "SELECT TOP(1)[Twr_Kod] FROM " + Baza + "[Towary] where [Twr_Kod] ='" + aKody + "'";
                            var wynikQuery2 = Sesja.Connection.Execute(sqlQuery2, out dummy, 0);
                            //MessageBox.Show("Kod: " + aKody , "Komunikat",
                            //  MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (wynikQuery2.EOF)
                            {
                                dobre_kody = false;
                                zly_kod = aKody;
                                //MessageBox.Show("Czy zły kod? "+ Convert.ToString(zly_kod), "Komunikat",
                                //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }
                        
                        }
                        kody.Clear();

                        //tu będzie knt_kod[k] zamiast sKnt_Kod
                        var sqlQuery3 = "SELECT TOP(1)[Knt_Kod]  FROM " + Baza + "[Kontrahenci] where [Knt_Kod] ='" + knt_kod[k] + "'";
                        var wynikQuery3 = Sesja.Connection.Execute(sqlQuery3, out dummy, 0);


                        if (wynikQuery.EOF && dobre_kody == true && !wynikQuery3.EOF)
                        {
                        
                            CDNHlmn.DokumentyHaMag PeZety = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
                            var PeZet = (CDNHlmn.IDokumentHaMag)PeZety.AddNew(null);
                             var zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));

                            //sKnt_Kod=knt_kod[k] -> odkomentować w głównym programie
                            var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
                            var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='" + knt_kod[k] + "'"];   //Tutaj kod ostrawki

                            //dodac typ PW
                            PeZet.Rodzaj = 303000;
                            PeZet.TypDokumentu = 303;
                            PeZet.Bufor = 1;
                            PeZet.Uwagi = aAkronimy;



                            // PeZet.NumerObcy = "";

                            //Ustawiamy date
                            PeZet.DataDok = data[k];
                            PeZet.DataWys = data[k];
                            PeZet.Podmiot = (CDNHeal.IPodmiot)kontrahent;

                            //atrybut dokumentu
                            string dea_kod_nr_dok = "NR DOKUMENTU FORCOM";
                            var AtrDokument = PeZet.Atrybuty;
                            var rAtrybut = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);
                            var Atr = AtrDokument.AddNew(null);
                            Atr.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybut["dea_Kod = '" + dea_kod_nr_dok + "'"];
                            Atr.Wartosc = atrWartosc;



                            /*List<string> kody = new List<string>();
                            string kod = "";*/
                            Sesja.Save();
                       
                            var id_PZ = PeZet.ID;



                        for (int a = 2; a < dataTable.Rows.Count; a++)
                            {
                                if (Convert.ToString(dataTable.Rows[a][9]) == aAkronimy)
                                {
                                    try
                                    {
                                        for (int ilosc = 0; ilosc < Convert.ToDouble(dataTable.Rows[a][2]); ilosc++)
                                        {
                                            Sesja = Login.CreateSession();
                                            PeZet = Sesja.CreateObject("CDN.DokumentyHaMag").Item("TrN_TrNId=" + id_PZ);
                                            CDNBase.ICollection pozycje = PeZet.Elementy;
                                            var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);

                                            //przypisywanie wartości pozycji
                                            CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_Kod='" + Convert.ToString(dataTable.Rows[a][0]) + "'"];
                                            //MessageBox.Show(zTowar.Kod, "Komunikat",
                                            //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            pozycja.TowarKod = zTowar.Kod;
                                            pozycja.Ilosc = 1.0;
                                            pozycja.Cena0 = 0;
                                            pozycja.Cena0WD = 0;
                                            pozycja.WartoscNetto = 0;
                                            pozycja.WartoscBrutto = 0;

                                            //var zasob = pozycja.Zasoby;



                                            var rDostawy = pozycja.Dostawy;
                                            string dea_Kod_NR_SERYJNY = "NUMER SERYJNY";

                                            var rDostawa = rDostawy.AddNew();
                                            pozycja.UstawDostawy();
                                            rDostawa.Cecha1_Kod = dea_Kod_NR_SERYJNY;
                                            rDostawa.Cecha1_Wartosc = Convert.ToString(dataTable.Rows[a][3]);
                                            rDostawa.Ilosc = 1.0;
                                            // MessageBox.Show(Convert.ToString(dataTable.Rows.Count), "Komunikat",
                                            // MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            if (dataTable.Columns.Count == 12 && String.IsNullOrEmpty(Convert.ToString(dataTable.Rows[a][11])) == false && Convert.ToString(dataTable.Rows[a][11]) == "T")
                                            {
                                                rDostawa.Cecha2_Kod = "LEASING";
                                                rDostawa.Cecha2_Wartosc = "Tak";
                                                //MessageBox.Show("Dodawanie Leasingu", "Komunikat",
                                                //MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }

                                            pozycja.UsunDostawy();


                                            // pozycja.UsunDostawy();


                                            Sesja.Save();
                                            var ZasPozycji = pozycja.Atrybuty;
                                            var rAtrybuty = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);

                                            //usuwanie istniejących atrybutów
                                            string TrEId = "SELECT TOP(1) [TrA_TrEId] FROM " + Baza + "[TraElemAtr] WHERE [TrA_TrEId] =" + pozycja.ID;
                                            var wynikExistQuery = Sesja.Connection.Execute(TrEId, out dummy, 0);
                                            while (!wynikExistQuery.EOF)
                                            {
                                                ZasPozycji.Delete(0);
                                                Sesja.Save();
                                                wynikExistQuery = Sesja.Connection.Execute(TrEId, out dummy, 0);
                                            }

                                            
                                            string wynik = "";

                                            string dea_Kod = "GWARACJA MC";
                                            string AtrQuery2 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            var wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr2 = ZasPozycji.AddNew(null);
                                            Atr2.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr2.Wartosc = wynik;


                                            dea_Kod = "GWARANCJA RODZAJ";
                                            AtrQuery2 = "SELECT TOP(1)  [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr3 = ZasPozycji.AddNew(null);
                                            Atr3.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr3.Wartosc = wynik;


                                            dea_Kod = "STAWKA AMORTYZACJI";
                                            AtrQuery2 = "SELECT TOP(1)  [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr4 = ZasPozycji.AddNew(null);
                                            Atr4.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr4.Wartosc = wynik;


                                            dea_Kod = "TYP_INDEKSU";
                                            AtrQuery2 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr5 = ZasPozycji.AddNew(null);
                                            Atr5.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr5.Wartosc = wynik;


                                            dea_Kod = "KST";
                                            AtrQuery2 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr7 = ZasPozycji.AddNew(null);
                                            Atr7.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr7.Wartosc = wynik;



                                            dea_Kod = "ST/WNIP";
                                            AtrQuery2 = "SELECT TOP(1) [TwA_WartoscTxt] FROM " + Baza + "[TwrAtrybuty] WHERE [TwA_DeAId] IN (SELECT [DeA_DeAId]  FROM " + Baza + "[DefAtrybuty] WHERE [DeA_Kod] = '" + dea_Kod + "') and [TwA_TwrId] IN (SELECT [Twr_TwrId]  FROM " + Baza + "[Towary] WHERE [Twr_Kod] = '" + Convert.ToString(dataTable.Rows[a][0]) + "')";
                                            wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                            if (wynikQueryAtr.EOF)
                                                wynik = "";
                                            else
                                                wynik = wynikQueryAtr.GetString().Trim();
                                            var Atr6 = ZasPozycji.AddNew(null);
                                            Atr6.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                            Atr6.Wartosc = wynik;

                                            dea_Kod = "NR ST/URZĄDZ";
                                            string countQuery = "SELECT count(*) FROM " + Baza + "[TrwaleAtrybuty] WHERE [SrA_DeAId] = 25 and [SrA_WartoscTxt]='" + Convert.ToString(dataTable.Rows[a][3]) + "' GROUP BY [SrA_WartoscTxt] HAVING count(*)>1";
                                            var wynikCountQuery = Sesja.Connection.Execute(countQuery, out dummy, 0);
                                            if (wynikCountQuery.EOF)
                                            {
                                                AtrQuery2 = "SELECT TOP(1) [SrT_NrInwent] FROM " + Baza + "[Trwale] Where SrT_SrTID IN (SELECT [SrA_SrTId] FROM " + Baza + "[TrwaleAtrybuty] WHERE [SrA_DeAId] = 25 and [SrA_WartoscTxt]='" + Convert.ToString(dataTable.Rows[a][3]) + "')";
                                                wynikQueryAtr = Sesja.Connection.Execute(AtrQuery2, out dummy, 0);
                                                if (wynikQueryAtr.EOF)
                                                    wynik = "";
                                                else
                                                    wynik = wynikQueryAtr.GetString().Trim();
                                                var Atr8 = ZasPozycji.AddNew(null);
                                                Atr8.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod + "'"];
                                                Atr8.Wartosc = wynik;
                                            }
                                            else
                                                MessageBox.Show("Istnieją " + wynikCountQuery.GetString() + "nazwy środka trwałego dla numeru seryjnego: " + Convert.ToString(dataTable.Rows[a][3]) + "\nNie dodano atrybutu 'NR ST/URZĄDZ' dla pozycji z towarem " + Convert.ToString(dataTable.Rows[a][0]) + " w dokumencie " + PeZet.NumerPelny, "Komunikat",
                                                   MessageBoxButtons.OK, MessageBoxIcon.Error);

                                            Sesja.Save();
                                            Sesja.Dispose();


                                        }
                                    

                                    }
                                    catch (Exception se)
                                    {
                                        
                                        MessageBox.Show("Błąd przy dodawaniu pozycji: " + Convert.ToString(se), "Komunikat",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);

                                    }
                                

                            }

                            }

                        }
                        else
                        {
                            if (zly_kod != "")
                            {
                                MessageBox.Show("Towar " + zly_kod + " nie istnieje w bazie. Nie importowano dokumentu " + atrWartosc, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            else if (wynikQuery3.EOF)
                            {
                                //tutaj będzie knt_kod[k] zamiast sKnt_Kod
                                MessageBox.Show("Kontrahent " + knt_kod[k] + " nie istnieje w bazie. Nie importowano dokumentu " + atrWartosc, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Dla " + atrWartosc + " istnieje już dokument PW", "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }

                        k++;
                        Sesja.Save();
                        Sesja.Dispose();
                       


                    }

                    catch (Exception e)
                    {
                      
                        MessageBox.Show("Błąd: " + e.Message + "\nNie dodano PW dla akronimu: " + aAkronimy, "Komunikat",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    }
                }
            
         
        
            pBar1.Value = 1;
            pBar1.Visible = false;
            MessageBox.Show("Zakończono generowanie PW.", "Komunikat",
               MessageBoxButtons.OK, MessageBoxIcon.Information);

            #endregion

           
        }


        #region SQL
        string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        #endregion
    }
}
