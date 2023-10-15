using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace smcPaketversand
{
    public class DokumentInfo
    {
        public string AdressNr;
        public int Kundennummer;
        public bool SendungErstellt = false;
        public string Benutzername = "";
        public string Passwort = "";
        public string DocNr = "";
        public string AppUser = "";
        public int numPackages = 0;
        public string Reference = "";
        public string Currency = "";
        public int MultiLieferscheine = 0;
        public string UnserZ = "";
        public DateTime TempDate;
        public string DeliveryDate = "";
        public string DHLDeliveryDate = "";
        public decimal Gewicht = 0;
        public string Anrede = "";
        public string Name = "";
        public string Vorname = "";
        public string Strasse = "";
        public int PLZ = 0;
        public string Ort = "";
        public string Kanton = "";
        public string Land = "";
        public string LiefAnrede = "";
        public string LiefName = "";
        public string LiefVorname = "";
        public string LiefStrasse = "";
        public int LiefPLZ = 0;
        public string LiefOrt = "";
        public string LiefKanton = "";
        public string LiefLand = "";
        public string LiefTel = "";
        public string Lieferart = "";
        public string Serviceart = "";
        public string Inhaltsbeschreibung = "";
        public string TrackingNr = "";
        public string MailAbsender = "";
        public string MailPasswort = "";
        public string SMTP = "";
        public string MailPort = "";
        public string MailD = "";
        public string MailF = "";
        public string MailI = "";
        public string MailE = "";
        public List<string> AlleLieferscheine;
        public string LiefEMail = "";
        public string Tel = "";
        public string Kontakt;
        public string LiefKontakt;
        public string Adresszeile1;
        public string Adresszeile2;
        public string LiefAdresszeile1;
        public string LiefAdresszeile2;

        public DokumentInfo(string xinipath)
        {
            this.numPackages = 1;
            if (!string.IsNullOrEmpty(xinipath))
            {
                string[] lines = File.ReadAllLines(xinipath, Encoding.Default);
                foreach (string line in lines)
                {
                    if (line.StartsWith("dfnAdressNr="))
                    {
                        AdressNr = line.Substring(line.IndexOf("=") + 1);
                    }
                    if (line.StartsWith("LoginUser="))
                    {
                        AppUser = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("dfnDokNrAUF="))
                    {
                        DocNr = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("dfsReferenz="))
                    {
                        Reference = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Adresszeile1="))
                    {
                        Adresszeile1 = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Adresszeile2="))
                    {
                       Adresszeile2 = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefAdresszeile1="))
                    {
                        LiefAdresszeile1 = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefAdresszeile2="))
                    {
                        LiefAdresszeile2 = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("dfsWaehrung="))
                    {
                        Currency = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("dfdStartSDatum="))
                    {
                        if (DateTime.TryParse(line.Substring(line.IndexOf("=") + 1), out TempDate) && string.IsNullOrEmpty(DeliveryDate))
                        {
                            DeliveryDate = TempDate.AddDays(1).ToString("yyyy-MM-ddT16:00:00");
                            DHLDeliveryDate = TempDate.AddDays(1).ToString("yyyy-MM-ddT16:00") + "GMT+01:00";
                        }
                    }
                    else if (line.StartsWith("dfdLiefertermin="))
                    {
                        if (DateTime.TryParse(line.Substring(line.IndexOf("=") + 1), out TempDate))
                        {
                            DeliveryDate = TempDate.ToString("yyyy-MM-ddT16:00:00");
                            DHLDeliveryDate = TempDate.ToString("yyyy-MM-ddT16:00") + "GMT+01:00";
                        }
                        else if (string.IsNullOrEmpty(DeliveryDate))
                        {
                            DeliveryDate = "";
                        }
                    }
                    else if (line.StartsWith("dfnTotGew="))
                    {
                        Decimal.TryParse(line.Substring(line.IndexOf("=") + 1), out Gewicht);
                    }
                    else if (line.StartsWith("Anrede="))
                    {
                        Anrede = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Name="))
                    {
                        Name = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Vorname="))
                    {
                        Vorname = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Kontakt="))
                    {
                        Kontakt = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Strasse="))
                    {
                        Strasse = line.Substring(line.IndexOf("=") + 1);
                        //Strasse = ));
                    }
                    else if (line.StartsWith("PLZ="))
                    {
                        Int32.TryParse(line.Substring(line.IndexOf("=") + 1), out PLZ);
                    }
                    else if (line.StartsWith("Ort="))
                    {
                        Ort = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Kanton="))
                    {
                        Kanton = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Land="))
                    {
                        Land = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefName="))
                    {
                        LiefName = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefVorname="))
                    {
                        LiefVorname = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefStrasse="))
                    {
                        LiefStrasse = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefKontakt="))
                    {
                        LiefKontakt = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefPLZ="))
                    {
                        Int32.TryParse(line.Substring(line.IndexOf("=") + 1), out LiefPLZ);
                    }
                    else if (line.StartsWith("LiefOrt="))
                    {
                        LiefOrt = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefKanton="))
                    {
                        LiefKanton = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefLand="))
                    {
                        LiefLand = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("LiefTel="))
                    {
                        LiefTel = line.Substring(line.IndexOf("=") + 1);
                        Tel = LiefTel;
                    }
                    else if (line.StartsWith("LiefEMail="))
                    {
                        LiefEMail = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("dfsLieferart="))
                    {
                        Lieferart = line.Substring(line.IndexOf("=") + 1);
                    }
                    else if (line.StartsWith("Z_Paketart="))
                    {
                        Serviceart = line.Substring(line.IndexOf("=") + 1);
                    }
                }
            }
        }
    }
}
