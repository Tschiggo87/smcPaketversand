using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SmcBasics;

namespace smcPaketversand
{
    public partial class frmPaketVersand : Form
    {
        decimal[,] d_packages;
        string[] s_packagecomment;
        int[] i_packagetype;
        int selectedPackage = -1;
        DokumentInfo dok;
        SmcConfig ini;
        SmcErrors err;
        SmcLog log;

        public frmPaketVersand(string dokumentInfoPfad, string docnr = "")
        {
            try
            {
                err = new SmcErrors();
                ini = new SmcConfig(AppDomain.CurrentDomain.BaseDirectory + "\\config\\smcPaketversand.ini", "52fafc33-1b82-4612-9234-6dcc107f3152");
                log = new SmcLog(SmcLog.LogLevel.Debug, "smcPaketversand", err: err);
                ini.Environment = "Prod";
                InitializeComponent();
                dateLiefertermin.MinDate = DateTime.Today.AddDays(1);
                d_packages = new decimal[100, 4];
                i_packagetype = new int[100];
                s_packagecomment = new string[100];
                SetPakettypen();
                HandleTabs();
                dok = new DokumentInfo(dokumentInfoPfad);
                if (dokumentInfoPfad.Equals(""))
                {
                    if (!docnr.Equals(""))
                    {
                        LoadDataFromDocNr(docnr);
                        SetDokument(docnr, lbAdressNr.Text);
                    }
                }
                else
                {

                    SetAdressen();
                    SetDokument(dok.DocNr, dok.AdressNr);


                    for (int i = 0; i < s_packagecomment.Length; i++)
                    {
                        s_packagecomment[i] = dok.Reference;
                    }
                }
                //for (int i = 0; i < 100; i++)
                //{
                //    d_packages[i, 0] = 1;
                //    d_packages[i, 1] = 1;
                //    d_packages[i, 2] = 1;
                //    d_packages[i, 3] = 1;
                //}

                cbPostPakete.SelectedIndex = 0;
                cbPostPakettyp.SelectedIndex = 0;
                cbDPDPakete.SelectedIndex = 0;
                cbDPDPakettyp.SelectedIndex = 0;
                cbDHLPakete.SelectedIndex = 0;
                cbDHLPakettyp.SelectedIndex = 0;
                cbFedExPakete.SelectedIndex = 0;
                cbFedExPakettyp.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void SetPakettypen()
        {
            try
            {
                DataTable dt = SmcDB.StoredProcedureDataTable("GetPakettypen", new Dictionary<string, object>(), "ProffixDB", ini);
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        cbPostPakettyp.Items.Add(new ComboBoxItem(dt.Rows[i]["Typname"].ToString(), Decimal.Parse(dt.Rows[i]["HoeheInCM"].ToString()), Decimal.Parse(dt.Rows[i]["LaengeInCM"].ToString()), Decimal.Parse(dt.Rows[i]["BreiteInCM"].ToString())));
                        cbDHLPakettyp.Items.Add(new ComboBoxItem(dt.Rows[i]["Typname"].ToString(), Decimal.Parse(dt.Rows[i]["HoeheInCM"].ToString()), Decimal.Parse(dt.Rows[i]["LaengeInCM"].ToString()), Decimal.Parse(dt.Rows[i]["BreiteInCM"].ToString())));
                        cbDPDPakettyp.Items.Add(new ComboBoxItem(dt.Rows[i]["Typname"].ToString(), Decimal.Parse(dt.Rows[i]["HoeheInCM"].ToString()), Decimal.Parse(dt.Rows[i]["LaengeInCM"].ToString()), Decimal.Parse(dt.Rows[i]["BreiteInCM"].ToString())));
                        cbFedExPakettyp.Items.Add(new ComboBoxItem(dt.Rows[i]["Typname"].ToString(), Decimal.Parse(dt.Rows[i]["HoeheInCM"].ToString()), Decimal.Parse(dt.Rows[i]["LaengeInCM"].ToString()), Decimal.Parse(dt.Rows[i]["BreiteInCM"].ToString())));
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void SetDokument(string docnr, string adressnr)
        {
            try
            {
                DataTable dt;
                if (!String.IsNullOrEmpty(docnr))
                {
                    lbAdressNr.Text = adressnr;
                    txtDokNr.Text = docnr;
                    dt = SmcDB.StoredProcedureDataTable("GetOpenLS", new Dictionary<string, object>() { { "@AdressNr", adressnr } }, "ProffixDB", ini);
                    clbPostLieferscheine.Items.Clear();
                    clbDHLLieferscheine.Items.Clear();
                    clbDPDLieferscheine.Items.Clear();
                    clbFedExLieferscheine.Items.Clear();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string item = dt.Rows[i]["DokumentNrAUF"].ToString();
                        clbPostLieferscheine.Items.Add(item);
                        clbDHLLieferscheine.Items.Add(item);
                        clbDPDLieferscheine.Items.Add(item);
                        clbFedExLieferscheine.Items.Add(item);
                        if (item.Equals(docnr))
                        {
                            clbPostLieferscheine.SetItemChecked(clbPostLieferscheine.Items.IndexOf(item), true);
                            clbDHLLieferscheine.SetItemChecked(clbDHLLieferscheine.Items.IndexOf(item), true);
                            clbDPDLieferscheine.SetItemChecked(clbDPDLieferscheine.Items.IndexOf(item), true);
                            clbFedExLieferscheine.SetItemChecked(clbFedExLieferscheine.Items.IndexOf(item), true);
                        }
                    }
                }
                if (!String.IsNullOrEmpty(dok.DeliveryDate) && Convert.ToDateTime(dok.DeliveryDate) >= DateTime.Today.AddDays(1))
                {
                    dateLiefertermin.Value = Convert.ToDateTime(dok.DeliveryDate);
                }
                dt = SmcDB.StoredProcedureDataTable("GetLieferarten", new Dictionary<string, object>(), "ProffixDB", ini);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cbLieferart.Items.Add(dt.Rows[i]["Bezeichnung"].ToString());
                }
                if (!String.IsNullOrEmpty(dok.Lieferart))
                {
                    int dokLieferartId = cbLieferart.FindString(dok.Lieferart);
                    if (dokLieferartId >= 0)
                    {
                        cbLieferart.SelectedIndex = dokLieferartId;
                    }
                    else if (cbLieferart.Items.Count > 0)
                    {
                        cbLieferart.SelectedIndex = 0;
                        dok.Lieferart = cbLieferart.SelectedItem.ToString();
                    }
                }
                else if (cbLieferart.Items.Count > 0)
                {
                    cbLieferart.SelectedIndex = 0;
                    dok.Lieferart = cbLieferart.SelectedItem.ToString();
                }
                SetAbholarten(dok.Lieferart);
            }
            catch (Exception ex) { }
        }

        public void SetAbholarten(string lieferart)
        {
            try
            {
                cbAbholart.Items.Clear();
                DataTable dt = SmcDB.StoredProcedureDataTable("GetAbholarten", new Dictionary<string, object>() { { "lieferart", lieferart } }, "ProffixDB", ini);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cbAbholart.Items.Add(dt.Rows[i]["Bezeichnung"].ToString());
                }
                if (cbAbholart.Items.Count > 0 && cbAbholart.SelectedIndex < 0)
                {
                    cbAbholart.SelectedIndex = 0;
                }
            }
            catch (Exception ex) { }
        }

        public void HandleTabs()
        {
            try
            {
                List<string> spediteure = ini.GetValue(ini.Environment + "_Application", "Spediteure").Split(',').ToList<string>();
                if (!spediteure.Contains("Post"))
                {
                    tabControl1.TabPages.Remove(tabPost);
                }
                if (!spediteure.Contains("DHL"))
                {
                    tabControl1.TabPages.Remove(tabDHL);
                }
                if (!spediteure.Contains("DPD"))
                {
                    tabControl1.TabPages.Remove(tabDPD);
                }
                if (!spediteure.Contains("FedEx"))
                {
                    tabControl1.TabPages.Remove(tabFedEx);
                }
            }
            catch (Exception ex) { }
        }

        public void SetAdressen()
        {
            try
            {
                DataTable datata = SmcDB.StoredProcedureDataTable("GetTel", new Dictionary<string, object>() { { "doknr", dok.DocNr } }, "ProffixDB", ini);
                lbETelefon.Text = datata.Rows[0]["tel"].ToString();
                dok.LiefTel = datata.Rows[0]["tel"].ToString();
                dok.LiefEMail = datata.Rows[0]["email"].ToString();
                if (String.IsNullOrEmpty(dok.LiefName) || String.IsNullOrEmpty(dok.LiefStrasse) || String.IsNullOrEmpty(dok.LiefOrt))
                {
                    lbEFirma.Text = dok.Name;
                    lbEStrasse.Text = dok.Strasse;
                    lbEPLZ.Text = dok.PLZ > 0 ? dok.PLZ.ToString() : "";
                    lbEOrt.Text = dok.Ort;
                    lbLand.Text = dok.Land;
                    lbEName.Text = dok.Kontakt;
                    lbEEmail.Text = dok.LiefEMail;
                }
                else
                {
                    lbEFirma.Text = dok.LiefName;
                    lbEStrasse.Text = dok.LiefStrasse;
                    lbEPLZ.Text = dok.LiefPLZ > 0 ? dok.LiefPLZ.ToString() : "";
                    lbEOrt.Text = dok.LiefOrt;
                    lbLand.Text = dok.LiefLand;
                    lbETelefon.Text = dok.LiefTel;
                    lbEName.Text = dok.LiefKontakt;
                    lbEEmail.Text = dok.LiefEMail;
                }
                if (lbETelefon.Text == "0")
                {
                    lbETelefon.Text = "";
                }
                //MessageBox.Show("Server=" + ini.GetValue(ini.Environment + "_ProffixDB", "DataSource") + ";Database=" + ini.GetValue(ini.Environment + "_ProffixDB", "Catalog") + ";User Id=" + ini.GetValue(ini.Environment + "_ProffixDB", "Username") + ";Password=" + ini.GetValue(ini.Environment + "_ProffixDB", "Password") + ";");
                DataTable dt = SmcDB.StoredProcedureDataTable("GetStammdaten", new Dictionary<string, object>() { { "username", dok.AppUser } }, "ProffixDB", ini);
                lbName.Text = "Lager";
                lbFirma.Text = dt.Rows[0]["Firma"].ToString();
                lbStrasse.Text = dt.Rows[0]["Strasse"].ToString();
                lbPLZ.Text = dt.Rows[0]["PLZ"].ToString();
                lbOrt.Text = dt.Rows[0]["Ort"].ToString();
                lbTelefon.Text = dt.Rows[0]["Telefon"].ToString();
                lbEmail.Text = dt.Rows[0]["EMail"].ToString();
            }
            catch (Exception ex) { }
        }

        private void FedExSend_Click(object sender, EventArgs e)
        {
            SavePackage(txtFedExGewicht.Text, txtFedExLaenge.Text, txtFedExBreite.Text, txtFedExHoehe.Text, txtFedExKundenref.Text, cbFedExPakettyp.SelectedIndex);
            try
            {
                for (int i = 0; i < dok.numPackages; i++)
                {
                    if (d_packages[i, 0] <= 0 || d_packages[i, 1] <= 0 || d_packages[i, 2] <= 0 || d_packages[i, 3] <= 0)
                    {
                        MessageBox.Show("Paket " + (i + 1) + " enthält ungültige Werte!");
                        return;
                    }
                }
                MessageBox.Show(dateLiefertermin.Value.ToString("yyyy-MM-ddTHH:mm") + " GMT+01:00");
                XNamespace version = "http://fedex.com/ws/ship/v26";
                XNamespace soapenv = "http://schemas.xmlsoap.org/soap/envelope/";
                for (int i = 0; i < dok.numPackages; i++)
                {
                    XDocument xml = new XDocument(
                        new XElement(soapenv + "Envelope",
                            new XAttribute(XNamespace.Xmlns + "soapenv", "http://schemas.xmlsoap.org/soap/envelope/"),
                            new XAttribute("xmlns", "http://fedex.com/ws/ship/v26"),
                            new XElement(soapenv + "Body",
                                new XElement(version + "ProcessShipmentRequest",
                                    new XElement(version + "WebAuthenticationDetail",
                                        new XElement(version + "UserCredential",
                                            new XElement(version + "Key", ini.GetValue(ini.Environment + "_FedEx", "Key")),
                                            new XElement(version + "Password", ini.GetValue(ini.Environment + "_FedEx", "Password"))
                                        )
                                    ),
                                    new XElement(version + "ClientDetail",
                                        new XElement(version + "AccountNumber", ini.GetValue(ini.Environment + "_FedEx", "AccountNumber")),
                                        new XElement(version + "MeterNumber", ini.GetValue(ini.Environment + "_FedEx", "MeterNumber"))
                                    ),
                                    new XElement(version + "Version",
                                        new XElement(version + "ServiceId", "ship"),
                                        new XElement(version + "Major", 26),
                                        new XElement(version + "Intermediate", 0),
                                        new XElement(version + "Minor", 0)
                                    ),
                                    new XElement(version + "RequestedShipment",
                                        new XElement(version + "ShipTimestamp", dateLiefertermin.Value.ToString("yyyy-MM-ddTHH:mm") + " GMT+01:00"),
                                        new XElement(version + "DropoffType", cbAbholart.Text),
                                        new XElement(version + "ServiceType", cbLieferart.Text),
                                        new XElement(version + "PackagingType", "YOUR_PACKAGING"),
                                        new XElement(version + "Shipper",
                                            new XElement(version + "Contact",
                                                new XElement(version + "PersonName", lbName.Text),
                                                new XElement(version + "CompanyName", lbFirma.Text),
                                                new XElement(version + "PhoneNumber", lbTelefon.Text),
                                                new XElement(version + "EMailAddress", lbEmail.Text)
                                            ),
                                            new XElement(version + "Address",
                                                new XElement(version + "StreetLines", lbStrasse.Text),
                                                new XElement(version + "City", lbOrt.Text),
                                                new XElement(version + "PostalCode", lbPLZ.Text),
                                                new XElement(version + "CountryCode", "CH")
                                            )
                                        ),
                                        new XElement(version + "Recipient",
                                            new XElement(version + "Contact",
                                                new XElement(version + "PersonName", lbEName.Text),
                                                new XElement(version + "CompanyName", lbEFirma.Text),
                                                new XElement(version + "PhoneNumber", lbETelefon.Text),
                                                new XElement(version + "EMailAddress", lbEEmail.Text)
                                            ),
                                            new XElement(version + "Address",
                                                new XElement(version + "StreetLines", lbEStrasse.Text),
                                                new XElement(version + "City", lbEOrt.Text),
                                                new XElement(version + "PostalCode", lbEPLZ.Text),
                                                new XElement(version + "CountryCode", "CH")
                                            )
                                        ),
                                        new XElement(version + "ShippingChargesPayment",
                                            new XElement(version + "PaymentType", "SENDER"),
                                            new XElement(version + "Payor",
                                                new XElement(version + "ResponsibleParty",
                                                    new XElement(version + "AccountNumber", ini.GetValue(ini.Environment + "_FedEx", "AccountNumber"))
                                                )
                                            )
                                        ),
                                        new XElement(version + "LabelSpecification",
                                            new XElement(version + "LabelFormatType", "COMMON2D"),
                                            new XElement(version + "ImageType", "PNG"),
                                            new XElement(version + "LabelStockType", "PAPER_7X4.75")
                                        ),
                                        new XElement(version + "PackageCount", 1),
                                        new XElement(version + "RequestedPackageLineItems",
                                            new XElement(version + "SequenceNumber", 1),
                                            new XElement(version + "Weight",
                                                new XElement(version + "Units", "KG"),
                                                new XElement(version + "Value", txtFedExGewicht.Text)
                                            ),
                                            new XElement(version + "Dimensions",
                                                new XElement(version + "Length", txtFedExLaenge.Text),
                                                new XElement(version + "Width", txtFedExBreite.Text),
                                                new XElement(version + "Height", txtFedExHoehe.Text),
                                                new XElement(version + "Units", "CM")
                                            ),
                                            new XElement(version + "CustomerReferences",
                                                new XElement(version + "CustomerReferenceType", "CUSTOMER_REFERENCE"),
                                                new XElement(version + "Value", txtFedExKundenref.Text)
                                            )
                                        )
                                    )
                                )
                            )
                        )
                    );

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"https://wsbeta.fedex.com:443/web-services/ship");
                    request.Headers.Add(@"SOAP:Action");
                    request.ContentType = "text/xml;charset=\"utf-8\"";
                    request.Accept = "text/xml";
                    request.Method = "POST";
                    using (Stream str = request.GetRequestStream())
                    {
                        xml.Save(str);
                    }


                    using (WebResponse response = request.GetResponse())
                    {
                        using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                        {
                            string soapResult = rd.ReadToEnd();
                            xml = XDocument.Parse(soapResult);
                            //MessageBox.Show(xml.ToString());
                        }
                    }
                    string trackingnumber = xml.Descendants("TrackingNumber").First().Value.ToString();
                    byte[] bytes = Convert.FromBase64String(xml.Descendants("image").First().Value.ToString());
                    Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Etiketten"));
                    var stream = new FileStream(Path.Combine(Environment.CurrentDirectory, "Etiketten", "FedEx_Etikette_" + trackingnumber + ".pdf"), FileMode.CreateNew);
                    var writer = new BinaryWriter(stream);
                    writer.Write(bytes, 0, bytes.Length);
                    writer.Close();
                    stream.Close();
                    Process.Start(Path.Combine(Environment.CurrentDirectory, "Etiketten", "FedEx_Etikette_" + trackingnumber + ".pdf"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PostSend_Click(object sender, EventArgs e)
        {
            SavePackage(txtPostGewicht.Text, txtPostLaenge.Text, txtPostBreite.Text, txtPostHoehe.Text, txtPostKundenref.Text, cbPostPakettyp.SelectedIndex);
            try
            {
                for (int i = 0; i < dok.numPackages; i++)
                {
                    if (d_packages[i, 0] <= 0 || d_packages[i, 1] <= 0 || d_packages[i, 2] <= 0 || d_packages[i, 3] <= 0)
                    {
                        MessageBox.Show("Paket " + (i + 1) + " enthält ungültige Werte!");
                        return;
                    }
                }
                XNamespace typ = "https://wsbc.post.ch/wsbc/barcode/v2_4/types";
                XNamespace soapenv = "http://schemas.xmlsoap.org/soap/envelope/";
                for (int i = 0; i < dok.numPackages; i++)
                {
                    XDocument xml = new XDocument(
                        new XElement(soapenv + "Envelope",
                            new XAttribute(XNamespace.Xmlns + "soapenv", soapenv),
                            new XAttribute(XNamespace.Xmlns + "typ", typ),
                            new XElement(soapenv + "Body",
                                new XElement(typ + "GenerateLabel",
                                    new XElement(typ + "Language", "de"),
                                    new XElement(typ + "Envelope",
                                        new XElement(typ + "LabelDefinition",
                                            new XElement(typ + "LabelLayout", "A6"),
                                            new XElement(typ + "PrintAddresses", "RecipientAndCustomer"),
                                            new XElement(typ + "ImageFileType", "png"),
                                            new XElement(typ + "ImageResolution", 300),
                                            new XElement(typ + "PrintPreview", false)
                                        ),
                                        new XElement(typ + "FileInfos",
                                            new XElement(typ + "FrankingLicense", ini.GetValue(ini.Environment + "_Post", "Frankierlizenz")),
                                            new XElement(typ + "PpFranking", false),
                                            new XElement(typ + "Customer",
                                                new XElement(typ + "Name1", lbName.Text),
                                                new XElement(typ + "Street", lbStrasse.Text),
                                                new XElement(typ + "ZIP", lbPLZ.Text),
                                                new XElement(typ + "City", lbOrt.Text)
                                            )
                                        ),
                                        new XElement(typ + "Data",
                                            new XElement(typ + "Provider",
                                                new XElement(typ + "Sending",
                                                    new XElement(typ + "Item",
                                                        new XElement(typ + "Recipient",
                                                            new XElement(typ + "Name1", lbEName.Text),
                                                            new XElement(typ + "Street", lbEStrasse.Text),
                                                            new XElement(typ + "ZIP", lbEPLZ.Text),
                                                            new XElement(typ + "City", lbEOrt.Text)
                                                        ),
                                                        new XElement(typ + "Attributes",
                                                            new XElement(typ + "PRZL", "PRI"),
                                                            new XElement(typ + "PRZL", "SI"),
                                                            new XElement(typ + "Dimensions",
                                                                new XElement(typ + "Weight", txtPostGewicht.Text)
                                                            )
                                                        )
                                                    )
                                                )
                                            )
                                        )
                                    )
                                )
                            )
                        )
                    );

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ini.GetValue(ini.Environment + "_Post", "URL"));
                    string encoded = Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(ini.GetValue(ini.Environment + "_Post", "Benutzername") + ":" + ini.GetValue(ini.Environment + "_Post", "Passwort")));
                    request.Headers.Add("Authorization", "Basic " + encoded);
                    request.ContentType = "text/xml;charset=\"utf-8\"";
                    request.Accept = "text/xml";
                    request.Method = "POST";
                    using (Stream str = request.GetRequestStream())
                    {
                        xml.Save(str);
                    }


                    using (WebResponse response = request.GetResponse())
                    {
                        using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                        {
                            string soapResult = rd.ReadToEnd();
                            xml = XDocument.Parse(soapResult);
                        }
                    }
                    XNamespace ns = "https://wsbc.post.ch/wsbc/barcode/v2_4/types";
                    string trackingnumber = xml.Descendants(ns + "Data").First().Descendants(ns + "IdentCode").First().Value.ToString();
                    //MessageBox.Show(trackingnumber);
                    string image = xml.Descendants(ns + "Data").First().Descendants(ns + "Label").First().Value.ToString();
                    //MessageBox.Show(image);

                    byte[] bytes = Convert.FromBase64String(image);

                    Image img;
                    using (MemoryStream ms = new MemoryStream(bytes))
                    {
                        img = Image.FromStream(ms);
                    }

                    Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Etiketten"));
                    //var stream = new FileStream(Path.Combine(Environment.CurrentDirectory, "Etiketten", "Post_Etikette_" + trackingnumber + ".pdf"), FileMode.CreateNew);
                    //var writer = new BinaryWriter(stream);
                    //writer.Write(bytes, 0, bytes.Length);
                    //writer.Close();
                    //stream.Close();
                    img.Save(Path.Combine(Environment.CurrentDirectory, "Etiketten", "Post_Etikette_" + trackingnumber + ".png"), System.Drawing.Imaging.ImageFormat.Png);
                    Process.Start(Path.Combine(Environment.CurrentDirectory, "Etiketten", "Post_Etikette_" + trackingnumber + ".png"));
                    List<string> allelieferscheine = new List<string>();
                    foreach (var item in clbPostLieferscheine.CheckedItems)
                    {
                        allelieferscheine.Add(item.ToString());
                    }
                    foreach (string doknr in allelieferscheine)
                    {
                        SmcDB.StoredProcedureNonQuery("SaveShipment", new Dictionary<string, object> { { "@dokumentnr", doknr }, {"@shipmentnr", trackingnumber}, { "@trackingnr", trackingnumber },
                                            { "@paketnr", (i+1) }, { "@kommentar", String.IsNullOrEmpty(s_packagecomment[i])? "Paket " + (i + 1) : s_packagecomment[i] } }, "ProffixDB", ini);
                    }
                    System.Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehlerhafte Eingabe. Bitte prüfen Sie die Eingabe und probieren Sie es noch einmal. " + ex.Message);
            }
        }

        private void LoadPackage()
        {
            try
            {
                if (cbPostPakete.SelectedIndex == 0)
                {
                    btnPrevPost.Visible = false;
                }
                else
                {
                    btnPrevPost.Visible = true;
                }
                if (cbPostPakete.SelectedIndex == (int)(numPostAnzPakete.Value - 1))
                {
                    btnNextPost.Visible = false;
                }
                else
                {
                    btnNextPost.Visible = true;
                }
                if (cbDHLPakete.SelectedIndex == 0)
                {
                    btnPrevDHL.Visible = false;
                }
                else
                {
                    btnPrevDHL.Visible = true;
                }
                if (cbDHLPakete.SelectedIndex == (int)(numDHLAnzPakete.Value - 1))
                {
                    btnNextDHL.Visible = false;
                }
                else
                {
                    btnNextDHL.Visible = true;
                }
                txtPostGewicht.Text = d_packages[selectedPackage, 0].ToString();
                txtPostLaenge.Text = d_packages[selectedPackage, 1].ToString();
                txtPostBreite.Text = d_packages[selectedPackage, 2].ToString();
                txtPostHoehe.Text = d_packages[selectedPackage, 3].ToString();
                txtPostKundenref.Text = s_packagecomment[selectedPackage];
                cbPostPakettyp.SelectedIndex = i_packagetype[selectedPackage];
                txtDHLGewicht.Text = d_packages[selectedPackage, 0].ToString();
                txtDHLLaenge.Text = d_packages[selectedPackage, 1].ToString();
                txtDHLBreite.Text = d_packages[selectedPackage, 2].ToString();
                txtDHLHoehe.Text = d_packages[selectedPackage, 3].ToString();
                txtDHLKundenref.Text = s_packagecomment[selectedPackage];
                cbDHLPakettyp.SelectedIndex = i_packagetype[selectedPackage];
            }
            catch (Exception ex) { }
        }

        private bool SavePackage(string Gewicht, string Laenge, string Breite, string Hoehe, string Kundenreferenz, int Pakettyp)
        {
            try
            {
                if (!Decimal.TryParse(Gewicht, out d_packages[selectedPackage, 0]))
                {
                    MessageBox.Show("Das Gewicht muss eine Zahl sein!");
                    return false;
                }
                if (!Decimal.TryParse(Laenge, out d_packages[selectedPackage, 1]))
                {
                    MessageBox.Show("Die Länge muss eine Zahl sein!");
                    return false;
                }
                if (!Decimal.TryParse(Breite, out d_packages[selectedPackage, 2]))
                {
                    MessageBox.Show("Die Breite muss eine Zahl sein!");
                    return false;
                }
                if (!Decimal.TryParse(Hoehe, out d_packages[selectedPackage, 3]))
                {
                    MessageBox.Show("Die Höhe muss eine Zahl sein!");
                    return false;
                }
                s_packagecomment[selectedPackage] = Kundenreferenz;
                i_packagetype[selectedPackage] = Pakettyp;
                //l_lieferscheine = new List<string>()[];
                //for (int i = 0; i < clbPostLieferscheine.CheckedItems.Count; i++)
                //{
                //    //l_lieferscheine[selectedPackage, ]
                //}
                return true;
            }
            catch (Exception ex) { return false; }
        }

        private void txtEKontaktName_Enter(object sender, EventArgs e)
        {
            if (txtEKontaktName.Text.Equals("Empfängername..."))
            {
                txtEKontaktName.Text = "";
            }
        }

        private void txtEKontaktName_Leave(object sender, EventArgs e)
        {
            if (txtEKontaktName.Text.Equals(""))
            {
                txtEKontaktName.Text = "Empfängername...";
            }
        }

        private void txtEFirma_Enter(object sender, EventArgs e)
        {
            if (txtEFirma.Text.Equals("Firmenname..."))
            {
                txtEFirma.Text = "";
            }
        }

        private void txtEFirma_Leave(object sender, EventArgs e)
        {
            if (txtEFirma.Text.Equals(""))
            {
                txtEFirma.Text = "Firmenname...";
            }
        }

        private void txtEStrasse_Enter(object sender, EventArgs e)
        {
            if (txtEStrasse.Text.Equals("Strasse..."))
            {
                txtEStrasse.Text = "";
            }
        }

        private void txtEStrasse_Leave(object sender, EventArgs e)
        {
            if (txtEStrasse.Text.Equals(""))
            {
                txtEStrasse.Text = "Strasse...";
            }
        }

        private void txtEPLZ_Enter(object sender, EventArgs e)
        {
            if (txtEPLZ.Text.Equals("PLZ..."))
            {
                txtEPLZ.Text = "";
            }
        }

        private void txtEPLZ_Leave(object sender, EventArgs e)
        {
            if (txtEPLZ.Text.Equals(""))
            {
                txtEPLZ.Text = "PLZ...";
            }
        }

        private void txtEOrt_Enter(object sender, EventArgs e)
        {
            if (txtEOrt.Text.Equals("Ort..."))
            {
                txtEOrt.Text = "";
            }
        }

        private void txtEOrt_Leave(object sender, EventArgs e)
        {
            if (txtEOrt.Text.Equals(""))
            {
                txtEOrt.Text = "Ort...";
            }
        }

        private void txtELand_Enter(object sender, EventArgs e)
        {
            if (txtELand.Text.Equals("Land..."))
            {
                txtELand.Text = "";
            }
        }

        private void txtELand_Leave(object sender, EventArgs e)
        {
            if (txtELand.Text.Equals(""))
            {
                txtELand.Text = "Land...";
            }
        }

        private void txtETelefon_Enter(object sender, EventArgs e)
        {
            if (txtETelefon.Text.Equals("Telefon..."))
            {
                txtETelefon.Text = "";
            }
        }

        private void txtETelefon_Leave(object sender, EventArgs e)
        {
            if (txtETelefon.Text.Equals(""))
            {
                txtETelefon.Text = "Telefon...";
            }
        }

        private void txtEEmail_Enter(object sender, EventArgs e)
        {
            if (txtEEmail.Text.Equals("E-Mail..."))
            {
                txtEEmail.Text = "";
            }
        }

        private void txtEEmail_Leave(object sender, EventArgs e)
        {
            if (txtEEmail.Text.Equals(""))
            {
                txtEEmail.Text = "E-Mail...";
            }
        }

        private void txtKontaktName_Enter(object sender, EventArgs e)
        {
            if (txtKontaktName.Text.Equals("Absendername..."))
            {
                txtKontaktName.Text = "";
            }
        }

        private void txtKontaktName_Leave(object sender, EventArgs e)
        {
            if (txtKontaktName.Text.Equals(""))
            {
                txtKontaktName.Text = "Absendername...";
            }
        }

        private void txtFirma_Enter(object sender, EventArgs e)
        {
            if (txtFirma.Text.Equals("Firmenname..."))
            {
                txtFirma.Text = "";
            }
        }

        private void txtFirma_Leave(object sender, EventArgs e)
        {
            if (txtFirma.Text.Equals(""))
            {
                txtFirma.Text = "Firmenname...";
            }
        }

        private void txtStrasse_Enter(object sender, EventArgs e)
        {
            if (txtStrasse.Text.Equals("Strasse..."))
            {
                txtStrasse.Text = "";
            }
        }

        private void txtStrasse_Leave(object sender, EventArgs e)
        {
            if (txtStrasse.Text.Equals(""))
            {
                txtStrasse.Text = "Strasse...";
            }
        }

        private void txtPLZ_Enter(object sender, EventArgs e)
        {
            if (txtPLZ.Text.Equals("PLZ..."))
            {
                txtPLZ.Text = "";
            }
        }

        private void txtPLZ_Leave(object sender, EventArgs e)
        {
            if (txtPLZ.Text.Equals(""))
            {
                txtPLZ.Text = "PLZ...";
            }
        }

        private void txtOrt_Enter(object sender, EventArgs e)
        {
            if (txtOrt.Text.Equals("Ort..."))
            {
                txtOrt.Text = "";
            }
        }

        private void txtOrt_Leave(object sender, EventArgs e)
        {
            if (txtOrt.Text.Equals(""))
            {
                txtOrt.Text = "Ort...";
            }
        }

        private void txtTelefon_Enter(object sender, EventArgs e)
        {
            if (txtTelefon.Text.Equals("Telefon..."))
            {
                txtTelefon.Text = "";
            }
        }

        private void txtTelefon_Leave(object sender, EventArgs e)
        {
            if (txtTelefon.Text.Equals(""))
            {
                txtTelefon.Text = "Telefon...";
            }
        }

        private void txtEmail_Enter(object sender, EventArgs e)
        {
            if (txtEmail.Text.Equals("E-Mail..."))
            {
                txtEmail.Text = "";
            }
        }

        private void txtEmail_Leave(object sender, EventArgs e)
        {
            if (txtEmail.Text.Equals(""))
            {
                txtEmail.Text = "E-Mail...";
            }
        }
        private void btnSaveAddress_Click(object sender, EventArgs e)
        {
            try
            {
                //Empfänger speichern
                lbEName.Text = txtEKontaktName.Text.ToString().EndsWith("...") ? "" : txtEKontaktName.Text;
                lbEFirma.Text = txtEFirma.Text.ToString().EndsWith("...") ? "" : txtEFirma.Text;
                lbEStrasse.Text = txtEStrasse.Text.ToString().EndsWith("...") ? "" : txtEStrasse.Text;
                lbEPLZ.Text = txtEPLZ.Text.ToString().EndsWith("...") ? "" : txtEPLZ.Text;
                lbEOrt.Text = txtEOrt.Text.ToString().EndsWith("...") ? "" : txtEOrt.Text;
                lbLand.Text = txtELand.Text.ToString().EndsWith("...") ? "" : txtELand.Text;
                lbETelefon.Text = txtETelefon.Text.ToString().EndsWith("...") ? "" : txtETelefon.Text;
                if (txtEEmail.Text.Length > 50)
                {
                    lbEEmail.Text = txtEEmail.Text.Substring(0, 50) + "\r\n" + txtEEmail.Text.Substring(50);
                }
                else
                {
                    lbEEmail.Text = txtEEmail.Text.ToString().EndsWith("...") ? "" : txtEEmail.Text;
                }

                //Absender speichern
                lbName.Text = txtKontaktName.Text.ToString().EndsWith("...") ? "" : txtKontaktName.Text;
                lbFirma.Text = txtFirma.Text.ToString().EndsWith("...") ? "" : txtFirma.Text;
                lbStrasse.Text = txtStrasse.Text.ToString().EndsWith("...") ? "" : txtStrasse.Text;
                lbPLZ.Text = txtPLZ.Text.ToString().EndsWith("...") ? "" : txtPLZ.Text;
                lbOrt.Text = txtOrt.Text.ToString().EndsWith("...") ? "" : txtOrt.Text;
                lbTelefon.Text = txtTelefon.Text.ToString().EndsWith("...") ? "" : txtTelefon.Text;
                if (txtEmail.Text.Length > 20)
                {
                    lbEmail.Text = txtEmail.Text.Substring(0, 20) + "\r\n" + txtEmail.Text.Substring(20);
                }
                else
                {
                    lbEmail.Text = txtEmail.Text.ToString().EndsWith("...") ? "" : txtEmail.Text;
                }
                pnAdresse.Visible = true;
                pnlEditAddress.Visible = false;
            }
            catch (Exception ex) { }
        }

        private void linkBearbeiten_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                //Empfänger speichern
                txtEKontaktName.Text = lbEName.Text;
                if (txtEKontaktName.Text.Equals(""))
                {
                    txtEKontaktName.Text = "Empfängername...";
                }
                txtEFirma.Text = lbEFirma.Text;
                if (txtEFirma.Text.Equals(""))
                {
                    txtEFirma.Text = "Firmenname...";
                }
                txtEStrasse.Text = lbEStrasse.Text;
                if (txtEStrasse.Text.Equals(""))
                {
                    txtEStrasse.Text = "Strasse...";
                }
                txtEPLZ.Text = lbEPLZ.Text;
                if (txtEPLZ.Text.Equals(""))
                {
                    txtEPLZ.Text = "PLZ...";
                }
                txtEOrt.Text = lbEOrt.Text;
                if (txtEOrt.Text.Equals(""))
                {
                    txtEOrt.Text = "Ort...";
                }
                txtELand.Text = lbLand.Text;
                if (txtELand.Text.Equals(""))
                {
                    txtELand.Text = "Land...";
                }
                txtETelefon.Text = lbETelefon.Text;
                if (txtETelefon.Text.Equals(""))
                {
                    txtETelefon.Text = "Telefon...";
                }
                txtEEmail.Text = lbEEmail.Text;
                if (txtEEmail.Text.Equals(""))
                {
                    txtEEmail.Text = "E-Mail...";
                }

                //Absender speichern
                txtKontaktName.Text = lbName.Text;
                if (txtKontaktName.Text.Equals(""))
                {
                    txtKontaktName.Text = "Absendername...";
                }
                txtFirma.Text = lbFirma.Text;
                if (txtFirma.Text.Equals(""))
                {
                    txtFirma.Text = "Firmenname...";
                }
                txtStrasse.Text = lbStrasse.Text;
                if (txtStrasse.Text.Equals(""))
                {
                    txtStrasse.Text = "Strasse...";
                }
                txtPLZ.Text = lbPLZ.Text;
                if (txtPLZ.Text.Equals(""))
                {
                    txtPLZ.Text = "PLZ...";
                }
                txtOrt.Text = lbOrt.Text;
                if (txtOrt.Text.Equals(""))
                {
                    txtOrt.Text = "Ort...";
                }
                txtTelefon.Text = lbTelefon.Text;
                if (txtTelefon.Text.Equals(""))
                {
                    txtTelefon.Text = "Telefon...";
                }
                txtEmail.Text = lbEmail.Text.Replace("\r\n", "");
                if (txtEmail.Text.Equals(""))
                {
                    txtEmail.Text = "E-Mail...";
                }

                pnAdresse.Visible = false;
                pnlEditAddress.Visible = true;
            }
            catch (Exception ex) { }
        }

        private void frmPaketVersand_Load(object sender, EventArgs e)
        {
            pnlEditAddress.Location = new Point(0, 122);
            imgSaveDokNr.Location = new Point(241, 9);

        }

        private void numAnzahlPakete_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                int selectedIdPost = cbPostPakete.SelectedIndex;
                int selectedIdDPD = cbDPDPakete.SelectedIndex;
                int selectedIdDHL = cbDHLPakete.SelectedIndex;
                int selectedIdFedEx = cbFedExPakete.SelectedIndex;

                dok.numPackages = (int)numPostAnzPakete.Value;

                numDHLAnzPakete.Value = numPostAnzPakete.Value;

                cbPostPakete.Items.Clear();
                cbDPDPakete.Items.Clear();
                cbDHLPakete.Items.Clear();
                cbFedExPakete.Items.Clear();

                for (int i = 1; i <= dok.numPackages; i++)
                {
                    cbPostPakete.Items.Add("Paket " + i);
                    cbDPDPakete.Items.Add("Paket " + i);
                    cbDHLPakete.Items.Add("Paket " + i);
                    cbFedExPakete.Items.Add("Paket " + i);
                }

                cbPostPakete.SelectedIndex = dok.numPackages > selectedIdPost ? selectedIdPost : 0;
                cbDPDPakete.SelectedIndex = dok.numPackages > selectedIdDPD ? selectedIdDPD : 0;
                cbDHLPakete.SelectedIndex = dok.numPackages > selectedIdDHL ? selectedIdDHL : 0;
                cbFedExPakete.SelectedIndex = dok.numPackages > selectedIdFedEx ? selectedIdFedEx : 0;
            }
            catch (Exception ex) { }
        }

        private void imgEditDokNr_Click(object sender, EventArgs e)
        {
            imgSaveDokNr.Visible = true;
            imgEditDokNr.Visible = false;
            txtDokNr.Enabled = true;
        }

        private void imgSaveDokNr_Click(object sender, EventArgs e)
        {
            LoadDataFromDocNr(txtDokNr.Text);
        }

        private void LoadDataFromDocNr(string docnr)
        {
            try
            {
                DataTable dt = SmcDB.StoredProcedureDataTable("GetVersandFromLs", new Dictionary<string, object>() { { "DokNr", docnr } }, "ProffixDB", ini);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Das Dokument mit der Nummer " + txtDokNr.Text + " konnte nicht gefunden werden.");
                }
                else
                {
                    //Empfänger speichern
                    string adressNr = dt.Rows[0]["EAdressNr"].ToString();
                    lbAdressNr.Text = adressNr;
                    
                    if (String.IsNullOrEmpty(dt.Rows[0]["LEName"].ToString()) || String.IsNullOrEmpty(dt.Rows[0]["LEStrasse"].ToString()) || String.IsNullOrEmpty(dt.Rows[0]["LEOrt"].ToString()))
                    {
                        if (String.IsNullOrEmpty(dt.Rows[0]["EKontakt"].ToString()))
                        {
                            lbEName.Text = "--";
                        }
                        else
                        {
                            lbEName.Text = dt.Rows[0]["EKontakt"].ToString();
                        }
                        lbEFirma.Text = dt.Rows[0]["EName"].ToString();
                        lbEStrasse.Text = dt.Rows[0]["EStrasse"].ToString() + " " + dt.Rows[0]["EHausNr"].ToString();
                        lbEPLZ.Text = dt.Rows[0]["EPLZ"].ToString();
                        lbEOrt.Text = dt.Rows[0]["EOrt"].ToString();
                        lbLand.Text = dt.Rows[0]["ELand"].ToString();
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(dt.Rows[0]["LEKontakt"].ToString()))
                        {
                            lbEName.Text = "--";
                        }
                        else
                        {
                            lbEName.Text = dt.Rows[0]["LEKontakt"].ToString();
                        }
                        lbEFirma.Text = dt.Rows[0]["LEName"].ToString();
                        lbEStrasse.Text = dt.Rows[0]["LEStrasse"].ToString() + " " + dt.Rows[0]["LEHausNr"].ToString();
                        lbEPLZ.Text = dt.Rows[0]["LEPLZ"].ToString();
                        lbEOrt.Text = dt.Rows[0]["LEOrt"].ToString();
                        lbLand.Text = dt.Rows[0]["LELand"].ToString();
                    }

                    DataTable datata = SmcDB.StoredProcedureDataTable("GetTel", new Dictionary<string, object>() { { "doknr", docnr } }, "ProffixDB", ini);
                    lbETelefon.Text = datata.Rows[0]["tel"].ToString();
                    dok.LiefTel = datata.Rows[0]["tel"].ToString();
                    dok.LiefEMail = datata.Rows[0]["email"].ToString();

                    lbEEmail.Text = dok.LiefEMail;
                    dok.Lieferart = dt.Rows[0]["Lieferart"].ToString();
                    if (DateTime.Parse(dt.Rows[0]["Liefertermin"].ToString()) > DateTime.Today.AddDays(1))
                        dateLiefertermin.Value = DateTime.Parse(dt.Rows[0]["Liefertermin"].ToString());

                    //Absender speichern
                    dt = SmcDB.StoredProcedureDataTable("GetStammdaten", new Dictionary<string, object>() { { "username", "Lager" } }, "ProffixDB", ini);
                    lbName.Text = "Lager";
                    lbFirma.Text = dt.Rows[0]["Firma"].ToString();
                    lbStrasse.Text = dt.Rows[0]["Strasse"].ToString();
                    lbPLZ.Text = dt.Rows[0]["PLZ"].ToString();
                    lbOrt.Text = dt.Rows[0]["Ort"].ToString();
                    lbTelefon.Text = dt.Rows[0]["Telefon"].ToString();
                    lbEmail.Text = dt.Rows[0]["EMail"].ToString();

                    cbLieferart.Items.Clear();
                    dt = SmcDB.StoredProcedureDataTable("GetLieferarten", new Dictionary<string, object>(), "ProffixDB", ini);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        cbLieferart.Items.Add(dt.Rows[i]["Bezeichnung"].ToString());
                    }
                    if (!String.IsNullOrEmpty(dok.Lieferart))
                    {
                        int dokLieferartId = cbLieferart.FindString(dok.Lieferart);
                        if (dokLieferartId >= 0)
                        {
                            cbLieferart.SelectedIndex = dokLieferartId;
                        }
                        else if (cbLieferart.Items.Count > 0)
                        {
                            cbLieferart.SelectedIndex = 0;
                            dok.Lieferart = cbLieferart.SelectedItem.ToString();
                        }
                    }
                    else if (cbLieferart.Items.Count > 0)
                    {
                        cbLieferart.SelectedIndex = 0;
                        dok.Lieferart = cbLieferart.SelectedItem.ToString();
                    }

                    numPostAnzPakete.Value = 1;
                    //cbLieferart = dt.Rows[0][""];

                    txtDokNr.Text = docnr;
                    //dok.DocNr = txtDokNr.Text;
                    imgSaveDokNr.Visible = false;
                    imgEditDokNr.Visible = true;
                    txtDokNr.Enabled = false;
                    SetAbholarten(dok.Lieferart);
                }
            }
            catch (Exception ex) { }
        }

        private void cbLieferart_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetAbholarten(cbLieferart.Text);
        }

        private void cbPostPakete_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (selectedPackage >= 0)
                {
                    if (SavePackage(txtPostGewicht.Text, txtPostLaenge.Text, txtPostBreite.Text, txtPostHoehe.Text, txtPostKundenref.Text, cbPostPakettyp.SelectedIndex))
                    {
                        selectedPackage = cbPostPakete.SelectedIndex;
                        LoadPackage();
                    }
                    else
                    {
                        int temp = selectedPackage;
                        selectedPackage = -5;
                        cbPostPakete.SelectedIndex = temp;
                    }
                }
                else if (selectedPackage == -5)
                {
                    selectedPackage = cbPostPakete.SelectedIndex;
                }
                else
                {
                    selectedPackage = cbPostPakete.SelectedIndex;
                    LoadPackage();
                }
            }
            catch (Exception ex) { }
        }

        private void btnNextPost_Click(object sender, EventArgs e)
        {
            cbPostPakete.SelectedIndex = selectedPackage + 1;
        }

        private void btnPrevPost_Click(object sender, EventArgs e)
        {
            cbPostPakete.SelectedIndex = selectedPackage - 1;
        }

        private void cbDHLPakettyp_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbDHLPakettyp.SelectedIndex == 0)
                {
                    txtDHLLaenge.Enabled = true;
                    txtDHLBreite.Enabled = true;
                    txtDHLHoehe.Enabled = true;
                }
                else
                {
                    txtDHLLaenge.Enabled = false;
                    txtDHLBreite.Enabled = false;
                    txtDHLHoehe.Enabled = false;
                    ComboBoxItem Selected = (ComboBoxItem)cbDHLPakettyp.SelectedItem;
                    txtDHLLaenge.Text = Selected.Laenge.ToString();
                    txtDHLBreite.Text = Selected.Breite.ToString();
                    txtDHLHoehe.Text = Selected.Hoehe.ToString();
                }
            }
            catch (Exception ex) { }
        }

        private void cbPostPakettyp_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbPostPakettyp.SelectedIndex == 0)
            {
                txtPostLaenge.Enabled = true;
                txtPostBreite.Enabled = true;
                txtPostHoehe.Enabled = true;
            }
            else
            {
                txtPostLaenge.Enabled = false;
                txtPostBreite.Enabled = false;
                txtPostHoehe.Enabled = false;
                ComboBoxItem Selected = (ComboBoxItem)cbPostPakettyp.SelectedItem;
                txtPostLaenge.Text = Selected.Laenge.ToString();
                txtPostBreite.Text = Selected.Breite.ToString();
                txtPostHoehe.Text = Selected.Hoehe.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<int> NichtGesendet = new List<int>();
            SavePackage(txtDHLGewicht.Text, txtDHLLaenge.Text, txtDHLBreite.Text, txtDHLHoehe.Text, txtDHLKundenref.Text, cbDHLPakettyp.SelectedIndex);
            try
            {
                if (!string.IsNullOrEmpty(lbETelefon.Text) && !string.IsNullOrEmpty(lbTelefon.Text))
                {
                    if (dateLiefertermin.Value >= DateTime.Today.AddDays(1) && dateLiefertermin.Value < DateTime.Today.AddDays(11))
                    {
                        for (int i = 0; i < dok.numPackages; i++)
                        {
                            if (d_packages[i, 0] <= 0 || d_packages[i, 1] <= 0 || d_packages[i, 2] <= 0 || d_packages[i, 3] <= 0)
                            {
                                MessageBox.Show("Paket " + (i + 1) + " enthält ungültige Werte!");
                                return;
                            }
                        }
                        object postAddress = new object();
                        //if (lbEStrasse.Text == dok.LiefStrasse)
                        //{
                        postAddress = new
                        {
                            postalCode = lbEPLZ.Text,
                            cityName = lbEOrt.Text,
                            countryCode = string.IsNullOrEmpty(lbLand.Text) ? "CH" : lbLand.Text,
                            //addressLine1 = string.IsNullOrEmpty(dok.Adresszeile1) ? dok.Adresszeile2 : dok.Adresszeile1,
                            addressLine1 = lbEStrasse.Text
                        };

                        //}
                        string ShipperName = lbName.Text;
                        string ShipperCompanyName = lbFirma.Text;
                        string ShipperPhoneNumber = lbTelefon.Text;
                        string ShipperStreet = lbStrasse.Text;
                        string ShipperCity = lbOrt.Text;
                        lbEName.Text = string.IsNullOrEmpty(lbEName.Text) ? "-" : lbEName.Text;
                        string ShipperPostalCode = lbPLZ.Text;
                        List<string> TrackNr = new List<string>();
                        for (int i = 0; i < dok.numPackages; i++)
                        {
                            try
                            {
                                //json start
                                object jsonstring = new
                                {
                                    plannedShippingDateAndTime = dateLiefertermin.Value.ToString("yyyy-MM-ddT16:30:00'GMT'+01:00"),
                                    pickup = new
                                    {
                                        isRequested = cbAbholart.Text == "Abholung beantragen" ? true : false
                                    },
                                    productCode = "N",
                                    localProductCode = "N",
                                    accounts = new object[] {
                                new {
                                    typeCode = "shipper",
                                    number = ini.GetValue(ini.Environment + "_DHL", "Kundennummer")
                                }
                            },
                                    customerDetails = new
                                    {
                                        shipperDetails = new
                                        {
                                            postalAddress = new
                                            {
                                                postalCode = lbPLZ.Text,
                                                cityName = lbOrt.Text,
                                                countryCode = "CH",
                                                addressLine1 = lbStrasse.Text
                                            },
                                            contactInformation = new
                                            {
                                                phone = lbTelefon.Text,
                                                companyName = lbFirma.Text,
                                                fullName = lbName.Text
                                            }
                                        },
                                        receiverDetails = new
                                        {
                                            postalAddress = postAddress,
                                            contactInformation = new
                                            {
                                                phone = string.IsNullOrEmpty(lbETelefon.Text) ? "-" : lbETelefon.Text,
                                                companyName = lbEFirma.Text,
                                                fullName = string.IsNullOrEmpty(lbEName.Text) ? "-" : lbEName.Text
                                            }
                                        }
                                    },
                                    content = new
                                    {
                                        packages = new object[]
                                        {
                                new {
                                    weight = d_packages[i, 0],
                                    dimensions = new
                                    {
                                        length = d_packages[selectedPackage, 1],
                                        width = d_packages[selectedPackage, 2],
                                        height = d_packages[selectedPackage, 3]
                                    },
                                    customerReferences = new object[]
                                    {
                                        new {
                                            value = String.IsNullOrEmpty(s_packagecomment[i]) ? "Paket " + (i + 1) : s_packagecomment[i].Substring(0, Math.Min(s_packagecomment[i].Length, 35)),
                                            typeCode = "CU"
                                        }
                                    }
                                }
                                        },
                                        isCustomsDeclarable = false,
                                        description = String.IsNullOrEmpty(s_packagecomment[i]) ? "Paket " + (i + 1) : s_packagecomment[i].Substring(0, Math.Min(s_packagecomment[i].Length, 35)),
                                        incoterm = cbLieferart.Text.Substring(0, 3),
                                        unitOfMeasurement = "metric"
                                    }
                                };

                                var dhl_json = JsonConvert.SerializeObject(jsonstring);
                                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(ini.GetValue(ini.Environment + "_DHL", "URL"));
                                String encoded = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(ini.GetValue(ini.Environment + "_DHL", "Benutzername") + ":" + ini.GetValue(ini.Environment + "_DHL", "Passwort")));
                                httpWebRequest.Headers.Add("Authorization", "Basic " + encoded);
                                httpWebRequest.ContentType = "application/json";
                                httpWebRequest.Method = "POST";
                                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                                {
                                    streamWriter.Write(dhl_json);
                                }
                                HttpWebResponse httpResponse;
                                try
                                {
                                    httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                                }
                                catch (WebException ex)
                                {
                                    httpResponse = ex.Response as HttpWebResponse;
                                }
                                dynamic result;
                                using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
                                {
                                    result = JsonConvert.DeserializeObject(streamReader.ReadToEnd());
                                }
                                //MessageBox.Show(result.ToString());
                                if (result.SelectToken("detail") != null)
                                {
                                    MessageBox.Show(result.SelectToken("detail").ToString());
                                    NichtGesendet.Add(i + 1);
                                    continue;
                                }
                                IncreaseVolume();
                                string shipmentidentificationnumber = result.SelectToken("shipmentTrackingNumber").ToString();
                                var bytes = Convert.FromBase64String(result.SelectToken("documents[0].content").ToString());
                                Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Etiketten"));
                                var stream = new FileStream(Path.Combine(Environment.CurrentDirectory, "Etiketten", "DHL_Etikette_" + shipmentidentificationnumber + ".pdf"), FileMode.CreateNew);
                                var writer = new BinaryWriter(stream);
                                writer.Write(bytes, 0, bytes.Length);
                                writer.Close();
                                Process.Start(Path.Combine(Environment.CurrentDirectory, "Etiketten", "DHL_Etikette_" + shipmentidentificationnumber + ".pdf"));

                                //PrintDocument(Path.Combine(Environment.CurrentDirectory, "Etiketten", "DHL_Etikette_" + shipmentidentificationnumber + ".pdf"));
                                //    using (var client = new System.Net.Http.HttpClient())
                                //{
                                //        HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(Endpoint + RESTApi);
                                //        httpWebRequest.ContentType = "application/json";
                                //        httpWebRequest.Method = "POST";
                                //        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes($"" + ini.GetValue(ini.Environment + "_DHL", "Benutzername") + ":" + ini.GetValue(ini.Environment + "_DHL", "Passwort"))));
                                //    response = await client.PostAsync(ini.GetValue(ini.Environment + "_DHL", "URL"), new StringContent(dhl_json, Encoding.UTF8, "application/json"));
                                //    responsestring = response.Content.ReadAsStringAsync().Result;
                                //    responsejson = JObject.Parse(responsestring);
                                //}
                                //MessageBox.Show(responsestring);
                                ////string shipmentnotification = responsejson.SelectToken("ShipmentResponse.Notification[0].Message").ToString();
                                //if (!string.IsNullOrEmpty(shipmentnotification))
                                //{
                                //    MessageBox.Show("Es gab einen Fehler bei der Verarbeitung der Anfrage. Die Meldung von der DHL lautet: " + shipmentnotification);
                                //    this.Close();
                                //}
                                //shipmentidentificationnumber = responsejson.SelectToken("ShipmentResponse.ShipmentIdentificationNumber").ToString();
                                string trackingnumber = result.SelectToken("packages[0].trackingNumber").ToString();
                                TrackNr.Add(trackingnumber);
                                //bytes = Convert.FromBase64String(responsejson.SelectToken("ShipmentResponse.LabelImage[0].GraphicImage").ToString());
                                //Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Etiketten"));
                                //stream = new FileStream(Path.Combine(Environment.CurrentDirectory, "Etiketten", "DHL_Etikette_" + shipmentidentificationnumber + ".pdf"), FileMode.CreateNew);
                                //writer = new BinaryWriter(stream);
                                //writer.Write(bytes, 0, bytes.Length);
                                //writer.Close();
                                //Process.Start(Path.Combine(Environment.CurrentDirectory, "Etiketten", "DHL_Etikette_" + shipmentidentificationnumber + ".pdf"));

                                SmcDB.StoredProcedureNonQuery("SaveDHLShipment", new Dictionary<string, object> { { "@dokumentnr", txtDokNr.Text }, {"@shipmentnr", shipmentidentificationnumber}, { "@trackingnr", trackingnumber },
                                            { "@paketnr", (i+1) }, { "@kommentar", String.IsNullOrEmpty(s_packagecomment[i])? "Paket " + (i + 1) : s_packagecomment[i] } }, "ProffixDB", ini);

                            }
                            catch
                            {
                                MessageBox.Show("Paket " + (i + 1) + " konnte nicht versendet werden.");
                                NichtGesendet.Add(i + 1);
                            }
                        }
                        //string EmpfängerMail;
                        //string Sprache = "D";
                        //using (SqlConnection conn = new SqlConnection(pxconnstring))
                        //{
                        //    conn.Open();
                        //    SqlCommand cmd = new SqlCommand("GetCustomerEmail", conn);
                        //    cmd.CommandType = CommandType.StoredProcedure;
                        //    cmd.Parameters.AddWithValue("@DokNr", lieferschein.DocNr);
                        //    using (SqlDataReader rdr = cmd.ExecuteReader())
                        //    {
                        //        rdr.Read();
                        //        Sprache = rdr["Sprache"].ToString();
                        //        Sprache = rdr["Sprache"].ToString();
                        //        EmpfängerMail = rdr["Mail"].ToString();
                        //    }
                        //}
                        if (!string.IsNullOrEmpty(ini.GetValue(ini.Environment + "_Mail", "Sender")) && TrackNr.Count > 0 && !string.IsNullOrEmpty(lbEEmail.Text))
                        {
                            SendMail(TrackNr);
                        }
                        if (NichtGesendet.Count > 0)
                        {
                            string output = System.Environment.NewLine;
                            foreach (int i in NichtGesendet)
                            {
                                output += "Paket " + i.ToString() + System.Environment.NewLine;
                            }
                            MessageBox.Show("Für die folgenden Pakete konnte keine Sendung erstellt werden:" + output);
                        }
                        else
                        {
                            this.Close();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Es wird eine Telefonnummer für den Sender und Empfänger benötigt.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Beim Erstellen der Lieferung ist ein Fehler aufgetreten" + ex.Message);
            }

        }

        private void cbDHLPakete_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (selectedPackage >= 0)
                {
                    if (SavePackage(txtDHLGewicht.Text, txtDHLLaenge.Text, txtDHLBreite.Text, txtDHLHoehe.Text, txtDHLKundenref.Text, cbDHLPakettyp.SelectedIndex))
                    {
                        selectedPackage = cbDHLPakete.SelectedIndex;
                        LoadPackage();
                    }
                    else
                    {
                        int temp = selectedPackage;
                        selectedPackage = -5;
                        cbDHLPakete.SelectedIndex = temp;
                    }
                }
                else if (selectedPackage == -5)
                {
                    selectedPackage = cbDHLPakete.SelectedIndex;
                }
                else
                {
                    selectedPackage = cbDHLPakete.SelectedIndex;
                    LoadPackage();
                }
            }
            catch (Exception ex) { }
        }

        private void numDHLAnzPakete_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                int selectedIdPost = cbPostPakete.SelectedIndex;
                int selectedIdDPD = cbDPDPakete.SelectedIndex;
                int selectedIdDHL = cbDHLPakete.SelectedIndex;
                int selectedIdFedEx = cbFedExPakete.SelectedIndex;

                dok.numPackages = (int)numDHLAnzPakete.Value;


                numPostAnzPakete.Value = numDHLAnzPakete.Value;
                cbPostPakete.Items.Clear();
                cbDPDPakete.Items.Clear();
                cbDHLPakete.Items.Clear();
                cbFedExPakete.Items.Clear();

                for (int i = 1; i <= dok.numPackages; i++)
                {
                    cbPostPakete.Items.Add("Paket " + i);
                    cbDPDPakete.Items.Add("Paket " + i);
                    cbDHLPakete.Items.Add("Paket " + i);
                    cbFedExPakete.Items.Add("Paket " + i);
                }

                cbPostPakete.SelectedIndex = dok.numPackages > selectedIdPost ? selectedIdPost : 0;
                cbDPDPakete.SelectedIndex = dok.numPackages > selectedIdDPD ? selectedIdDPD : 0;
                cbDHLPakete.SelectedIndex = dok.numPackages > selectedIdDHL ? selectedIdDHL : 0;
                cbFedExPakete.SelectedIndex = dok.numPackages > selectedIdFedEx ? selectedIdFedEx : 0;
            }
            catch (Exception ex) { }
        }


        private void btnNextDHL_Click(object sender, EventArgs e)
        {
            cbDHLPakete.SelectedIndex = selectedPackage + 1;
        }

        private void btnPrevDHL_Click(object sender, EventArgs e)
        {
            cbDHLPakete.SelectedIndex = selectedPackage - 1;
        }

        private void cbDHLPakettyp_TextChanged(object sender, EventArgs e)
        {
            if (cbDHLPakettyp.FindStringExact(cbDHLPakettyp.Text) >= 0 && !string.IsNullOrEmpty(cbDHLPakettyp.Text))
            {
                cbDHLPakettyp.SelectedItem = cbDHLPakettyp.Items[cbDHLPakettyp.FindStringExact(cbDHLPakettyp.Text)];
            }
        }

        private void cbFedExPakettyp_TextChanged(object sender, EventArgs e)
        {
            if (cbFedExPakettyp.FindStringExact(cbFedExPakettyp.Text) >= 0 && !string.IsNullOrEmpty(cbFedExPakettyp.Text))
            {
                cbFedExPakettyp.SelectedItem = cbFedExPakettyp.Items[cbFedExPakettyp.FindStringExact(cbFedExPakettyp.Text)];
            }
        }

        private void cbDPDPakettyp_TextChanged(object sender, EventArgs e)
        {
            if (cbDPDPakettyp.FindStringExact(cbDPDPakettyp.Text) >= 0 && !string.IsNullOrEmpty(cbDPDPakettyp.Text))
            {
                cbDPDPakettyp.SelectedItem = cbDPDPakettyp.Items[cbDPDPakettyp.FindStringExact(cbDPDPakettyp.Text)];
            }
        }

        private void cbPostPakettyp_TextChanged(object sender, EventArgs e)
        {
            if (cbPostPakettyp.FindStringExact(cbPostPakettyp.Text) >= 0 && !string.IsNullOrEmpty(cbPostPakettyp.Text))
            {
                cbPostPakettyp.SelectedItem = cbPostPakettyp.Items[cbPostPakettyp.FindStringExact(cbPostPakettyp.Text)];
            }
        }

        class ComboBoxItem
        {
            string displayValue;
            public decimal Hoehe;
            public decimal Laenge;
            public decimal Breite;

            public ComboBoxItem(string d, decimal h, decimal l, decimal b)
            {
                displayValue = d;
                Hoehe = h;
                Laenge = l;
                Breite = b;
            }

            public override string ToString()
            {
                return displayValue;
            }
        }

        public void SendMail(List<string> trackingNummern)
        {
            try
            {
                string absender = ini.GetValue(ini.Environment + "_Mail", "Sender");
                string subject = "DHL Sendung wurde in Auftrag gegeben";
                string body = "Absender: " + lbFirma.Text + Environment.NewLine
                     + "Empfänger: " + lbEFirma.Text + Environment.NewLine
                     + "Versandtermin: " + dateLiefertermin.Value.ToString("dd-MM-yyyy") + Environment.NewLine
                     + "Anzahl Pakete: " + dok.numPackages + Environment.NewLine
                     + Environment.NewLine
                     + "Die Sendung kann unter folgender URL getrackt werden: https://www.dhl.ch/exp-de/express/sendungsverfolgung.html" + Environment.NewLine
                     + Environment.NewLine
                     + "Tracking-Nr.:" + Environment.NewLine;

                foreach (string tracknr in trackingNummern)
                {
                    body += tracknr + Environment.NewLine;
                }

                try
                {
                    var client = new SmtpClient(ini.GetValue(ini.Environment + "_Mail", "SMTP"), Int32.Parse(ini.GetValue(ini.Environment + "_Mail", "Port")))
                    {
                        Credentials = new NetworkCredential(ini.GetValue(ini.Environment + "_Mail", "Sender"), ini.GetValue(ini.Environment + "_Mail", "Password")),
                        EnableSsl = true
                    };
                    client.Send(ini.GetValue(ini.Environment + "_Mail", "Sender"), lbEEmail.Text, subject, body);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Exception");
                }
            }
            catch (Exception ex) { }
        }

        private async void IncreaseVolume()
        {
            try
            {
                int ProgrammId = 159;
                SmcPxREST rest = new SmcPxREST("https://ts.smc-computer.ch:4343/pxapi/v4", "REST", "A<m*VpjuZ:nu}bSf,WkQ", "PXSMC", "VOL");
                await rest.Patch<string>("/Extension/Okolo/Programmierung/" + ProgrammId + "/increment", "", "PxSessionId", err: err);
            }
            catch (Exception ex) {
            }
        }
    }
}
