using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using iQDataProvider;
using iQPUBLIC;
using System.Linq;
using BarcodeReader;

namespace BOLSET1IQ
{
    static class Module1
    {
        //public static int AmtNum = 0;
        //public static string BillOfLading_Synonyms = "BOL#BillOfLading";        
        //public static RetStructF3 returnZones = new RetStructF3();

        #region IQ Class Objects
        public static RetStructF3 objRetStructF3 = new RetStructF3();
        public static RetStructF5 objRetStructF5 = new RetStructF5();
        public static RetStructF7 objRetStructF7 = new RetStructF7();
        public static RetStructF27 objStructF27 = new RetStructF27();
        #endregion

        #region IQDataProvider Objects
        public static ClsCNC oClsCnc = new ClsCNC();
        public static clsCNCSBR ObjSbr = new clsCNCSBR();
        // public static clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();
        #endregion

        #region Lists for clscncword
        public static List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
        public static List<int> conflevel = new List<int>();
        public static List<string> lstConfFlag = new List<string>();    //for flag 
        public static List<string> lstRemarks = new List<string>();    //for  Remark
        #endregion

        public static DataSet dsKeys = new DataSet();
        //public static DataTable dt = new DataTable();

        //For MicroROI
        public static int MX1, MX2, MY1, MY2;
        // public static int MX1 = 1381, MX2 = 2469, MY1 = 231, MY2 = 501;
        //public static clsCnCWord[] TopKeywords = new clsCnCWord[10];

        #region All Regex
        //public static Regex NewRegexForAlphanumeric = new Regex("[^a-zA-Z0-9-,.#/:;_*]+");
        public static Regex NewRegexForAlphanumeric = new Regex("[^a-zA-Z0-9-,.#/:;_*~&]+");
        public static Regex regex = new Regex("^([a-zA-Z]*[-,.#/:;_*]?[0-9]+[-,.#/:;_*]?[a-zA-Z]*[-,.#:;/_*]?)*$");
        public static Regex regex1 = new Regex("^([a-zA-Z]*[-,./#:;_*]?[a-zA-Z]*[-,./#:;_*]?[0-9]+[-,./#:;_*]?)*$");
        public static Regex regexVal = new Regex("^([a-zA-Z]*[-,.#/_~@]?[0-9]+[-,.#/_~@]?[a-zA-Z]*[-,.#/_~@]?)*$");
        #region Date Regex
        public static Regex regex1Val = new Regex("^([a-zA-Z]*[-,./#_~@]?[a-zA-Z]*[-,./#_~@]?[0-9]+[-,./#_~@]?)*$");
        public static Regex InvalidDate = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]*(Map|MAP|mAP|MAp)(\d{4}|(\d{3}[a-zA-Z])|(\d{1}[a-zA-Z]\d{2}))");
        public static Regex MDYwithChar = new Regex(@"(1[012]|0?[1-9])[/](3[01]|[12][0-9]|0?[1-9]|[1-9])[/](\d{4})(?:/[A-Za-z])?");
        public static Regex MDYwithComma = new Regex(@"((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(3[01]|[12][0-9]|0?[1-9]|[1-9]),\s((?:[0-9]{2})?[0-9]{2}))");
        public static Regex DthMY = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])th\s(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s((?:[0-9]{2})?[0-9]{2})");
        public static Regex DMY = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])(\s*)(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s((?:[0-9]{2})?[0-9]{2})");
        public static Regex MDY = new Regex(@"((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)[ ](3[01]|[12][0-9]|0?[1-9]|[1-9])[ ](\d{4}))");
        #endregion
        //public static Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$");       //HH:MM
        //public static Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$|(([0-1][0-9])|([2][0-3])):([0-5][0-9]):([0-5][0-9])");      //HH:MM or HH:MM:SS
        #region Time Regex
        public static Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$|(([0-9]|[0-1][0-9])|([2][0-3])):([0-5][0-9]):([0-5][0-9])|(^([0-9]|[0-1][0-9]|[2][0-3]):([0-5][0-9])(\s{0,1})(AM|PM|Am|aM|pM|Am|Pm|am|pm{2,2}))");    //HH:MM or HH:MM:SS with AM/PM
        public static Regex Timewithoutcolon = new Regex(@"(([0-9])(\s{0,1})(AM|PM|Am|aM|pM|Am|Pm|am|pm{2,2}))");
        #endregion
        #region Telephone Regex
        public static Regex RegTel = new Regex(@"(\d{3}\s+\d{3}\s+\d{4}|\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\-\s*\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\.\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*/\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\-\s*\d{3}\s*/\s*\d{4}|\d{3}\s*/\s*\d{3}\s*/\s*\d{4})");
        public static Regex RegTelwithMob = new Regex(@"(\d{3}\s+\d{3}\s+\d{4}|\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\-\s*\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\.\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*/\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\-\s*\d{3}\s*/\s*\d{4}|\d{3}\s*/\s*\d{3}\s*/\s*\d{4}|(\d{10}))");
        #endregion
        #endregion

        public static List<string> ShouldNotStartWith = new List<string>() { "ORDERNUMBER","ORDERNO" };
        public static List<string> ShouldNotContains = new List<string>() { "ATN", "ATTN", "DO NOT STACK","DONOTSTACK","ORDER","LOCATION", "CALL", "APPT", "PCS", "PLT", "NOTIFY", "FOR", "UP","LBS","HOUR","BORDER","MABD","SHIP","OF","LADING","FREIGHT","PICKUP","CODE","STREET","APT","FROM","DRIVE","CLAS","LIFT","LIFTGATE","REFERENCE","PAGE","HANDLEWITHCARE","NORTH","HANDLE WITH CARE","DATE","RECEIVING","FREIGHT","11 THIS","COLLECT","PREPAID" };
        // public static Regex MobNo = new Regex("^[0-9]{10}$");
        public static List<string> SpecialCharacters = new List<string>() {"\"", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "{", "}", "[", "]", ":", ";", ",", "<", ">", ".", "?", "/", "~", "`" ,"|", "•","·" };
        public static List<string> SplCharsToMakeRed = new List<string>() { ".", ",", "/", ":", ";", "(", ")", "\"", "!","&","_","*" };
        public static List<string> IgnoreWordforLeftSplChar = new List<string>() {":","#" };//if leftword of value ends with : or # then ignore that word

        #region Variables for Hazmat and ERT
        //To store HazmatTel for ERT Code
        public static string EMerTel = string.Empty;
        public static String[] ExcludeWords = { ":", ".", "|", "#", "#:", "CHEMTREC", "IS", "VERISK", "3E", "INFOTRAC" };
        #endregion

        #region RA keywords
        public static List<string> RAKeywords = new List<string>() {"RA","RA#","RMA","RMA#", "RGA", "RGA#", "RETURN AUTHORIZATION","RETURNAUTHORIZATION" };
        #endregion

            #region Search Keywords from Document
        public static bool bFoundSearchKeywordTable = false;
        public static string imageName = "";

        public static string ProNo = "";

        public static DataSet MakeDs(clsCncMetaData oMeta, int intCurrPageNumber)
        {
            DataSet ds = new DataSet();
            //bool bFoundSearchKeywordTable = false;
            if (!bFoundSearchKeywordTable)
            {
                //iQPUBLIC.PublicComponents.htMyVariable = new Hashtable();
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword") == true)
                {
                    iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSet1_keyword");
                }
                DataTable dtkeyword = F3_SearchKeyword(oMeta, intCurrPageNumber);
                ds.Tables.Add(dtkeyword);
                iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSet1_keyword", ds);
                bFoundSearchKeywordTable = true;
            }
            return ds;

        }

        #region F3searchkeyword for Multi Page

        public static DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        {
            var dtKeywordInfo = new DataTable();
            dtKeywordInfo = MakeDt();
            try
            {
                // -----------------------------------------BILL OF LADING NUMBER/PURCHASE ORDER NUMBER/SHIPPER NUMBER KEYWORD------------------------------------------------------------
                // Dim sKeyarray() As String = {"Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "lhre Bestellung per FAX", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero"}

                //var sKeyarray1 = new string[] { "Bill of Lading Number", "BOL#","BOL", "B/L #","Bill of Lading No", "BOL No","Bill of L ading No", "BOL ID","Bill of Lading #",
                //                               "BOL Number","Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","Po", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No","PO#(s)",
                //                               "Customers PO","po#","cust po#","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","Customer PO#","Consisnee's Refer.noe/po No","P.O. No",
                //                               "Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","P/O","CUST. ORDER NO.","PO Reference", "Shipper Number",  "Ship ID", "Shipper Ref",
                //                               "Shipper No", "Shipment",  "Shipping Order No",  "Shipment ID","Shipment No","Shipment number","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.",
                //                               "Emergency", "Emergency Phone Numbers","Emergency Response Phone#","Emergency Contact","Emergency Response Phone Number","EMERGENCYCONTACTNUMBER:","24hrEMERGENCY PHONENUMBERS","Shipment Numbers",
                //                               "CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone","Emergencies","ChemTel","Chem tel",
                //                               "CHEMTREC","INFOTRAC","BOLNO","Shippers Release No","Ship Ref","Cust order","Load","Load ID","Load Number","Load No","SHIPPER NUM","shpmt num",
                //                               "CUSTOMER ORDER NUM","RA","RMA","RMA#","RGA","RGA#","RA#","SHIPPER REF NUMBER","Shipment Reference Number","Shipper NOS","RETURN AUTHORIZATION","BOL/SHIPMENT NO","BE OF LADING NO","BIL OF LADING NO","Bill Number",
                //                               "ponumber","shipperno","Customer POs:","Bill of Lading#","B/L#","Ship ID#:","B0L#","B L#:","P O #:","P/O#","P O. NUMBER(S)","PO NUM BER","BOL/Order#:","DO/BOL","SH PPER NUMBER","PO NUMBE R:","CUSTOMER ORDER #","SHIPPER#","SALES ORDER NBR/CUSTOMER P.O",
                //                               "Bil of Lading Number:","BIE OF LADING NO","LOAD#","BILL OF LADING - ME","SHIPMEN T NUMBER","SHIPMEN T NUMBE R","BILLOF LADING NUMBER","CUSTOMER PO INFORMATION","B0 L #","BILL OF LADING / BOOKING NUMBER","Shipper ID","BOLNUMBER",
                //                               "BILLOFLADING","P O NO","SID/BOL#:","BOLNBR","SHIPPERNOS","Shippers #","SHIPPER #","BOL NO (LD#)","CUSTOMER ORDER NO(S)","ORDER ID / PO","B L #:","BOL/SHIPMENT#","BL#","SHIPPER REF #/ORIG BOL #","BILL OF LADING / PACKING SLIP","SHIP NO","SHIPMENT IDENTIFICATION NO",
                //                               "P0/ORDER NUMBERS","SHIPPERS REFERENCE","DELIVERY/PO NUMBER","BDL OF LADING NO","CUSTOMER ORDER N","CUSTOME R P.O.","SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","BILL OF LADINGNUMBER","PO/REFERENCE NO","P0 #","CUSTOMER P0",
                //                                "ORDER/BILL OF LADING NO","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING REF NO","SHIPPER REFERENCE NO","BHI OF LADING #","BILL OFLADING NO","CUSTOMERORDER NUM","BDL NO","SHIPPERSNO","Shipper Reference","Shipment #","Underlying BOLs","CustomerPO#",
                //                                "Bill Of Lading Nos","BOL NUM","P. O. NUMBER","BM OF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BILL OF LAD ING NUMBER"};

                //var sKeyarray2 = new string[] { "Bill of Lading Number", "BOL#","BOL", "B/L #","Bill of Lading No", "BOL No","Bill of L ading No", "BOL ID","Bill of Lading #",
                //                               "BOL Number","Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","Po", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No","PO#(s)",
                //                               "Customers PO","po#","cust po#","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","Customer PO#","Consisnee's Refer.noe/po No","P.O. No",
                //                               "Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","P/O","CUST. ORDER NO.","PO Reference", "Shipper Number",  "Ship ID", "Shipper Ref",
                //                               "Shipper No", "Shipment",  "Shipping Order No",  "Shipment ID","Shipment No","Shipment number","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.",
                //                               "BOLNO","Shippers Release No","Ship Ref","Cust order","SHIPPER NUM","shpmt num","CUSTOMER ORDER NUM","RA","RMA","RMA#","RGA","RGA#","RA#","SHIPPER REF NUMBER","Shipment Reference Number","Shipper NOS","RETURN AUTHORIZATION","BOL/SHIPMENT NO","BE OF LADING NO","BIL OF LADING NO","Bill Number",
                //                               "ponumber","shipperno","Customer POs:","Bill of Lading#","B/L#","Ship ID#:","B0L#","B L#:","P O #:","P/O#","P O. NUMBER(S)","PO NUM BER","BOL/Order#:","DO/BOL","SH PPER NUMBER","PO NUMBE R:","CUSTOMER ORDER #","SHIPPER#","SALES ORDER NBR/CUSTOMER P.O",
                //                               "Bil of Lading Number:","BIE OF LADING NO","BILL OF LADING - ME","SHIPMEN T NUMBER","SHIPMEN T NUMBE R","BILLOF LADING NUMBER","CUSTOMER PO INFORMATION","B0 L #","BILL OF LADING / BOOKING NUMBER","Shipper ID","BOLNUMBER",
                //                               "BILLOFLADING","P O NO","SID/BOL#:","BOLNBR","SHIPPERNOS","Shippers #","SHIPPER #","BOL NO (LD#)","CUSTOMER ORDER NO(S)","ORDER ID / PO","B L #:","BOL/SHIPMENT#","BL#","SHIPPER REF #/ORIG BOL #","BILL OF LADING / PACKING SLIP","SHIP NO","SHIPMENT IDENTIFICATION NO",
                //                               "P0/ORDER NUMBERS","SHIPPERS REFERENCE","DELIVERY/PO NUMBER","BDL OF LADING NO","CUSTOMER ORDER N","CUSTOME R P.O.","SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","BILL OF LADINGNUMBER","PO/REFERENCE NO","P0 #","CUSTOMER P0",
                //                                "ORDER/BILL OF LADING NO","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING REF NO","SHIPPER REFERENCE NO","BHI OF LADING #","BILL OFLADING NO","CUSTOMERORDER NUM","BDL NO","SHIPPERSNO","Shipper Reference","Shipment #","Underlying BOLs","CustomerPO#",
                //                                "Bill Of Lading Nos","BOL NUM","P. O. NUMBER","BM OF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BILL OF LAD ING NUMBER"};

                var sKeyarray1 = new string[] {"SHIPPER BILL OF LADING NUMBER:","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING / BOOKING NUMBER","BILL OF LADING / PACKING SLIP","Bill of Lading Number",
                                               "Bil of Lading Number:","BILL OF LADING REF NO","Bill Of Lading Nos","ORDER/BILL OF LADING NO","Bill of Lading Num",
                                               "Bill of Lading No","BIL OF LADING NO","BILL OF LADING - ME","Bill of Lading #","Bill of Lading#","Bill of Lading","Pro/BOL Number",
                                               "BOL Number","BOLNUMBER","BOL NUM","BOL NBR","BOLNBR","BOL NO (LD#)","BOL No","BOLNO","BOL/SHIPMENT NO","BOL/SHIPMENT#","BOL/Order#:","BOL ID","DO/BOL","SID/BOL#:","Underlying BOLs",
                                               "BOL#","BOL","B/L Document No.","B/L NUMBER","B/L NO.","B/L #","B/L#","B L #:","B L#:","BL#","Bill Number","RETURN AUTHORIZATION","RGA#","RMA#","RA#","RGA","RMA","RA",
                                               "SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","CUSTOMER PURCHASE ORDER","CUSTOMERPURCHASEORDER#","PURCHASE-ORDER-NO","Purchase Order","PURCHASEORDER#","Customer PO Number",
                                               "Customer PO No","Cust P.O. No","CUSTOMER PO INFORMATION","SALES ORDER NBR/CUSTOMER P.O","Customer PO#","Customer POs:","Customer PO",
                                               "CustomerPO#","cust po#","PO/REFERENCE NO","P0/ORDER NUMBERS","PO Reference","DELIVERY/PO NUMBER","P O. NUMBER(S)","P.O. NUMBER(S)","P. O. NUMBER","PO Number",
                                               "ponumber","PO No.(s)","P O NO","P.O. No","ORDER ID / PO","P O #:","PO#(s)","P/O#","po#","Po","Customer Order Number","CUSTOMER ORDER NO(S)","Customer Order No","CUSTOMER ORDER NUM","CUSTOMERORDERNUMBER",
                                               "CUST. ORDER NO.","CUSTOMER ORDER N","CUSTOMER ORDER #","Cust order",
                                               "SHIPPER REF #/ORIG BOL #","SHIPPER REFERENCE NO","SHIPPER REF NUMBER","Shippers Release No","SHIPPERS REFERENCE","Shipper Ref","Shipper Number","SH PPER NUMBER","SHIPPER NUM","Shipper NOS",
                                               "Shipper No","SHIPPERSNO","shipperno","Shipper ID","SHIPPER #","SHIPPER#","Shipment Reference Number","Shipment Reference","SHIPMENT IDENTIFICATION NO","SHIPMENT NUMBER(S)","Shipment number",
                                               "Shipment No","shpmt num","Shipment ID","Shipment #","Shipment","Ship ID#:","Ship ID","SHIP NO","Ship Ref","Shipping Order No",
                                               "CHEMTREC","ChemTel","INFOTRAC","CHEM EMER #","Emergency Response Phone Number","Emergency Response Phone#","Emergency Response","Emergency Phone Numbers","EMERGENCY PHONE","EMERGENCY CONTACT CALL:",
                                               "Emergency Contact Phone","Emergency Contact","Emergency TEL","chemical Emergencies","Emergencies","Emergency"};

                var sKeyarray2 = new string[] {"SHIPPER BILL OF LADING NUMBER:","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING / BOOKING NUMBER","BILL OF LADING / PACKING SLIP","Bill of Lading Number",
                                               "Bil of Lading Number:","BILL OF LADING REF NO","Bill Of Lading Nos","ORDER/BILL OF LADING NO","Bill of Lading Num",
                                               "Bill of Lading No","BIL OF LADING NO","BILL OF LADING - ME","Bill of Lading #","Bill of Lading#","Bill of Lading","Pro/BOL Number",
                                               "BOL Number","BOLNUMBER","BOL NUM","BOL NBR","BOLNBR","BOL NO (LD#)","BOL No","BOLNO","BOL/SHIPMENT NO","BOL/SHIPMENT#","BOL/Order#:","BOL ID","DO/BOL","SID/BOL#:","Underlying BOLs",
                                               "BOL#","BOL","B/L Document No.","B/L NUMBER","B/L NO.","B/L #","B/L#","B L #:","B L#:","BL#","Bill Number","RETURN AUTHORIZATION","RGA#","RMA#","RA#","RGA","RMA","RA",
                                               "SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","CUSTOMER PURCHASE ORDER","CUSTOMERPURCHASEORDER#","PURCHASE-ORDER-NO","Purchase Order","PURCHASEORDER#","Customer PO Number",
                                               "Customer PO No","Cust P.O. No","CUSTOMER PO INFORMATION","SALES ORDER NBR/CUSTOMER P.O","Customer PO#","Customer POs:","Customer PO",
                                               "CustomerPO#","cust po#","PO/REFERENCE NO","P0/ORDER NUMBERS","PO Reference","DELIVERY/PO NUMBER","P O. NUMBER(S)","P.O. NUMBER(S)","P. O. NUMBER","PO Number",
                                               "ponumber","PO No.(s)","P O NO","P.O. No","ORDER ID / PO","P O #:","PO#(s)","P/O#","po#","Po","Customer Order Number","CUSTOMER ORDER NO(S)","Customer Order No","CUSTOMER ORDER NUM","CUSTOMERORDERNUMBER",
                                               "CUST. ORDER NO.","CUSTOMER ORDER N","CUSTOMER ORDER #","Cust order",
                                               "SHIPPER REF #/ORIG BOL #","SHIPPER REFERENCE NO","SHIPPER REF NUMBER","Shippers Release No","SHIPPERS REFERENCE","Shipper Ref","Shipper Number","SH PPER NUMBER","SHIPPER NUM","Shipper NOS",
                                               "Shipper No","SHIPPERSNO","shipperno","Shipper ID","SHIPPER #","SHIPPER#","Shipment Reference Number","Shipment Reference","SHIPMENT IDENTIFICATION NO","SHIPMENT NUMBER(S)","Shipment number",
                                               "Shipment No","shpmt num","Shipment ID","Shipment #","Shipment","Ship ID#:","Ship ID","SHIP NO","Ship Ref","Shipping Order No"};

                var sKeyarray = new string[sKeyarray1.Length];
                //var sKeyarray = new string[] { "Bill of Lading Number", "BOL", "B/L", "Manifest", "Bill of Lading No", "BOL No","Bill of L ading No", "BI LL 0 F LADING","BOL ID","Shipper's Bill of Ladias No","BOL Number",
                //                               "Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.","BOLNo","Po", "Customer Order Number",
                //                               "Customer Order No","Customer PO Number","Customer PO No","Customer's PO","cust po#","PO Number","Cust order","Purchase Ader","Customer PO","Consisnee's Refer.noe/po No",
                //                               "P.O. No","Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","Cust.P.O.#","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","CUST. ORDER NO.","Order","No/Customer PO#","PO No.(s)",
                //                               "CustomerPO","Order Number","Shipper Number",  "Ship ID","Shipper Ref", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No","Shipper No", "Shipment",
                //                               "Shipping Order No", "Shipper No/Customer","Shipment ID","Shipment No","Shipment number","SID","SHIPPER","Shipper Ref#","Shippers Release No","Pickup Number","TMS ID","ASN","Packing Slip Number","Ship Ref",
                //                               "Emergency","Emergency Phone Numbers","Emergency Response Phone","Emergency Contact","EMERGENCYCONTACTNUMBER:",
                //                               "24hrEMERGENCY PHONENUMBERS","CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone",
                //                               "ChemTel","CHEMTREC","INFOTRAC", };

                // Dim sKeyarray() As String = {"Your Contact"}
                // -----------------------------------------------------------------------------------------------------------------------------------
                // Dim sKeyarray() As String = {"Invoice No:"}

                //Array.Sort(sKeyarray);

                //DataTable dt = oMeta.DocDictionary;

                //var CNC = new ClsCNC();
                //// Dim Words As clsCnCWord() = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999)

                //var PossibleInvoices = new List<clsCnCWord>();
                ////var oRetF3 = new RetStructF3();
                //var oCnc = new ClsCNC();
                //var conflevel = new List<int>();
                //conflevel.Clear();


                //System.Data.DataRow[] foundRows;

                for (int iPage = intCurrPageNo; iPage <= oMeta.PageCount; iPage++)
                //for (int iPage = intCurrPageNo; iPage <= intCurrPageNo; iPage++)
                {
                    //int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == intCurrPageNo.ToString()).Max(row => row["Line_No"]));
                    //int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == iPage.ToString()).Max(row => row["Line_No"]));
                    //for (int iLine = 1; iLine <= max; iLine++)
                    for (int iLine = 1; iLine <= oMeta.Page[iPage].LineCount; iLine++)
                    {
                        //foundRows = dt.Select("[Page_no]=" + iPage + " AND [Line_no]=" + iLine);
                        clsCnCLine oLine = oMeta.Page[iPage].Line[iLine];
                        //int RowCount = foundRows.Length;
                        for (int iWord = 1; iWord <= oLine.WordCount; iWord++)
                        {
                            // int WordCount= Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == "1" && row["Line_No"].ToString() ==iLine.ToString()).Max(row => row["Word_No"]));
                            clsCnCWord cncWord = oLine.Word[iWord];
                            int WordNo = Convert.ToInt32(cncWord.WordNumber);
                            //int WordNo = Convert.ToInt32(foundRows[iWord]["Word_no"]);
                            // string Word = foundRows[iWord]["word_Otext"].ToString();
                            string Word = cncWord.strWord;
                            //int WordLength = Convert.ToInt32(foundRows[iWord]["Word_Length"]);
                            int WordLength = cncWord.strWord.Length;
                            //string WordType = foundRows[iWord]["Word_Type"].ToString();
                            int lengthDiff;
                            if (!string.IsNullOrEmpty(Word))
                            {
                                //if (WordType.ToUpper() != "N")
                                if (!Word.All(char.IsNumber))
                                {
                                    if (iPage > intCurrPageNo)
                                    {
                                        sKeyarray = sKeyarray2;
                                        //Array.Sort(sKeyarray);                                        
                                    }
                                    else
                                    {
                                        sKeyarray = sKeyarray1;
                                        //Array.Sort(sKeyarray);                                        
                                    }
                                    //Array.Sort(sKeyarray, (x, y) => x.Split().Length.CompareTo(y.Split().Length));
                                    //Array.Reverse(sKeyarray);
                                    for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
                                    {
                                        var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
                                        lengthDiff = WordLength - skeysplit[0].Length;
                                        if (skeysplit.Length > 1)
                                        {
                                            if (lengthDiff >= -2 && lengthDiff <= 2)
                                            {
                                                if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                                {
                                                    if (iWord > oLine.WordCount)
                                                    {
                                                    }
                                                    else
                                                    {
                                                        int iIndex;
                                                        var loopTo3 = skeysplit.Length - 1;
                                                        for (iIndex = 1; iIndex <= loopTo3; iIndex++)
                                                        {
                                                            if (iWord + iIndex > oLine.WordCount)
                                                                continue;
                                                            //lengthDiff = Convert.ToInt32(foundRows[iWord + iIndex]["Word_Length"]) - skeysplit[iIndex].Length;
                                                            clsCnCWord NxtWord = oLine.Word[iWord + iIndex];
                                                            //string oWord = foundRows[iWord + iIndex]["word_Otext"].ToString();
                                                            string oWord = NxtWord.strWord;
                                                            lengthDiff = oWord.Length - skeysplit[iIndex].Length;
                                                            if (oWord.Contains(":"))
                                                            {
                                                                string[] splitword;
                                                                if (oWord.Contains(":") && oWord.Any(char.IsDigit) && (oWord.Count(char.IsDigit) > 1))
                                                                {
                                                                    splitword = oWord.Split(':');
                                                                    oWord = splitword[0];
                                                                    lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
                                                                }
                                                                //else if(oWord.Contains("#"))
                                                                //{
                                                                //    splitword = oWord.Split('#');
                                                                //    oWord = splitword[0];
                                                                //    lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
                                                                //}
                                                            }

                                                            if (lengthDiff >= -2 && lengthDiff <= 2)
                                                            {
                                                                //if (CheckValidWord(foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                if (CheckValidWord(oWord.ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                {
                                                                    if (skeysplit.Length - 1 != iIndex)
                                                                    {
                                                                        continue;
                                                                    }
                                                                    else
                                                                    {
                                                                        DataRow dr;

                                                                        dr = dtKeywordInfo.NewRow();
                                                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                        dr["Page No"] = cncWord.PageNo;
                                                                        dr["Line No"] = cncWord.LineNo;
                                                                        dr["Word No"] = WordNo;
                                                                        dr["LineWordNo"] = WordNo;
                                                                        dr["X1"] = cncWord.Left;
                                                                        dr["Y1"] = cncWord.Top;
                                                                        dr["X2"] = cncWord.Right;
                                                                        dr["Y2"] = cncWord.Bottom;
                                                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                        dtKeywordInfo.Rows.Add(dr);
                                                                        iWord = iWord + loopTo3;
                                                                        jLoop = sKeyarray.Length - 1;
                                                                        break;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    break;
                                                                }
                                                            }
                                                            else if (iIndex == loopTo3)
                                                            {
                                                                //string word = foundRows[iWord + iIndex]["word_Otext"].ToString();
                                                                if (oWord.ToUpper().StartsWith(skeysplit[iIndex].ToUpper()))
                                                                {
                                                                    //string keyword = Word.Substring(0, skeysplit[0].Length);
                                                                    //if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
                                                                    //{
                                                                    //    if (!Module1.RAKeywords.Contains(keyword) || keyword.ToUpper() == "RA")
                                                                    //    {
                                                                    //        continue;
                                                                    //    }
                                                                    //}
                                                                    if (oWord.ToUpper().Any(char.IsDigit))
                                                                    {
                                                                        DataRow dr;

                                                                        dr = dtKeywordInfo.NewRow();
                                                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                        dr["Page No"] = cncWord.PageNo;
                                                                        dr["Line No"] = cncWord.LineNo;
                                                                        dr["Word No"] = WordNo;
                                                                        dr["LineWordNo"] = WordNo;
                                                                        dr["X1"] = cncWord.Left;
                                                                        dr["Y1"] = cncWord.Top;
                                                                        dr["X2"] = cncWord.Right;
                                                                        dr["Y2"] = cncWord.Bottom;
                                                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                        dtKeywordInfo.Rows.Add(dr);
                                                                        iWord = iWord + loopTo3;
                                                                        jLoop = sKeyarray.Length - 1;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            { break; }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        //else if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                        else if (lengthDiff >= -2 && lengthDiff <= 2)
                                        {
                                            //if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                            if (CheckValidWord(Word.Trim(), new string[] { skeysplit[0] }) == true)
                                            {
                                                DataRow dr;

                                                dr = dtKeywordInfo.NewRow();
                                                dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                dr["Page No"] = cncWord.PageNo;
                                                dr["Line No"] = cncWord.LineNo;
                                                dr["Word No"] = WordNo;
                                                dr["LineWordNo"] = WordNo;
                                                dr["X1"] = cncWord.Left;
                                                dr["Y1"] = cncWord.Top;
                                                dr["X2"] = cncWord.Right;
                                                dr["Y2"] = cncWord.Bottom;
                                                dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                dtKeywordInfo.Rows.Add(dr);
                                                jLoop = sKeyarray.Length - 1;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (Word.ToUpper().StartsWith(skeysplit[0].ToUpper()))
                                            {
                                                string[] StartsPO = { "PO:", "PO#", "PO-" };
                                                string keyword = Word.Substring(0, skeysplit[0].Length);
                                                if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
                                                {
                                                    if (!Module1.RAKeywords.Contains(keyword) || keyword.ToUpper() == "RA")
                                                    {
                                                        continue;
                                                    }
                                                }
                                                if (keyword.ToUpper() == "PO" && (!StartsPO.Any(Word.ToUpper().StartsWith)))
                                                {
                                                    continue;
                                                }
                                                if (Word.Any(char.IsDigit))
                                                {
                                                    DataRow dr;

                                                    dr = dtKeywordInfo.NewRow();
                                                    dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                    dr["Page No"] = cncWord.PageNo;
                                                    dr["Line No"] = cncWord.LineNo;
                                                    dr["Word No"] = WordNo;
                                                    dr["LineWordNo"] = WordNo;
                                                    dr["X1"] = cncWord.Left;
                                                    dr["Y1"] = cncWord.Top;
                                                    dr["X2"] = cncWord.Right;
                                                    dr["Y2"] = cncWord.Bottom;
                                                    dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                    dtKeywordInfo.Rows.Add(dr);
                                                    jLoop = sKeyarray.Length - 1;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                // throw;
                //MessageBox.Show(ex.Message, "SearchKeywords");
            }

            return dtKeywordInfo;
        }

        //public static DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        //{
        //    var dtKeywordInfo = new DataTable();
        //    dtKeywordInfo = MakeDt();
        //    try
        //    {
        //        // -----------------------------------------BILL OF LADING NUMBER/PURCHASE ORDER NUMBER/SHIPPER NUMBER KEYWORD------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "lhre Bestellung per FAX", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero"}

        //        var sKeyarray1 = new string[] { "Bill of Lading Number", "BOL#","BOL", "B/L #","Bill of Lading No", "BOL No","Bill of L ading No", "BOL ID","Bill of Lading #",
        //                                       "BOL Number","Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","Po", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No","PO#(s)",
        //                                       "Customers PO","po#","cust po#","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","Customer PO#","Consisnee's Refer.noe/po No","P.O. No",
        //                                       "Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","P/O","CUST. ORDER NO.","PO Reference", "Shipper Number",  "Ship ID", "Shipper Ref",
        //                                       "Shipper No", "Shipment",  "Shipping Order No",  "Shipment ID","Shipment No","Shipment number","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.",
        //                                       "Emergency", "Emergency Phone Numbers","Emergency Response Phone#","Emergency Contact","Emergency Response Phone Number","EMERGENCYCONTACTNUMBER:","24hrEMERGENCY PHONENUMBERS","Shipment Numbers",
        //                                       "CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone","Emergencies","ChemTel","Chem tel",
        //                                       "CHEMTREC","INFOTRAC","BOLNO","Shippers Release No","Ship Ref","Cust order","Load","Load ID","Load Number","Load No","SHIPPER NUM","shpmt num",
        //                                       "CUSTOMER ORDER NUM","RA","RMA","RMA#","RGA","RGA#","RA#","SHIPPER REF NUMBER","Shipment Reference Number","Shipper NOS","RETURN AUTHORIZATION","BOL/SHIPMENT NO","BE OF LADING NO","BIL OF LADING NO","Bill Number",
        //                                       "ponumber","shipperno","Customer POs:","Bill of Lading#","B/L#","Ship ID#:","B0L#","B L#:","P O #:","P/O#","P O. NUMBER(S)","PO NUM BER","BOL/Order#:","DO/BOL","SH PPER NUMBER","PO NUMBE R:","CUSTOMER ORDER #","SHIPPER#","SALES ORDER NBR/CUSTOMER P.O",
        //                                       "Bil of Lading Number:","BIE OF LADING NO","LOAD#","BILL OF LADING - ME","SHIPMEN T NUMBER","SHIPMEN T NUMBE R","BILLOF LADING NUMBER","CUSTOMER PO INFORMATION","B0 L #","BILL OF LADING / BOOKING NUMBER","Shipper ID","BOLNUMBER",
        //                                       "BILLOFLADING","P O NO","SID/BOL#:","BOLNBR","SHIPPERNOS","Shippers #","SHIPPER #","BOL NO (LD#)","CUSTOMER ORDER NO(S)","ORDER ID / PO","B L #:","BOL/SHIPMENT#","BL#","SHIPPER REF #/ORIG BOL #","BILL OF LADING / PACKING SLIP","SHIP NO","SHIPMENT IDENTIFICATION NO",
        //                                       "P0/ORDER NUMBERS","SHIPPERS REFERENCE","DELIVERY/PO NUMBER","BDL OF LADING NO","CUSTOMER ORDER N","CUSTOME R P.O.","SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","BILL OF LADINGNUMBER","PO/REFERENCE NO","P0 #","CUSTOMER P0",
        //                                        "ORDER/BILL OF LADING NO","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING REF NO","SHIPPER REFERENCE NO","BHI OF LADING #","BILL OFLADING NO","CUSTOMERORDER NUM","BDL NO","SHIPPERSNO","Shipper Reference","Shipment #","Underlying BOLs","CustomerPO#",
        //                                        "Bill Of Lading Nos","BOL NUM","P. O. NUMBER","BM OF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BILL OF LAD ING NUMBER"};

        //        var sKeyarray2 = new string[] { "Bill of Lading Number", "BOL#","BOL", "B/L #","Bill of Lading No", "BOL No","Bill of L ading No", "BOL ID","Bill of Lading #",
        //                                       "BOL Number","Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","Po", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No","PO#(s)",
        //                                       "Customers PO","po#","cust po#","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","Customer PO#","Consisnee's Refer.noe/po No","P.O. No",
        //                                       "Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","P/O","CUST. ORDER NO.","PO Reference", "Shipper Number",  "Ship ID", "Shipper Ref",
        //                                       "Shipper No", "Shipment",  "Shipping Order No",  "Shipment ID","Shipment No","Shipment number","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.",
        //                                       "BOLNO","Shippers Release No","Ship Ref","Cust order","SHIPPER NUM","shpmt num","CUSTOMER ORDER NUM","RA","RMA","RMA#","RGA","RGA#","RA#","SHIPPER REF NUMBER","Shipment Reference Number","Shipper NOS","RETURN AUTHORIZATION","BOL/SHIPMENT NO","BE OF LADING NO","BIL OF LADING NO","Bill Number",
        //                                       "ponumber","shipperno","Customer POs:","Bill of Lading#","B/L#","Ship ID#:","B0L#","B L#:","P O #:","P/O#","P O. NUMBER(S)","PO NUM BER","BOL/Order#:","DO/BOL","SH PPER NUMBER","PO NUMBE R:","CUSTOMER ORDER #","SHIPPER#","SALES ORDER NBR/CUSTOMER P.O",
        //                                       "Bil of Lading Number:","BIE OF LADING NO","BILL OF LADING - ME","SHIPMEN T NUMBER","SHIPMEN T NUMBE R","BILLOF LADING NUMBER","CUSTOMER PO INFORMATION","B0 L #","BILL OF LADING / BOOKING NUMBER","Shipper ID","BOLNUMBER",
        //                                       "BILLOFLADING","P O NO","SID/BOL#:","BOLNBR","SHIPPERNOS","Shippers #","SHIPPER #","BOL NO (LD#)","CUSTOMER ORDER NO(S)","ORDER ID / PO","B L #:","BOL/SHIPMENT#","BL#","SHIPPER REF #/ORIG BOL #","BILL OF LADING / PACKING SLIP","SHIP NO","SHIPMENT IDENTIFICATION NO",
        //                                       "P0/ORDER NUMBERS","SHIPPERS REFERENCE","DELIVERY/PO NUMBER","BDL OF LADING NO","CUSTOMER ORDER N","CUSTOME R P.O.","SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","BILL OF LADINGNUMBER","PO/REFERENCE NO","P0 #","CUSTOMER P0",
        //                                        "ORDER/BILL OF LADING NO","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING REF NO","SHIPPER REFERENCE NO","BHI OF LADING #","BILL OFLADING NO","CUSTOMERORDER NUM","BDL NO","SHIPPERSNO","Shipper Reference","Shipment #","Underlying BOLs","CustomerPO#",
        //                                        "Bill Of Lading Nos","BOL NUM","P. O. NUMBER","BM OF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BILL OF LAD ING NUMBER"};

        //        //var sKeyarray1 = new string[] {"SHIPPER BILL OF LADING NUMBER:","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING / BOOKING NUMBER","BILL OF LADING / PACKING SLIP","Bill of Lading Number","Bill of LadingNumber","BILL OF LADINGNUMBER",
        //        //                               "Bil of Lading Number:","BILL OF LAD ING NUMBER","BILLOF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BM OF LADING NUMBER","BILL OF LADING REF NO","Bill Of Lading Nos","ORDER/BILL OF LADING NO",
        //        //                               "Bill of Lading No","Bill of L ading No","BILL OFLADING NO","BIL OF LADING NO","BIE OF LADING NO","BDL OF LADING NO","BE OF LADING NO","BILL OF LADING - ME","Bill of Lading #","BHI OF LADING #","Bill of Lading#","Bill of Lading",
        //        //                               "BILLOFLADING","BOL Number","BOLNUMBER","BOL NUM","BOL NBR","BOLNBR","BOL NO (LD#)","BOL No","BOLNO","BDL NO","BOL/SHIPMENT NO","BOL/SHIPMENT#","BOL/Order#:","BOL ID","DO/BOL","SID/BOL#:","Underlying BOLs",
        //        //                               "BOL#","B0 L #","B0L#","BOL","B/L Document No.","B/L NUMBER","B/L NO.","B/L #","B/L#","B L #:","B L#:","BL#","Bill Number","RETURN AUTHORIZATION","RGA#","RMA#","RA#","RGA","RMA","RA",
        //        //                               "SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","CUSTOMER PURCHASE ORDER","PURCHASE-ORDER-NO","Purchase Order","Purchase Ader","Consisnee's Refer.noe/po No","customer p.o. number","Customer PO Number",
        //        //                               "customer p.o. no","Customer P.O. No","Customer PO No","Cust P.O. No","CUSTOMER PO INFORMATION","SALES ORDER NBR/CUSTOMER P.O","Customer PO#","Customer POs:","Customers PO","Customer PO","CUSTOME R P.O.",
        //        //                               "CUSTOMER P0","CustomerPO#","cust po#","PO/REFERENCE NO","P0/ORDER NUMBERS","PO Reference","DELIVERY/PO NUMBER","P O. NUMBER(S)","P.O. NUMBER(S)","P. O. NUMBER","PO Number","PO NUM BER","PO NUMBE R:",
        //        //                               "ponumber","PO No.(s)","P O NO","P.O. No","ORDER ID / PO","P O #:","P0 #","PO#(s)","P/O#","po#","P/O","Po","Customer Order Number","CUSTOMER ORDER NO(S)","Customer Order No","CUSTOMER ORDER NUM","CUSTOMERORDERNUMBER",
        //        //                               "CUSTOMERORDER NUM","CUST. ORDER NO.","CUSTOMER ORDER N","CUSTOMER ORDER #","Cust order",
        //        //                               "SHIPPER REF #/ORIG BOL #","SHIPPER REFERENCE NO","SHIPPER REF NUMBER","Shippers Release No","SHIPPERS REFERENCE","Shipper Reference","Shipper Ref","Shipper Number","SH PPER NUMBER","SHIPPER NUM","Shipper NOS",
        //        //                               "Shipper No","SHIPPERNOS","SHIPPERSNO","shipperno","Shipper ID","Shippers #","SHIPPER #","SHIPPER#","Shipment Reference Number","SHIPMENT IDENTIFICATION NO","Shipment Numbers","Shipment number","SHIPMEN T NUMBE R",
        //        //                               "SHIPMEN T NUMBER","Shipment No","shpmt num","Shipment ID","Shipment #","Shipment","Ship ID#:","Ship ID","SHIP NO","Ship Ref","Shipping Order No",
        //        //                               "CHEMTREC","ChemTel","Chem tel","INFOTRAC","CHEM EMER #","Emergency Response Phone Number","Emergency Response Phone#","Emergency Response","Emergency Phone Numbers","EMERGENCY PHONE","EMERGENCY CONTACT CALL:",
        //        //                               "Emergency Contact Phone","EMERGENCYCONTACTNUMBER:","Emergency Contact","Emergency TEL","24hrEMERGENCY PHONENUMBERS","chemical Emergencies","Emergencies","Emergency",
        //        //                               "Load Number","Load No","Load ID","LOAD#","Load"};

        //        //var sKeyarray2 = new string[] {"SHIPPER BILL OF LADING NUMBER:","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING / BOOKING NUMBER","BILL OF LADING / PACKING SLIP","Bill of Lading Number","Bill of LadingNumber","BILL OF LADINGNUMBER",
        //        //                               "Bil of Lading Number:","BILL OF LAD ING NUMBER","BILLOF LADING NUMBER","BIOF LADING NUMBER","BLOF LADING NUMBER","BM OF LADING NUMBER","BILL OF LADING REF NO","Bill Of Lading Nos","ORDER/BILL OF LADING NO",
        //        //                               "Bill of Lading No","Bill of L ading No","BILL OFLADING NO","BIL OF LADING NO","BIE OF LADING NO","BDL OF LADING NO","BE OF LADING NO","BILL OF LADING - ME","Bill of Lading #","BHI OF LADING #","Bill of Lading#","Bill of Lading",
        //        //                               "BILLOFLADING","BOL Number","BOLNUMBER","BOL NUM","BOL NBR","BOLNBR","BOL NO (LD#)","BOL No","BOLNO","BDL NO","BOL/SHIPMENT NO","BOL/SHIPMENT#","BOL/Order#:","BOL ID","DO/BOL","SID/BOL#:","Underlying BOLs",
        //        //                               "BOL#","B0 L #","B0L#","BOL","B/L Document No.","B/L NUMBER","B/L NO.","B/L #","B/L#","B L #:","B L#:","BL#","Bill Number","RETURN AUTHORIZATION","RGA#","RMA#","RA#","RGA","RMA","RA",
        //        //                               "SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","CUSTOMER PURCHASE ORDER","PURCHASE-ORDER-NO","Purchase Order","Purchase Ader","Consisnee's Refer.noe/po No","customer p.o. number","Customer PO Number",
        //        //                               "customer p.o. no","Customer P.O. No","Customer PO No","Cust P.O. No","CUSTOMER PO INFORMATION","SALES ORDER NBR/CUSTOMER P.O","Customer PO#","Customer POs:","Customers PO","Customer PO","CUSTOME R P.O.",
        //        //                               "CUSTOMER P0","CustomerPO#","cust po#","PO/REFERENCE NO","P0/ORDER NUMBERS","PO Reference","DELIVERY/PO NUMBER","P O. NUMBER(S)","P.O. NUMBER(S)","P. O. NUMBER","PO Number","PO NUM BER","PO NUMBE R:",
        //        //                               "ponumber","PO No.(s)","P O NO","P.O. No","ORDER ID / PO","P O #:","P0 #","PO#(s)","P/O#","po#","P/O","Po","Customer Order Number","CUSTOMER ORDER NO(S)","Customer Order No","CUSTOMER ORDER NUM","CUSTOMERORDERNUMBER",
        //        //                               "CUSTOMERORDER NUM","CUST. ORDER NO.","CUSTOMER ORDER N","CUSTOMER ORDER #","Cust order",
        //        //                               "SHIPPER REF #/ORIG BOL #","SHIPPER REFERENCE NO","SHIPPER REF NUMBER","Shippers Release No","SHIPPERS REFERENCE","Shipper Reference","Shipper Ref","Shipper Number","SH PPER NUMBER","SHIPPER NUM","Shipper NOS",
        //        //                               "Shipper No","SHIPPERNOS","SHIPPERSNO","shipperno","Shipper ID","Shippers #","SHIPPER #","SHIPPER#","Shipment Reference Number","SHIPMENT IDENTIFICATION NO","Shipment Numbers","Shipment number","SHIPMEN T NUMBE R",
        //        //                               "SHIPMEN T NUMBER","Shipment No","shpmt num","Shipment ID","Shipment #","Shipment","Ship ID#:","Ship ID","SHIP NO","Ship Ref","Shipping Order No"};

        //        var sKeyarray = new string[sKeyarray1.Length];
        //        //var sKeyarray = new string[] { "Bill of Lading Number", "BOL", "B/L", "Manifest", "Bill of Lading No", "BOL No","Bill of L ading No", "BI LL 0 F LADING","BOL ID","Shipper's Bill of Ladias No","BOL Number",
        //        //                               "Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.","BOLNo","Po", "Customer Order Number",
        //        //                               "Customer Order No","Customer PO Number","Customer PO No","Customer's PO","cust po#","PO Number","Cust order","Purchase Ader","Customer PO","Consisnee's Refer.noe/po No",
        //        //                               "P.O. No","Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","Cust.P.O.#","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","CUST. ORDER NO.","Order","No/Customer PO#","PO No.(s)",
        //        //                               "CustomerPO","Order Number","Shipper Number",  "Ship ID","Shipper Ref", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No","Shipper No", "Shipment",
        //        //                               "Shipping Order No", "Shipper No/Customer","Shipment ID","Shipment No","Shipment number","SID","SHIPPER","Shipper Ref#","Shippers Release No","Pickup Number","TMS ID","ASN","Packing Slip Number","Ship Ref",
        //        //                               "Emergency","Emergency Phone Numbers","Emergency Response Phone","Emergency Contact","EMERGENCYCONTACTNUMBER:",
        //        //                               "24hrEMERGENCY PHONENUMBERS","CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone",
        //        //                               "ChemTel","CHEMTREC","INFOTRAC", };

        //        // Dim sKeyarray() As String = {"Your Contact"}
        //        // -----------------------------------------------------------------------------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice No:"}

        //        //Array.Sort(sKeyarray);

        //        DataTable dt = oMeta.DocDictionary;

        //        //var CNC = new ClsCNC();
        //        //// Dim Words As clsCnCWord() = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999)

        //        //var PossibleInvoices = new List<clsCnCWord>();
        //        ////var oRetF3 = new RetStructF3();
        //        //var oCnc = new ClsCNC();
        //        //var conflevel = new List<int>();
        //        //conflevel.Clear();


        //        System.Data.DataRow[] foundRows;

        //        for (int iPage = intCurrPageNo; iPage <= oMeta.PageCount; iPage++)
        //        //for (int iPage = intCurrPageNo; iPage <= intCurrPageNo; iPage++)
        //        {
        //            //int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == intCurrPageNo.ToString()).Max(row => row["Line_No"]));
        //            int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == iPage.ToString()).Max(row => row["Line_No"]));
        //            for (int iLine = 1; iLine <= max; iLine++)
        //            {
        //                foundRows = dt.Select("[Page_no]=" + iPage + " AND [Line_no]=" + iLine);
        //                //foreach (DataRow drDataRow in foundRows)
        //                int RowCount = foundRows.Length;
        //                for (int iWord = 0; iWord < RowCount; iWord++)
        //                {
        //                    // int WordCount= Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == "1" && row["Line_No"].ToString() ==iLine.ToString()).Max(row => row["Word_No"]));

        //                    int WordNo = Convert.ToInt32(foundRows[iWord]["Word_no"]);
        //                    string Word = foundRows[iWord]["word_Otext"].ToString();
        //                    int WordLength = Convert.ToInt32(foundRows[iWord]["Word_Length"]);
        //                    string WordType = foundRows[iWord]["Word_Type"].ToString();
        //                    int lengthDiff;
        //                    if (!string.IsNullOrEmpty(WordType))
        //                    {
        //                        if (WordType.ToUpper() != "N")
        //                        {
        //                            if(iPage>intCurrPageNo)
        //                            {
        //                                sKeyarray = sKeyarray2;
        //                                Array.Sort(sKeyarray);                                        
        //                            }
        //                            else
        //                            {
        //                                sKeyarray = sKeyarray1;
        //                                Array.Sort(sKeyarray);                                        
        //                            }
        //                            //Array.Sort(sKeyarray, (x, y) => x.Split().Length.CompareTo(y.Split().Length));
        //                            //Array.Reverse(sKeyarray);
        //                            for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
        //                            {
        //                                var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
        //                                lengthDiff = WordLength - skeysplit[0].Length;
        //                                if (skeysplit.Length > 1)
        //                                {
        //                                    if (lengthDiff >= -2 && lengthDiff <= 2)
        //                                    {
        //                                        if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                                        {
        //                                            if (iWord >= RowCount)
        //                                            {
        //                                            }
        //                                            else
        //                                            {
        //                                                int iIndex;
        //                                                var loopTo3 = skeysplit.Length - 1;
        //                                                for (iIndex = 1; iIndex <= loopTo3; iIndex++)
        //                                                {
        //                                                    if (iWord + iIndex >= RowCount)
        //                                                        continue;
        //                                                    lengthDiff = Convert.ToInt32(foundRows[iWord + iIndex]["Word_Length"]) - skeysplit[iIndex].Length;
        //                                                    string oWord = foundRows[iWord + iIndex]["word_Otext"].ToString();
        //                                                    if (oWord.Contains(":"))
        //                                                    {
        //                                                        string[] splitword;
        //                                                        if (oWord.Contains(":") && oWord.Any(char.IsDigit) && (oWord.Count(char.IsDigit) > 1))
        //                                                        {
        //                                                            splitword = oWord.Split(':');
        //                                                            oWord = splitword[0];
        //                                                            lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
        //                                                        }
        //                                                        //else if(oWord.Contains("#"))
        //                                                        //{
        //                                                        //    splitword = oWord.Split('#');
        //                                                        //    oWord = splitword[0];
        //                                                        //    lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
        //                                                        //}
        //                                                    }

        //                                                    if (lengthDiff >= -2 && lengthDiff <= 2)
        //                                                    {
        //                                                        //if (CheckValidWord(foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
        //                                                        if (CheckValidWord(oWord.ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
        //                                                        {
        //                                                            if (skeysplit.Length - 1 != iIndex)
        //                                                            {
        //                                                                continue;
        //                                                            }
        //                                                            else
        //                                                            {
        //                                                                DataRow dr;

        //                                                                dr = dtKeywordInfo.NewRow();
        //                                                                dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                                dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                                                dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                                                dr["Word No"] = WordNo;
        //                                                                dr["LineWordNo"] = WordNo;
        //                                                                dr["X1"] = foundRows[iWord]["X1"];
        //                                                                dr["Y1"] = foundRows[iWord]["Y1"];
        //                                                                dr["X2"] = foundRows[iWord]["X2"];
        //                                                                dr["Y2"] = foundRows[iWord]["Y2"];
        //                                                                dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                                                dtKeywordInfo.Rows.Add(dr);
        //                                                                //iWord = iWord + loopTo3;
        //                                                                //jLoop = sKeyarray.Length - 1;
        //                                                                //break;
        //                                                            }
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            break;
        //                                                        }
        //                                                    }
        //                                                    else if (iIndex == loopTo3)
        //                                                    {
        //                                                        string word = foundRows[iWord + iIndex]["word_Otext"].ToString();
        //                                                        if (foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().StartsWith(skeysplit[iIndex].ToUpper()))
        //                                                        {
        //                                                            //string keyword = Word.Substring(0, skeysplit[0].Length);
        //                                                            //if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
        //                                                            //{
        //                                                            //    if (!Module1.RAKeywords.Contains(keyword) || keyword.ToUpper() == "RA")
        //                                                            //    {
        //                                                            //        continue;
        //                                                            //    }
        //                                                            //}
        //                                                            if (foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Any(char.IsDigit))
        //                                                            {
        //                                                                DataRow dr;

        //                                                                dr = dtKeywordInfo.NewRow();
        //                                                                dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                                dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                                                dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                                                dr["Word No"] = WordNo;
        //                                                                dr["LineWordNo"] = WordNo;
        //                                                                dr["X1"] = foundRows[iWord]["X1"];
        //                                                                dr["Y1"] = foundRows[iWord]["Y1"];
        //                                                                dr["X2"] = foundRows[iWord]["X2"];
        //                                                                dr["Y2"] = foundRows[iWord]["Y2"];
        //                                                                dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                                                dtKeywordInfo.Rows.Add(dr);
        //                                                                //iWord = iWord + loopTo3;
        //                                                                //jLoop = sKeyarray.Length - 1;
        //                                                                //break;
        //                                                            }
        //                                                        }
        //                                                    }
        //                                                    else
        //                                                    { break; }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                                //else if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                                else if (lengthDiff >= -2 && lengthDiff <= 2)
        //                                {
        //                                    //if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                                    if (CheckValidWord(Word.Trim(), new string[] { skeysplit[0] }) == true)
        //                                    {
        //                                        DataRow dr;

        //                                        dr = dtKeywordInfo.NewRow();
        //                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                        dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                        dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                        dr["Word No"] = WordNo;
        //                                        dr["LineWordNo"] = WordNo;
        //                                        dr["X1"] = foundRows[iWord]["X1"];
        //                                        dr["Y1"] = foundRows[iWord]["Y1"];
        //                                        dr["X2"] = foundRows[iWord]["X2"];
        //                                        dr["Y2"] = foundRows[iWord]["Y2"];
        //                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                        dtKeywordInfo.Rows.Add(dr);
        //                                        //jLoop = sKeyarray.Length - 1;
        //                                        //break;
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    if (Word.ToUpper().StartsWith(skeysplit[0].ToUpper()))
        //                                    {
        //                                        string[] StartsPO = { "PO:", "PO#","PO-" };
        //                                        string keyword = Word.Substring(0, skeysplit[0].Length);
        //                                        if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
        //                                        {
        //                                            if (!Module1.RAKeywords.Contains(keyword) || keyword.ToUpper() == "RA")
        //                                            {
        //                                                continue;
        //                                            }
        //                                        }
        //                                        if (keyword.ToUpper() == "PO" && (!StartsPO.Any(Word.ToUpper().StartsWith)))
        //                                        {
        //                                            continue;
        //                                        }
        //                                        if (Word.Any(char.IsDigit))
        //                                        {
        //                                            DataRow dr;

        //                                            dr = dtKeywordInfo.NewRow();
        //                                            dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                            dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                            dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                            dr["Word No"] = WordNo;
        //                                            dr["LineWordNo"] = WordNo;
        //                                            dr["X1"] = foundRows[iWord]["X1"];
        //                                            dr["Y1"] = foundRows[iWord]["Y1"];
        //                                            dr["X2"] = foundRows[iWord]["X2"];
        //                                            dr["Y2"] = foundRows[iWord]["Y2"];
        //                                            dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                            dtKeywordInfo.Rows.Add(dr);
        //                                            //jLoop = sKeyarray.Length - 1;
        //                                            //break;
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }

        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        // throw;
        //        //MessageBox.Show(ex.Message, "SearchKeywords");
        //    }

        //    return dtKeywordInfo;
        //}
        #endregion


        #region F3searchkeyword for Single Page
        //public static DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        //{
        //    var dtKeywordInfo = new DataTable();
        //    dtKeywordInfo = MakeDt();
        //    try
        //    {
        //        // -----------------------------------------BILL OF LADING NUMBER/PURCHASE ORDER NUMBER/SHIPPER NUMBER KEYWORD------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "lhre Bestellung per FAX", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero"}

        //        var sKeyarray = new string[] { "Bill of Lading Number", "BOL#","BOL", "B/L #","Bill of Lading No", "BOL No","Bill of L ading No", "BOL ID","HBL","MBL","Bill of Lading #",
        //                                       "BOL Number","Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","Po", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No","PO#(s)",
        //                                       "Customers PO","po#","cust po#","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","Customer PO#","Consisnee's Refer.noe/po No","P.O. No",
        //                                       "Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","P/O","CUST. ORDER NO.","PO Reference", "Shipper Number",  "Ship ID", "Shipper Ref",
        //                                       "Shipper No", "Shipment",  "Shipping Order No",  "Shipment ID","Shipment No","Shipment number","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.",
        //                                       "Emergency", "Emergency Phone Numbers","Emergency Response Phone#","Emergency Contact","Emergency Response Phone Number","EMERGENCYCONTACTNUMBER:","24hrEMERGENCY PHONENUMBERS","Shipment Numbers",
        //                                       "CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone","Emergencies","ChemTel","Chem tel",
        //                                       "CHEMTREC","INFOTRAC","BOLNO","Shippers Release No","Ship Ref","Cust order","Load","Load ID","Load Number","Load No","SHIPPER NUM","shpmt num",
        //                                       "CUSTOMER ORDER NUM","RA","RMA","RGA","RGA#","RA#","SHIPPER REF NUMBER","Shipment Reference Number","Shipper NOS","RETURN AUTHORIZATION","BOL/SHIPMENT NO","BE OF LADING NO","BIL OF LADING NO","Bill Number",
        //                                       "ponumber","shipperno","Customer POs:","Bill of Lading#","B/L#","Ship ID#:","B0L#","B L#:","P O #:","P/O#","P O. NUMBER(S)","PO NUM BER","BOL/Order#:","DO/BOL","SH PPER NUMBER","PO NUMBE R:","CUSTOMER ORDER #","SHIPPER#","SALES ORDER NBR/CUSTOMER P.O",
        //                                       "Bil of Lading Number:","BIE OF LADING NO","LOAD#","BILL OF LADING - ME","SHIPMEN T NUMBER","SHIPMEN T NUMBE R","BILLOF LADING NUMBER","CUSTOMER PO INFORMATION","B0 L #","BILL OF LADING / BOOKING NUMBER","Shipper ID","BOLNUMBER",
        //                                       "BILLOFLADING","P O NO","SID/BOL#:","BOLNBR","SHIPPERNOS","Shippers #","SHIPPER #","BOL NO (LD#)","CUSTOMER ORDER NO(S)","ORDER ID / PO","B L #:","BOL/SHIPMENT#","BL#","SHIPPER REF #/ORIG BOL #","BILL OF LADING / PACKING SLIP","SHIP NO","SHIPMENT IDENTIFICATION NO",
        //                                       "P0/ORDER NUMBERS","SHIPPERS REFERENCE","DELIVERY/PO NUMBER","BDL OF LADING NO","CUSTOMER ORDER N","CUSTOME R P.O.","SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS","BILL OF LADINGNUMBER","PO/REFERENCE NO","P0 #","CUSTOMER P0",
        //                                        "ORDER/BILL OF LADING NO","BILL OF LADING NUMBER/SHIPMENT NUMBER","BILL OF LADING REF NO","SHIPPER REFERENCE NO","BHI OF LADING #","BILL OFLADING NO","CUSTOMERORDER NUM","BDL NO"};

        //        //var sKeyarray = new string[] { "Bill of Lading Number", "BOL", "B/L", "Manifest", "Bill of Lading No", "BOL No","Bill of L ading No", "BI LL 0 F LADING","BOL ID","Shipper's Bill of Ladias No","BOL Number",
        //        //                               "Bill of Lading","Bill of LadingNumber","B/L NO.","B/L NUMBER","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.","BOLNo","Po", "Customer Order Number",
        //        //                               "Customer Order No","Customer PO Number","Customer PO No","Customer's PO","cust po#","PO Number","Cust order","Purchase Ader","Customer PO","Consisnee's Refer.noe/po No",
        //        //                               "P.O. No","Purchase Order","Cust P.O. No","CUSTOMERORDERNUMBER","P.O. NUMBER(S)","Cust.P.O.#","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER","CUST. ORDER NO.","Order","No/Customer PO#","PO No.(s)",
        //        //                               "CustomerPO","Order Number","Shipper Number",  "Ship ID","Shipper Ref", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No","Shipper No", "Shipment",
        //        //                               "Shipping Order No", "Shipper No/Customer","Shipment ID","Shipment No","Shipment number","SID","SHIPPER","Shipper Ref#","Shippers Release No","Pickup Number","TMS ID","ASN","Packing Slip Number","Ship Ref",
        //        //                               "Emergency","Emergency Phone Numbers","Emergency Response Phone","Emergency Contact","EMERGENCYCONTACTNUMBER:",
        //        //                               "24hrEMERGENCY PHONENUMBERS","CHEM EMER #","EMERGENCY CONTACT CALL:","EMERGENCY PHONE","chemical Emergencies","Emergency Response","Emergency TEL","Emergency Contact Phone",
        //        //                               "ChemTel","CHEMTREC","INFOTRAC", };

        //        // Dim sKeyarray() As String = {"Your Contact"}
        //        // -----------------------------------------------------------------------------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice No:"}

        //        Array.Sort(sKeyarray);

        //        DataTable dt = oMeta.DocDictionary;

        //        //var CNC = new ClsCNC();
        //        //// Dim Words As clsCnCWord() = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999)

        //        //var PossibleInvoices = new List<clsCnCWord>();
        //        ////var oRetF3 = new RetStructF3();
        //        //var oCnc = new ClsCNC();
        //        //var conflevel = new List<int>();
        //        //conflevel.Clear();


        //        System.Data.DataRow[] foundRows;


        //        int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == intCurrPageNo.ToString()).Max(row => row["Line_No"]));

        //        for (int iLine = 1; iLine <= max; iLine++)
        //        {
        //            foundRows = dt.Select("[Page_no]="+intCurrPageNo+" AND [Line_no]=" + iLine);
        //            //foreach (DataRow drDataRow in foundRows)
        //            int RowCount = foundRows.Length;
        //            for (int iWord = 0; iWord < RowCount; iWord++)
        //            {
        //                // int WordCount= Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == "1" && row["Line_No"].ToString() ==iLine.ToString()).Max(row => row["Word_No"]));

        //                int WordNo = Convert.ToInt32(foundRows[iWord]["Word_no"]);
        //                string Word = foundRows[iWord]["word_Otext"].ToString();
        //                int WordLength = Convert.ToInt32(foundRows[iWord]["Word_Length"]);
        //                string WordType = foundRows[iWord]["Word_Type"].ToString();
        //                int lengthDiff;
        //                if (!string.IsNullOrEmpty(WordType))
        //                {
        //                    if (WordType.ToUpper() != "N")
        //                    {
        //                        for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
        //                        {
        //                            var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
        //                            lengthDiff = WordLength - skeysplit[0].Length;
        //                            if (skeysplit.Length > 1)
        //                            {
        //                                if (lengthDiff >= -2 && lengthDiff <= 2)
        //                                {
        //                                    if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                                    {
        //                                        if (iWord >= RowCount)
        //                                        {
        //                                        }
        //                                        else
        //                                        {
        //                                            int iIndex;
        //                                            var loopTo3 = skeysplit.Length - 1;
        //                                            for (iIndex = 1; iIndex <= loopTo3; iIndex++)
        //                                            {
        //                                                if (iWord + iIndex >= RowCount)
        //                                                    continue;
        //                                                lengthDiff = Convert.ToInt32(foundRows[iWord + iIndex]["Word_Length"]) - skeysplit[iIndex].Length;
        //                                                string oWord = foundRows[iWord + iIndex]["word_Otext"].ToString();
        //                                                if(oWord.Contains(":"))
        //                                                {
        //                                                    string[] splitword;
        //                                                    if (oWord.Contains(":")&& oWord.Any(char.IsDigit)&&(oWord.Count(char.IsDigit)>1))
        //                                                    {
        //                                                        splitword = oWord.Split(':');
        //                                                        oWord = splitword[0];
        //                                                        lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
        //                                                    }
        //                                                    //else if(oWord.Contains("#"))
        //                                                    //{
        //                                                    //    splitword = oWord.Split('#');
        //                                                    //    oWord = splitword[0];
        //                                                    //    lengthDiff = Convert.ToInt32(oWord.Length - skeysplit[iIndex].Length);
        //                                                    //}
        //                                                }                                                        

        //                                                if (lengthDiff >= -2 && lengthDiff <= 2)
        //                                                {
        //                                                    //if (CheckValidWord(foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
        //                                                    if (CheckValidWord(oWord.ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
        //                                                    {
        //                                                        if (skeysplit.Length - 1 != iIndex)
        //                                                        {
        //                                                            continue;
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            DataRow dr;

        //                                                            dr = dtKeywordInfo.NewRow();
        //                                                            dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                            dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                                            dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                                            dr["Word No"] = WordNo;
        //                                                            dr["LineWordNo"] = WordNo;
        //                                                            dr["X1"] = foundRows[iWord]["X1"];
        //                                                            dr["Y1"] = foundRows[iWord]["Y1"];
        //                                                            dr["X2"] = foundRows[iWord]["X2"];
        //                                                            dr["Y2"] = foundRows[iWord]["Y2"];
        //                                                            dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                                            dtKeywordInfo.Rows.Add(dr);
        //                                                        }
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        break;
        //                                                    }
        //                                                }
        //                                                else if (iIndex == loopTo3)
        //                                                {
        //                                                    string word = foundRows[iWord + iIndex]["word_Otext"].ToString();
        //                                                    if (foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().StartsWith(skeysplit[iIndex].ToUpper()))
        //                                                    {
        //                                                        //string keyword = Word.Substring(0, skeysplit[0].Length);
        //                                                        //if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
        //                                                        //{
        //                                                        //    if (!Module1.RAKeywords.Contains(keyword) || keyword.ToUpper() == "RA")
        //                                                        //    {
        //                                                        //        continue;
        //                                                        //    }
        //                                                        //}
        //                                                        if (foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Any(char.IsDigit))
        //                                                        {
        //                                                            DataRow dr;

        //                                                            dr = dtKeywordInfo.NewRow();
        //                                                            dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                            dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                                            dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                                            dr["Word No"] = WordNo;
        //                                                            dr["LineWordNo"] = WordNo;
        //                                                            dr["X1"] = foundRows[iWord]["X1"];
        //                                                            dr["Y1"] = foundRows[iWord]["Y1"];
        //                                                            dr["X2"] = foundRows[iWord]["X2"];
        //                                                            dr["Y2"] = foundRows[iWord]["Y2"];
        //                                                            dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                                            dtKeywordInfo.Rows.Add(dr);
        //                                                        }
        //                                                    }
        //                                                }
        //                                                else
        //                                                { break; }
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            //else if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                            else if (lengthDiff >= -2 && lengthDiff <= 2)
        //                            {
        //                                //if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
        //                                if (CheckValidWord(Word.Trim(), new string[] { skeysplit[0] }) == true)
        //                                {
        //                                    DataRow dr;

        //                                    dr = dtKeywordInfo.NewRow();
        //                                    dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                    dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                    dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                    dr["Word No"] = WordNo;
        //                                    dr["LineWordNo"] = WordNo;
        //                                    dr["X1"] = foundRows[iWord]["X1"];
        //                                    dr["Y1"] = foundRows[iWord]["Y1"];
        //                                    dr["X2"] = foundRows[iWord]["X2"];
        //                                    dr["Y2"] = foundRows[iWord]["Y2"];
        //                                    dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                    dtKeywordInfo.Rows.Add(dr);
        //                                }
        //                            }
        //                            else
        //                            {
        //                                if(Word.ToUpper().StartsWith(skeysplit[0].ToUpper()))
        //                                {
        //                                    string[] StartsPO = {"PO:","PO#" };
        //                                    string keyword = Word.Substring(0, skeysplit[0].Length);
        //                                    if (Module1.RAKeywords.Contains(skeysplit[0].ToUpper()))
        //                                    {
        //                                        if (!Module1.RAKeywords.Contains(keyword)||keyword.ToUpper()=="RA")
        //                                        {
        //                                            continue;
        //                                        }
        //                                    }
        //                                    if(keyword.ToUpper()=="PO"&&(!StartsPO.Any(Word.ToUpper().StartsWith)))
        //                                    {
        //                                        continue;
        //                                    }
        //                                    if (Word.Any(char.IsDigit))
        //                                    {
        //                                        DataRow dr;

        //                                        dr = dtKeywordInfo.NewRow();
        //                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                        dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
        //                                        dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
        //                                        dr["Word No"] = WordNo;
        //                                        dr["LineWordNo"] = WordNo;
        //                                        dr["X1"] = foundRows[iWord]["X1"];
        //                                        dr["Y1"] = foundRows[iWord]["Y1"];
        //                                        dr["X2"] = foundRows[iWord]["X2"];
        //                                        dr["Y2"] = foundRows[iWord]["Y2"];
        //                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
        //                                        dtKeywordInfo.Rows.Add(dr);
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }

        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        // throw;
        //        //MessageBox.Show(ex.Message, "SearchKeywords");
        //    }

        //    return dtKeywordInfo;
        //}
        #endregion

        //public static DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        //{
        //    var dtKeywordInfo = new DataTable();
        //    dtKeywordInfo = MakeDt();
        //    try
        //    {
        //        // -----------------------------------------BILL OF LADING NUMBER/PURCHASE ORDER NUMBER/SHIPPER NUMBER KEYWORD------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "lhre Bestellung per FAX", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero"}

        //        var sKeyarray = new string[] { "Bill of Lading Number", "BOL #", "B/L #", "Manifest #", "Bill of Lading No", "Shipper's bill of loading no.", "BOL#", "BOL No","Bill of L ading No","Bill of Lading", "BI LL 0 F LADING","BOL ID",
        //                                       "Shipper's Bill of Ladias No","BOL Number","Bill of Lading #","Bill of LadingNumber","B/L NO.","B/L NUMBER","Purchase Order No","Purchase Order Number","Po#","Po #", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No",
        //                                       "Customer's PO","cust po#","purchase order","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","P.O. #","Consisnee's Refer.noe/po No","P.O. No","Purchase Order #","Cust P.O. No",
        //                                       "CUSTOMERORDERNUMBER","P.O. NUMBER(S)","Cust.P.O.#","PURCHASE-ORDER-NO","CUSTOMER PURCHASE ORDER #","P/O #","CUST. ORDER NO.","PONumber","Order", "Order #","Shipper Number",  "Ship ID#", "Shipper Ref #", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No", "Shipper No", "Shipment #", "Shipper's No","Shipper #", "Shipping Order No", "SID#", "Shipper No/Customer",
        //                                       "Shipment ID","Shipment No","Shipment number","SID","SHIPPER","Shipper Ref#","","No/Customer PO#","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No.","24 hr Emergency Phone Numbers", "Emergency Response Number:", "Emergency Response Phone#",
        //                                       "Emergency Contact Telephone #", "Emergency Contact","For 24 HR. Chemical Emergency Assistance Call", "Chemical Emergency Response", "Emergency Response Phone Number","24-HR Emergency","Emergency Contact Number","In case of Emergency","Chemtrec Emergency 24-Hour Number",
        //                                       "EMERGENCYCONTACTNUMBER:","24hrEMERGENCY PHONENUMBERS","CHEM EMER #","Emergency 24-Hour Number","EMERGENCY CONTACT CALL:","HAZMAT EMERGENCY PHONE","Emergency Phone:","Chemical Emergency","chemical Emergencies","Emergency Number","Emergency Contact Phone",
        //                                       "24 Hour Emergency","24 HR Emergency","EMERGENCY ASSISTANCE","Emergency Response","EMERGENCY RESPONSE#","Emergency TEL #","Consignee","Ship To","Destination / Consignee","Deliver to","Destination/Consignee","Shippers","TO","TO (CONSIGNEE):","CONSIGNEE PHONE NO","Destination","Customer Phone #:","CustomerPO#","Order Number" };

        //        // Dim sKeyarray() As String = {"Your Contact"}
        //        // -----------------------------------------------------------------------------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice No:"}

        //        Array.Sort(sKeyarray);

        //        DataTable dt = oMeta.DocDictionary;

        //        var CNC = new ClsCNC();
        //        // Dim Words As clsCnCWord() = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999)

        //        var PossibleInvoices = new List<clsCnCWord>();
        //        //var oRetF3 = new RetStructF3();
        //        var oCnc = new ClsCNC();
        //        var conflevel = new List<int>();
        //        conflevel.Clear();
        //        int iPage;
        //        //var loopTo = oMeta.PageCount;
        //        for (iPage = 1; iPage <= 1; iPage++)
        //        {
        //            int h = iQPUBLIC.PublicComponents.ImageHeight;
        //            int w = iQPUBLIC.PublicComponents.ImageWidth;
        //            //clsCnCWord[] Words = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999);
        //            clsCnCLine[] CncLines = CNC.GetLinesFromROI(oMeta.Page, iPage, 0, 0, 9999, 9999);
        //            var wordsArray = new clsCnCWord[0];
        //            var oClsCnc = new ClsCNC();
        //            var oWord = new clsCnCWord();
        //            if (CncLines is object)
        //            {
        //                if (CncLines.Length != 0)
        //                {
        //                    for (int line = 0; line <= CncLines.Length - 1; line++)
        //                    {
        //                        for (int i = 1; i <= CncLines[line].WordCount; i++)
        //                        {
        //                            try
        //                            {
        //                                for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
        //                                {
        //                                    var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
        //                                    if (skeysplit.Length > 1)
        //                                    {
        //                                        if (CheckValidWord(CncLines[line].Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[0] }) == true)
        //                                        {
        //                                            // If Words(i).strWord.ToString.ToUpper.Trim.Replace("'", "`").Contains(skeysplit(0).ToUpper.Trim) Then

        //                                            if (i >= CncLines[line].WordCount)
        //                                            {
        //                                            }
        //                                            else
        //                                            {
        //                                                int iIndex;
        //                                                var loopTo3 = skeysplit.Length - 1;
        //                                                for (iIndex = 0; iIndex <= loopTo3; iIndex++)
        //                                                {
        //                                                    if (i + iIndex > CncLines[line].WordCount)
        //                                                        continue;
        //                                                    if (CheckValidWord(CncLines[line].Word[i + iIndex].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[iIndex] }) == true)
        //                                                    {

        //                                                        // If Words(i + iIndex).strWord.ToString.ToUpper.Trim.Contains(skeysplit(iIndex).Trim.ToUpper()) Then

        //                                                        if (skeysplit.Length - 1 != iIndex)
        //                                                        {
        //                                                            continue;
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            DataRow drDataRow = dtKeywordInfo.NewRow();
        //                                                            drDataRow["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                            drDataRow["Page No"] = iPage;
        //                                                            drDataRow["Line No"] = CncLines[line].Word[i].LineNo;
        //                                                            drDataRow["Word No"] = i + iIndex;
        //                                                            drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
        //                                                            // drDataRow("X1") = Words(i).Top
        //                                                            // drDataRow("Y1") = Words(i).Left
        //                                                            // drDataRow("X2") = Words(i + iIndex).Bottom
        //                                                            // drDataRow("Y2") = Words(i + iIndex).Right
        //                                                            // drDataRow("Word No") = Words(i).WordNumber + iIndex
        //                                                            drDataRow["X1"] = CncLines[line].Word[i].Left;
        //                                                            drDataRow["Y1"] = CncLines[line].Word[i].Top;
        //                                                            drDataRow["X2"] = CncLines[line].Word[i + iIndex].Right;
        //                                                            drDataRow["Y2"] = CncLines[line].Word[i + iIndex].Bottom;
        //                                                            dtKeywordInfo.Rows.Add(drDataRow);
        //                                                        }
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        break;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    //else if (Words[i].strWord.ToString().ToUpper().Trim().Replace("'", "`").Contains(sKeyarray[jLoop].Trim().ToUpper()))
        //                                    else if (CheckValidWord(CncLines[line].Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[0] }) == true)
        //                                    {
        //                                        DataRow drDataRow;
        //                                        //if (sKeyarray[jLoop].Trim().ToUpper() == "PO" && Words[i].strWord.ToString().ToUpper().Trim().Contains("PO") && Words[i].strWord.ToString().ToUpper() != "PO")
        //                                        //{
        //                                        //    continue;
        //                                        //}

        //                                        drDataRow = dtKeywordInfo.NewRow();
        //                                        drDataRow["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                        drDataRow["Page No"] = iPage;
        //                                        drDataRow["Line No"] = CncLines[line].Word[i].LineNo;
        //                                        drDataRow["Word No"] = i;
        //                                        drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
        //                                        // drDataRow("X1") = Words(i).Top
        //                                        // drDataRow("Y1") = Words(i).Left
        //                                        // drDataRow("X2") = Words(i).Bottom
        //                                        // drDataRow("Y2") = Words(i).Right
        //                                        // drDataRow("Word No") = Words(i).WordNumber
        //                                        drDataRow["X1"] = CncLines[line].Word[i].Left;
        //                                        drDataRow["Y1"] = CncLines[line].Word[i].Top;
        //                                        drDataRow["X2"] = CncLines[line].Word[i].Right;
        //                                        drDataRow["Y2"] = CncLines[line].Word[i].Bottom;
        //                                        dtKeywordInfo.Rows.Add(drDataRow);
        //                                    }
        //                                }
        //                            }
        //                            catch (Exception ex)
        //                            { }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //MessageBox.Show(ex.Message, "SearchKeywords");
        //    }

        //    return dtKeywordInfo;
        //}
        public static DataTable MakeDt()
        {
            var dtKeyword = new DataTable();
            dtKeyword.Columns.Add("Keyword", typeof(string));
            dtKeyword.Columns.Add("Page No", typeof(string));
            dtKeyword.Columns.Add("Line No", typeof(int));
            dtKeyword.Columns.Add("Word No", typeof(string));
            dtKeyword.Columns.Add("LineWordNo", typeof(string));
            dtKeyword.Columns.Add("X1", typeof(string));
            dtKeyword.Columns.Add("Y1", typeof(string));
            dtKeyword.Columns.Add("X2", typeof(string));
            dtKeyword.Columns.Add("Y2", typeof(string));
            dtKeyword.Columns.Add("NoofWords", typeof(string));
            //dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
            return dtKeyword;
        }
        //private static bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        //{
        //    bool bFlag = false;
        //    int lengthofword;
        //    // Dim AllMatches As MatchCollection = Regex.Matches(sWord, "[a-zA-Z0-9]+", RegexOptions.IgnoreCase)
        //    // For Each SingleMatch As Match In AllMatches
        //    // If Array.IndexOf(arrAllWords, SingleMatch.Value.ToUpper()) >= 0 Then
        //    // bFlag = True
        //    // End If
        //    // Next
        //    var objRestruct056 = new iQPUBLIC.clsCNCSBR();
        //    iQPUBLIC.RetStructIQSBR058 objRestruct058;
        //    if (!string.IsNullOrEmpty(sWord))
        //    {

        //        // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
        //        var arrFinalkeyword = new List<string>();
        //        sWord = sWord.Replace("'", "");
        //        sWord = sWord.Replace("`", "");

        //        lengthofword = sWord.Length;
        //        if (lengthofword <= 5)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length <= 5)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 6 & lengthofword <= 8)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 4 && arrcheckkeywords[iLoop].Length <= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 9 & lengthofword <= 11)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 9 && arrcheckkeywords[iLoop].Length <= 11)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 12)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }

        //        for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)

        //        {
        //            objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
        //            if (lengthofword <= 5)
        //            {
        //                if (objRestruct058.CharChanged < 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 6 & lengthofword <= 8)
        //            {
        //                if (objRestruct058.CharChanged < 2)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 9 & lengthofword <= 11)
        //            {
        //                if (objRestruct058.CharChanged < 3)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 12)
        //            {
        //                if (objRestruct058.CharChanged < 4)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //        }
        //    }

        //    return bFlag;
        //}

        //changed 

        #region changed by komal
        //private static bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        //{
        //    bool bFlag = false;
        //    int lengthofword;

        //    var objRestruct056 = new iQPUBLIC.clsCNCSBR();
        //    iQPUBLIC.RetStructIQSBR058 objRestruct058;
        //    if (!string.IsNullOrEmpty(sWord))
        //    {

        //        // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
        //        var arrFinalkeyword = new List<string>();
        //        sWord = sWord.Replace("'", "");
        //        sWord = sWord.Replace("`", "");

        //        if (SpecialCharacters.Any(sWord.Contains))
        //        {
        //            sWord = Regex.Replace(sWord, @"[^0-9a-zA-Z]+", "");
        //            // sWord = RemoveSpecialCharacters(sWord);
        //        }
        //        lengthofword = sWord.Length;
        //        if (lengthofword <= 3)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length <= 3)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword > 3 & lengthofword <= 5)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length > 2 && arrcheckkeywords[iLoop].Length <= 6)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }

        //        }
        //        else if (lengthofword >= 6 & lengthofword <= 8)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 4 && arrcheckkeywords[iLoop].Length <= 10)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 9 & lengthofword <= 11)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 6 && arrcheckkeywords[iLoop].Length <= 14)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 12)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }

        //        for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)

        //        {
        //            objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
        //            if (SpecialCharacters.Any(arrFinalkeyword[iLoop].Contains))
        //            {
        //                arrFinalkeyword[iLoop] = Regex.Replace(arrFinalkeyword[iLoop], @"[^0-9a-zA-Z]+", "");
        //                // sWord = RemoveSpecialCharacters(sWord);
        //            }
        //            int  lengthofkeyword = arrFinalkeyword[iLoop].Length;
        //            if (lengthofword <= 3)
        //            {
        //                if (objRestruct058.CharChanged < 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            if ((lengthofword > 3 && lengthofword <= 5)&&lengthofkeyword>3)
        //            {
        //                if (objRestruct058.CharChanged <= 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 6 & lengthofword <= 8)
        //            {
        //                if (objRestruct058.CharChanged <= 2)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 9 & lengthofword <= 11)
        //            {
        //                if (objRestruct058.CharChanged <= 3)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 12)
        //            {
        //                if (objRestruct058.CharChanged <= 4)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //        }
        //    }

        //    return bFlag;
        //}

        //changed on 10/12

        #region changed on 10/12
        private static bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        {
            List<string> exactMatchingKeywords = new List<string>() {"PONUMBER","SHIPPERNO","BOL#","LADING#","B/L","B/L#","P/O#","BOLNO","SHIPPER#","LOAD#","BILLOF","BOLNUMBER","BILLOFLADING","BOLNBR","SHIPPERNOS","SHIPPERS","SHIP", "RA", "RA#", "RMA", "RMA#", "RGA", "RGA#","OFLADING","CUSTOMERORDER","P0ORDER","SHIPPERSNO","CUSTOMERPO#","BIOF","BLOF","BOL/ORDER#","BOL/SHIPMENT#","BOLSHIPMENT","PURCHASEORDER#","CUSTOMERPURCHASEORDER#","DOBOL","PROBOL" };
            bool bFlag = false;
            int lengthofword, KeywordLength;

            var objRestruct056 = new iQPUBLIC.clsCNCSBR();
            iQPUBLIC.RetStructIQSBR058 objRestruct058;
            if (!string.IsNullOrEmpty(sWord))
            {
                // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
                var arrFinalkeyword = new List<string>();
                sWord = sWord.Replace("'", "");
                sWord = sWord.Replace("`", "");
                sWord = sWord.Replace(":", "");
                sWord = sWord.Replace("-", "");
                sWord = sWord.Replace(";", "");
                sWord = sWord.Replace(".", "");

                if (SpecialCharacters.Any(sWord.Contains)&&(!sWord.Contains("#")))
                {
                    sWord = Regex.Replace(sWord, @"[^0-9a-zA-Z]+", "");
                    // sWord = RemoveSpecialCharacters(sWord);
                }
                if (SpecialCharacters.Any(sWord.StartsWith)&&(!SpecialCharacters.Contains(sWord)))
                {
                    sWord = sWord.Remove(0, 1);
                    //arrcheckkeywords[0] = Regex.Replace(arrcheckkeywords[0], @"[^0-9a-zA-Z]+", "");
                    // sWord = RemoveSpecialCharacters(sWord);
                }
                if (SpecialCharacters.Any(arrcheckkeywords[0].Contains) && (!arrcheckkeywords[0].Contains("#")))
                {
                    arrcheckkeywords[0] = Regex.Replace(arrcheckkeywords[0], @"[^0-9a-zA-Z]+", "");
                    // sWord = RemoveSpecialCharacters(sWord);
                }
                lengthofword = sWord.Length;
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace("'", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace("`", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace(":", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace("-", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace(";", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace(".", "");

                KeywordLength = arrcheckkeywords[0].Length;

                if (sWord.ToUpper() == arrcheckkeywords[0].ToUpper())
                {
                    return true;
                }

                if (KeywordLength <= 3)
                {
                    if (lengthofword <= 3)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }
                else if (KeywordLength > 3 & KeywordLength <= 5)
                {                    
                    if (lengthofword > 3 && lengthofword <= 7)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }
                else if (KeywordLength >= 6 & KeywordLength <= 8)
                {

                    if (lengthofword >= 4 && lengthofword <= 10)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }

                }
                else if (KeywordLength >= 9 & KeywordLength <= 11)
                {
                    if (lengthofword >= 6 && lengthofword <= 14)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }
                else if (KeywordLength >= 12)
                {
                    if (lengthofword >= 8)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }

                for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)
                {
                    objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true,CompareOptions.EliminateBlank);
                    if (SpecialCharacters.Any(arrFinalkeyword[iLoop].Contains) && (!arrcheckkeywords[0].Contains("#")))
                    {
                        arrFinalkeyword[iLoop] = Regex.Replace(arrFinalkeyword[iLoop], @"[^0-9a-zA-Z]+", "");
                        // sWord = RemoveSpecialCharacters(sWord);
                    }
                    int lengthofkeyword = arrFinalkeyword[iLoop].Length;
                    if (exactMatchingKeywords.Contains(arrcheckkeywords[0].ToUpper()))
                    {
                        if (objRestruct058.CharChanged == 0)
                        {
                            bFlag = true;
                        }
                    }
                    else if (Module1.RAKeywords.Contains(arrcheckkeywords[0].ToUpper()) && Module1.RAKeywords.Contains(sWord.ToUpper()))
                    {
                        if (Module1.RAKeywords.Contains(sWord))
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofkeyword <= 3)
                    {
                        if (objRestruct058.CharChanged < 1)
                        {
                            bFlag = true;
                        }
                    }
                    
                    else if ((lengthofkeyword > 3 && lengthofkeyword <= 5) && lengthofword > 3)
                    {
                        if (arrFinalkeyword[iLoop].ToUpper() == "LOAD")
                        {
                            if (objRestruct058.CharChanged == 0)
                            {
                                bFlag = true;
                            }
                        }
                        else if (objRestruct058.CharChanged <= 1)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofkeyword >= 6 & lengthofkeyword <= 8)
                    {
                        //if(sWord.ToUpper()== "QUIPMENT")
                        //{
                        //    bFlag = false;
                        //}
                        if (arrFinalkeyword[iLoop].ToUpper() == "SHIPMENT"|| arrFinalkeyword[iLoop].ToUpper() == "CHEMTEL")
                        {
                            if (objRestruct058.CharChanged <= 1)
                            {
                                bFlag = true;
                            }
                        }
                        else if (objRestruct058.CharChanged <= 2)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofkeyword >= 9 & lengthofkeyword <= 11)
                    {
                        if (arrFinalkeyword[iLoop].ToUpper() == "EMERGENCY")
                        {
                            if (objRestruct058.CharChanged <= 2)
                            {
                                bFlag = true;
                            }
                        }
                        else if (objRestruct058.CharChanged <= 3)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofkeyword >= 12)
                    {
                        if (objRestruct058.CharChanged <= 4)
                        {
                            bFlag = true;
                        }
                    }
                }
            }

            return bFlag;
        }
        #endregion
        #endregion
        #endregion
                
        #region Barcode Reading
        public static List<string> Barcodes = new List<string>();
        public static List<string> func_ReadBarcode()
        {
            // List<string> Barcodes = new List<string>();
            BarcodeReader.clsReadBarCode clsReadBarCode = new clsReadBarCode();
            Barcodes = clsReadBarCode.ReadBarcodeFromImage(iQPUBLIC.PublicComponents.PrimaryImagePath, 0, 0);


            return Barcodes;
        }
        #endregion

        #region Replace Functions 
        public static string FuncReplace(string ReplaceStr)
        {
            if (ReplaceStr[ReplaceStr.Length - 1] == ',' || ReplaceStr[ReplaceStr.Length - 1] == '.')
            {
                ReplaceStr = ReplaceStr.Remove(ReplaceStr.Length - 1, 1);
            }
            ReplaceStr = ReplaceStr.Replace("#", "");
            //if (ReplaceStr.Contains("."))
            //{
            //    ReplaceStr = ReplaceStr.Replace(".", "");
            //}

            return ReplaceStr;
        }
        public static string RemoveSpecialCharacters(string StrWord)
        {
            string strRetWord = string.Empty;
            try
            {
                //char[] SpecialCharacters = new char[] { '!','@','#','$','%','^','&','*','(',')','_','+','-','=',',','.',':',';','/'};



                // var MatchValues = SpecialCharacters.Where(StrWord => StrWord.Contains())
                foreach (string ch in SpecialCharacters)
                {
                    if (StrWord.Contains(ch))
                    {
                        StrWord = StrWord.Replace(ch, "");
                    }
                }
            }
            catch (Exception ex)
            {
                //throw;
            }
            return StrWord;
        }
        public static string ReplaceAlphanumericWord(string strword)
        {
            foreach (char ch in strword)
            {
                if (char.IsLetter(ch))
                {
                    char ReplacedChar = Module1.AlphatoNumeric_OnlyNumericFields(ch);
                    strword = strword.Replace(ch, ReplacedChar);
                    break;
                }
            }
            return strword;
        }
        public static string ReplaceBraces(string strword)
        {
            strword = strword.Replace("(", "");
            strword = strword.Replace(")", "");
            strword = strword.Replace("[", "");
            strword = strword.Replace("]", "");
            strword = strword.Replace("{", "");
            strword = strword.Replace("}", "");

            return strword;
        }
        
        //public static clsCnCWord ReplaceAlphanumericWord(clsCnCWord cncWord)
        //{
        //    foreach (char ch in cncWord.strWord)
        //    {
        //        if (char.IsLetter(ch))
        //        {
        //            char ReplacedChar = Module1.AlphatoNumeric_OnlyNumericFields(ch);
        //            cncWord.strWord = cncWord.strWord.Replace(ch, ReplacedChar);
        //            break;
        //        }
        //    }
        //    return cncWord;
        //}
        public static char AlphatoNumeric_OnlyNumericFields(char chValue)
        {
            switch (chValue)
            {
                case 'O':
                case 'Q':
                case 'D':
                case 'o':
                case 'Þ':
                case 'Ó':
                case 'ö':
                case 'ô':
                case 'ò':
                case 'ó':
                case 'õ':
                case 'ø':
                case 'Ö':
                case 'Ô':
                case 'Ò':
                //case 'Ó':
                case 'Õ':
                case 'Ø':
                    {
                        return '0';
                    }


                case 'S':
                    {
                        return '5';
                    }

                case 'I':
                case 'L':
                case 'i':
                case 'l':
                case 'î':
                case 'ï':
                case '¡':
                case 'ì':
                case 'í':
                case 'Ï':
                case 'Î':
                case 'Ì':
                case 'Í'
               :
                    {
                        return '1';
                    }

                case 'B':
                    {
                        return '8';
                    }

                case 'Z':
                case 'z':
                case 'ž':
                case 'Ž':
                    {
                        return '2';
                    }

                case 'b':
                    {
                        return '6';
                    }

                case 'g':
                case 'q':
                    {
                        return '9';
                    }

                default:
                    {
                        return chValue;
                    }
            }


        }
        #endregion

        #region FindPatternFor Multiple
        public static string GetPattern(string inputText)
        {
            if (inputText.Trim().Length == 0)
            {
                return "";
            }
            int intAsciiCode = 0;
            string strTemp = "";
            inputText = inputText.Trim().ToLower();

            foreach (char Ch in inputText)
            {
                intAsciiCode = Convert.ToInt32(Ch);
                if ((intAsciiCode >= 65 && intAsciiCode <= 90) || (intAsciiCode >= 97 && intAsciiCode <= 122))
                    strTemp = strTemp + "a";
                else if (intAsciiCode >= 48 && intAsciiCode <= 57)
                    strTemp = strTemp + "n";
                else if (intAsciiCode == 32)
                    strTemp = strTemp + "b";
                else
                    strTemp = strTemp + "s";
            }

            return strTemp;
        }
        #endregion
        //public static clsCnCWord func_GetDefaultWord(int PageNumber)
        //{
        //    clsCnCWord oWord = new clsCnCWord();
        //    var _with1 = oWord;
        //    _with1.X1Char = "5";
        //    _with1.Y1Char = "5";
        //    _with1.X2Char = "5";
        //    _with1.Y2Char = "5";
        //    _with1.Confidence = 90;
        //    //.intLineNumber = 0
        //    _with1.LineNo = 1;
        //    _with1.PageNo = PageNumber;
        //    _with1.Left = 250;
        //    _with1.Right = 250;
        //    _with1.Top = 400;
        //    _with1.Bottom = 400;
        //    _with1.strWord = "";
        //    _with1.ConfString = "";

        //    return oWord;

        //}
    }
}