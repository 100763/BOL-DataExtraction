using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;

namespace BOLSET4IQ
{
    static class Module1
    {

        public static List<packagecode> PackageCodes = new List<packagecode>();

        public static string Synonyms = "OTY~DESC~Description~DESCRIPT]ON~Pallet~Pallets~* WElGHT (SUBJECT~WElGHT~DEDCRIPTION~LBS~FRT~Freight~Packaging~Pieces~CODE CLASS~NMFC CODE~WEIGHT - LB~WGHT~Description~PC~PKG~NMFC Item~Gross Weight~PACKAGES~Type~NMFC No~NMFC number~WGT~Freight ID Description~PACKAGE TYPE~Quantity~Weight~WElGHT~Cl~Item Description~Commodity Description~Class~NMFC~Qty~Pieces~Units~Qty~CTNS~Pieces/Quantity~Package~WT~Pkgs";
        #region IQ Class Objects
        public static RetStructF3 objRetStructF3 = new RetStructF3();
        #endregion

        public static clsCnCWord ObjDescription = new clsCnCWord();
        public static clsCnCWord ObjNMFC = new clsCnCWord();
        public static clsCnCWord ObjClassCode = new clsCnCWord();
        public static clsCnCWord ObjWeight = new clsCnCWord();
        public static clsCnCWord ObjQuantity = new clsCnCWord();
        public static clsCnCWord ObjHazmatCode = new clsCnCWord();
        public static clsCnCWord ObjPalletCode = new clsCnCWord();
        public static clsCnCWord ObjWord = new clsCnCWord();

        public static int LineNo;

        public static DataTable dtHeader = null, dtDetail = null, dtMapTable = null, dtDetailFinal = null;


        public static List<clsCnCWord> PossibleWords = new List<clsCnCWord>();

        public static clsCnCWord func_GetDefaultWord(int PageNumber)
        {
            clsCnCWord oWord = new clsCnCWord();
            var _with1 = oWord;
            _with1.X1Char = "5";
            _with1.Y1Char = "5";
            _with1.X2Char = "5";
            _with1.Y2Char = "5";
            _with1.Confidence = 90;
            //.intLineNumber = 0
            _with1.LineNo = 1;
            _with1.PageNo = PageNumber;
            _with1.Left = 250;
            _with1.Right = 250;
            _with1.Top = 400;
            _with1.Bottom = 400;
            _with1.strWord = "";
            _with1.ConfString = "";

            return oWord;

        }

        public static void GetPackageCode()
        {
            if (PackageCodes != null && PackageCodes.Count > 0)
                return;

            PackageCodes.Add(new packagecode("BG", "BAGS"));
            PackageCodes.Add(new packagecode("BG", "BAG"));
            PackageCodes.Add(new packagecode("BL", "BALES"));
            PackageCodes.Add(new packagecode("BR", "BARRELS"));
            PackageCodes.Add(new packagecode("BS", "BASKETS"));
            PackageCodes.Add(new packagecode("BS", "BASKET"));
            PackageCodes.Add(new packagecode("BX", "BOXES"));
            PackageCodes.Add(new packagecode("BX", "BOX"));
            PackageCodes.Add(new packagecode("BK", "BUCKETS"));
            PackageCodes.Add(new packagecode("BK", "BUCKET"));
            PackageCodes.Add(new packagecode("BD", "BUNDLES"));
            PackageCodes.Add(new packagecode("BD", "BUNDLE"));
            PackageCodes.Add(new packagecode("BD", "BDL"));
            PackageCodes.Add(new packagecode("CN", "CANS"));
            PackageCodes.Add(new packagecode("CT", "CARTONS"));
            PackageCodes.Add(new packagecode("CT", "CARTON"));
            PackageCodes.Add(new packagecode("CT", "CTN"));
            PackageCodes.Add(new packagecode("CS", "CASES"));
            PackageCodes.Add(new packagecode("CS", "CASE"));
            PackageCodes.Add(new packagecode("CO", "COILS"));
            PackageCodes.Add(new packagecode("CO", "COIL"));
            PackageCodes.Add(new packagecode("CY", "CYLINDERS"));
            PackageCodes.Add(new packagecode("CY", "CYLINDER"));
            PackageCodes.Add(new packagecode("DR", "DRUMS"));
            PackageCodes.Add(new packagecode("DR", "DRUM"));
            PackageCodes.Add(new packagecode("EA", "EACH"));
            PackageCodes.Add(new packagecode("EN", "ENVELOPES"));
            PackageCodes.Add(new packagecode("EN", "ENVELOPES"));
            PackageCodes.Add(new packagecode("KT", "KITS"));
            PackageCodes.Add(new packagecode("LS", "LOOSE"));
            PackageCodes.Add(new packagecode("NS", "NO "));
            PackageCodes.Add(new packagecode("JT", "JOINT"));
            PackageCodes.Add(new packagecode("JT", "JOINTS"));
            PackageCodes.Add(new packagecode("KT", "KITS"));
            PackageCodes.Add(new packagecode("LS", "LOOSE"));
            PackageCodes.Add(new packagecode("NS", "NO PACKAGE CODE"));
            PackageCodes.Add(new packagecode("PK", "PACKAGES"));
            PackageCodes.Add(new packagecode("PK", "PACKAGE"));
            PackageCodes.Add(new packagecode("PL", "PAILS"));
            PackageCodes.Add(new packagecode("PT", "PALLETS"));
            PackageCodes.Add(new packagecode("PT", "PALLET"));
            PackageCodes.Add(new packagecode("PT", "PLT"));
            PackageCodes.Add(new packagecode("PT", "PALLET(S)"));
            PackageCodes.Add(new packagecode("PC", "PIECES"));
            PackageCodes.Add(new packagecode("RE", "REELS"));
            PackageCodes.Add(new packagecode("RL", "ROLLS"));
            PackageCodes.Add(new packagecode("RL", "ROLL"));
            PackageCodes.Add(new packagecode("RT", "RT"));
            //PackageCodes.Add(new packagecode("SW", "SHRINK WRAPPED PALLET"));
            PackageCodes.Add(new packagecode("SK", "SKIDS"));
            PackageCodes.Add(new packagecode("SK", "SKID"));
            PackageCodes.Add(new packagecode("SP", "SPOOLS"));
            PackageCodes.Add(new packagecode("TK", "TANKS"));
            PackageCodes.Add(new packagecode("TK", "TANK"));
            PackageCodes.Add(new packagecode("TB", "TOTE BINS"));
            PackageCodes.Add(new packagecode("TU", "TUBE"));
            PackageCodes.Add(new packagecode("TU", "TUBES"));
            PackageCodes.Add(new packagecode("UT", "UNITS"));
            PackageCodes.Add(new packagecode("UT", "UNIT"));
            PackageCodes.Add(new packagecode("CA", "CA"));
        }

        // supporting function 
        public static string func_RemoveSpecialCharacter(string str_Data)
        {
            string[] array_RemoveSpecialCharacter = new[] { " ", "  ", ":", ".", "/", "!", "@", "#", "$", "%", "^", "*", "'", @"\", ";", "_", "(", ")", "|", "[", "]", ",", "-" };

            foreach (string str_RemoveSpecialCharacter in array_RemoveSpecialCharacter)
                str_Data = str_Data.Replace(str_RemoveSpecialCharacter, "");

            return str_Data.ToUpper().Trim();
        }
    }



    public class packagecode
    {
        public string Key;
        public string Value;

        public packagecode(string k, string val)
        {
            Key = k;
            Value = val;
        }
    }
}