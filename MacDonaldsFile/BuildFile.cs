using System;
using System.Linq;
using System.Collections.Generic;

using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace MacDonaldsFile
{
    public class BuildFile
    {

        public static void CreateFile(DAL.DataLayer dal, string sdg_name)
        {


            try
            {


                var oraCon = dal.GetOracleConnection();
                if (oraCon.State != System.Data.ConnectionState.Open)
                {
                    oraCon.Open();
                }
                sdg_name = sdg_name.Replace("\\", "").Replace("//", "");


                var ph = dal.GetPhraseByName("Location folders");
                string localFilePath = ph.PhraseEntries.FirstOrDefault(x => x.PhraseName == "Macdonalds").PhraseDescription;


                if (Directory.Exists(localFilePath) == false)
                {
                    Directory.CreateDirectory(localFilePath);
                }
                string LocalnewFile = sdg_name + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                string LocalnewFilePath = localFilePath + LocalnewFile;

                List<MacdonaldCls> List = new List<MacdonaldCls>();
                List = SelectSdg(sdg_name, oraCon);
                var csv = new StringBuilder();
                //  הזמנה מספר_דגימה  תאריך_דיגום שעת דיגום שם הדוגם תאריך_אישור_תוצאות  שם סניף     מספר הסניף  שם מנהל המשמרת תיאור_דוגמאות   יצרן מק_ט_בדיקה_מכון למיקרו תוצאה   יחידות_מידה טווח שמעליו הבדיקה אינה תקינה

                csv.AppendLine('"' + "הזמנה" + '"' + "," + '"' + "מספר_דגימה" + '"' + "," + '"' + "תאריך_דיגום" + '"' + "," + '"' + "שעת דיגום" + '"' + "," + '"' + "שם הדוגם" + '"' + "," + '"' + "תאריך_אישור_תוצאות" + '"' + "," + '"' + "שם ומספר סניף" + '"' + "," + '"' + "שם מנהל המשמרת" + '"' + "," + '"' + "תיאור_דוגמאות" + '"' + "," + '"' + "יצרן" + '"' + "," + '"' + "מק_ט_בדיקה_מכון למיקרו" + '"' + "," + '"' + "תוצאה" + '"' + "," + '"' + "יחידות_מידה" + '"' + "," + '"' + "טווח שמעליו הבדיקה אינה תקינה " + '"');

                foreach (var s in List)
                {
                    //SampleNo, string _Charecteristic, string _CharecteristicDesc, string _COAApprovalTime, string _Plant, string _MaterialDesc, string _QuantitativeResult, string _RangedQuantitativeResult, string _SeparatedResults, string _QualitativeResults, string _Catalog)
                    var SDGLine = string.Format('"' + "{0}" + '"' + "," + '"' + "{1}" + '"' + "," + '"' + "{2}" + '"' + "," + '"' + "{3}" + '"' + "," + '"' + "{4}" + '"' + "," + '"' + "{5}" + '"' + "," + '"' + "{6}" + '"' + "," + '"' + "{7}" + '"' + "," + '"' + "{8}" + '"' + "," + '"' + "{9}" + '"' + "," + '"' + "{10}" + '"' + "," + '"' + "{11}" + '"' + "," + '"' + "{12}" + '"' + "," + '"' + "{13}" + '"' + "," + '"' + "{14}" + '"', s.hazmana, s.dgima, s.taarich_digum, s.shaatDigum, s.dogemName, s.tarich_Ishur, s.snif_name_and_number, s.menael_Mishmeret, s.teur_Dugma, s.yatzran, s.makat, s.Tozaa, s.Mida, "","");
                    csv.AppendLine(SDGLine);
                }
                if (!Directory.Exists(localFilePath))
                {
                    Directory.CreateDirectory(localFilePath);
                }
                File.WriteAllText(LocalnewFilePath, csv.ToString(), System.Text.Encoding.UTF8);
                Log("File Saved in " + LocalnewFilePath);

            }

            catch (Exception e)
            {
                Log("Erorr on create Macdonalds file " + e.Message);
            }
        }
        public static void Log(string s)
        {
            try
            {
                Console.WriteLine(s);
                string LogPath = ConfigurationManager.AppSettings["LogPath"];
                using (FileStream file = new FileStream(LogPath + DateTime.Now.ToString("dd-MM-yyyy") + ".log", FileMode.Append, FileAccess.Write))
                {
                    var streamWriter = new StreamWriter(file);
                    streamWriter.WriteLine(s + " - " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    streamWriter.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Log File Error :" + e.Message + " - " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }


        }


        static List<MacdonaldCls> SelectSdg(string sdg_name, OracleConnection oraCon)
        {
            MacdonaldCls im = new MacdonaldCls();
            List<MacdonaldCls> ListIM = new List<MacdonaldCls>();


            string sql3 = " SELECT d.name hazmana, s.name dgima, DELIVERY_DATE taarich_digum , LIMS.get_field_by_col_name('שעת דיגום',U_Sdg_Client,U_Lab_Info_Id,s.sample_id) shaatDigum,    " +
 "        sop.FULL_NAME dogemName, S.Authorised_On tarich_Ishur, LIMS.get_field_by_col_name('מספר סניף', U_Sdg_Client, U_Lab_Info_Id, s.sample_id) snif_name_and_number, " +
 "        LIMS.get_field_by_col_name('שם מנהל/ ת משמרת', U_Sdg_Client, U_Lab_Info_Id, s.sample_id) menael_Mishmeret,    " +
 "      replace(s.description, ',', ' ')    teur_Dugma, LIMS.get_field_by_col_name('יצרן', U_Sdg_Client, U_Lab_Info_Id, s.sample_id) yatzran, Ttex.U_Headline_English makat, R.Formatted_Result Tozaa, U.Name Mida " +
 "        FROM LIMS_SYS.sdg d, LIMS_SYS.U_Test_Template_Ex_User ttex, LIMS_SYS.sdg_user du, LIMS_SYS.Unit u, LIMS_SYS.sample s, LIMS_SYS.Sample_User su, LIMS_SYS.aliquot a, " +
 "        LIMS_SYS.test t, LIMS_SYS.aliquot_user au, LIMS_SYS.result r, LIMS_SYS.operator sop " +
 "   WHERE R.Test_Id = T.Test_Id AND U.Unit_Id = Ttex.U_Default_Unit " +
 "   AND sop.operator_id = u_sampled_by AND Au.Aliquot_Id = A.Aliquot_Id " +
 "   AND T.Aliquot_Id = A.Aliquot_Id" +
 "   AND A.Sample_Id = S.Sample_Id" +
 "   AND R.Reported = 'T' AND Ttex.U_Test_Template_Ex_Id = Au.U_Test_Template_Extended         " +
 "   AND s.Sample_Id = Su.Sample_Id AND D.Sdg_Id = du.Sdg_Id AND D.Sdg_Id = S.Sdg_Id           " +
 "   AND D.Status <> 'X' AND s.Status <> 'X' AND a.Status <> 'X' AND t.Status <> 'X'           " +
 "   AND r.Status <> 'X'  AND R.Formatted_Result IS NOT NULL and d.name = '" + sdg_name.Trim() + "'" +
 "   AND rownum< 1000 ORDER BY S.Sample_Id,A.Aliquot_Id";




            OracleCommand cmd3 = new OracleCommand(sql3, oraCon);
            var rdr3 = cmd3.ExecuteReader();

            if (rdr3.HasRows)
            {

                while (rdr3.Read())
                {




                    im.hazmana = nvl(rdr3["hazmana"]);
                    im.dgima = nvl(rdr3["dgima"]) ;
                    im.taarich_digum = nvl(rdr3["taarich_digum"]) ;
                    im.shaatDigum = nvl(rdr3["shaatDigum"]) ;
                    im.dogemName = nvl(rdr3["dogemName"]) ;
                    im.tarich_Ishur = nvl(rdr3["tarich_Ishur"]) ;
                    im.snif_name_and_number = nvl(rdr3["snif_name_and_number"]) ;
                    im.menael_Mishmeret = nvl(rdr3["menael_Mishmeret"]) ;
                    im.teur_Dugma = nvl(rdr3["teur_Dugma"]) ;
                    im.yatzran = nvl(rdr3["yatzran"]) ;
                    im.makat = nvl(rdr3["makat"]) ;
                    im.Tozaa = nvl(rdr3["Tozaa"]) ;
                    im.Mida = nvl(rdr3["Mida"]) ;
                 //   im.description = nvl(rdr3["description"]) ;
                 //   im.u_sdg_client = nvl(rdr3["u_sdg_client"]) ;
                    ListIM.Add(im);
                }
            }




            return ListIM;
        }

        private static string nvl(object p)
        {
          
            if (p == null)
            {
                return "";
            }
            else
            {
                return p.ToString().Trim();
            }




        }
    }
}
