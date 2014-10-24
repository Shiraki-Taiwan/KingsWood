using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class FaxService_PrintOfSelect : System.Web.UI.Page
{
    protected double nTotalPiece = 0;

    protected double nTotalVolume = 0;

    protected double nTotalWeight = 0;

    protected double nTotalBoard = 0;

    protected List<DbData> dbData = new List<DbData>();

    protected string szVesselName = String.Empty;

    protected string szVesselList = String.Empty;

    protected string szVesselOwner = String.Empty;

    protected string szVesselDate = String.Empty;

    protected string szGroupType = String.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string szVesselListID = Request.QueryString["vlid"];
            string szVesselLine = Request.QueryString["line"];
            string szIds = Request.QueryString["ids"];

            if (String.IsNullOrEmpty(szVesselListID))
                throw new HttpException(404, "Not find");
            else if (String.IsNullOrEmpty(szVesselLine))
                throw new HttpException(404, "Not find");
            else if (String.IsNullOrEmpty(szIds))
                throw new HttpException(404, "Not find");

            szGroupType = Request.QueryString["t"];

            szVesselListID = szVesselListID.Replace("-", "").Replace("'", "");
            szIds = szIds.Replace("-", "").Replace("'", "");

            string[] rgIds = szIds.Split(',');

            for (int i = 0; i < rgIds.Length; i++)
            {
                rgIds[i] = string.Format("'{0}'", rgIds[i]);
            }

            szIds = string.Join(",", rgIds);

            string dbPath = System.IO.File.ReadAllLines(Server.MapPath("../App_Data/connection.txt"))[0];

            using (OleDbConnection conn = new OleDbConnection(String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", Server.MapPath("../App_Data/" + dbPath))))
            {
                using (OleDbCommand cmd = conn.CreateCommand())
                {
                    try
                    {
                        conn.Open();
                        /*
                        cmd.CommandText = "select top 1 * from VesselList where ID = @vlid";
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@vlid", szVesselListID);
                        */
                        cmd.CommandText = "select top 1 * from VesselList where ID = ?";
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("?", szVesselListID);

                        using (OleDbDataReader rd = cmd.ExecuteReader())
                        {
                            if (rd.HasRows)
                            {
                                rd.Read();
                                szVesselName = rd["VesselName"].ToString();
                                szVesselList = rd["VesselNo"].ToString();
                                szVesselOwner = rd["Owner"].ToString();
                                szVesselDate = Convert.ToDateTime(rd["Date"]).ToShortDateString();
                            }
                            rd.Close();
                        }

                        cmd.CommandText =
                            "select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum, sum(Board) as TotalBoard " +
                            "from FreightForm " +
                            "where ID in (" + szIds + ") and VesselID = ?" +
                            "group by FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry " +
                            "order by FreightForm.ID";
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("?", szVesselListID);

                        using (OleDbDataReader rd = cmd.ExecuteReader())
                        {
                            while (rd.Read())
                            {
                                DbData data = new DbData(rd);

                                dbData.Add(data);

                                nTotalWeight += data.TotalWeightSum;
                                nTotalBoard += data.TotalBoard;
                                nTotalPiece += data.TotalPiece;

                                if (data.NeededVolume > 0)
                                    nTotalVolume += data.NeededVolume;
                                else
                                    nTotalVolume += data.TotalVolume;
                            }
                            rd.Close();
                        }

                        foreach (DbData item in dbData)
                        {
                            cmd.CommandText =
                                "select FreightOwner.ID, FreightOwner.Name " +
                                "from FreightOwner, FormToOwner " +
                                "where FormToOwner.FormID = ? " +
                                "and FormToOwner.OwnerID = FreightOwner.ID " +
                                "and FormToOwner.VesselLine = ?";
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("?", item.ID);
                            cmd.Parameters.AddWithValue("?", szVesselLine);

                            using (OleDbDataReader rd = cmd.ExecuteReader())
                            {
                                if (rd.Read())
                                {
                                    item.FO_ID = rd["ID"].ToString();
                                    item.FO_Name = rd["Name"].ToString();
                                }
                                rd.Close();
                            }
                        }
                    }
                    catch (Exception)
                    {
                        dbData.Clear();
                    }
                    finally
                    {
                        try { conn.Close(); }
                        catch { }
                    }
                }
            }
        }
    }

    public class DbData
    {
        public string ID { get; set; }

        public string FO_ID { get; set; }

        public string FO_Name { get; set; }

        public string IsCheckedText
        {
            get
            {
                if (IsChecked) return "";
                else return "未核對";
            }
        }

        public bool IsChecked { get; set; }

        public double TotalPiece;
        public double TotalVolume;
        public double TotalBoard;
        public double TotalWeightSum;
        public double NeededVolume;

        public string Memo = "";

        public DbData(OleDbDataReader reader)
        {
            ID = reader["ID"].ToString();
            IsChecked = Convert.ToBoolean(reader["IsChecked"]);

            if (double.TryParse(reader["TotalPiece"].ToString(), out TotalPiece) == false)
            {
                // TotalPiece = Convert.ToDouble(reader["TotalPiece"]);
                Memo += reader["TotalPiece"].ToString();
            }

            if (double.TryParse(reader["TotalVolume"].ToString(), out TotalVolume) == false)
            {
                // TotalVolume = Convert.ToDouble(reader["TotalVolume"]);
                Memo += reader["TotalVolume"].ToString();
            }

            if (double.TryParse(reader["TotalBoard"].ToString(), out TotalBoard) == false)
            {
                // TotalBoard = Convert.ToDouble(reader["TotalBoard"]);
                Memo += reader["TotalBoard"].ToString();
            }

            if (double.TryParse(reader["TotalWeightSum"].ToString(), out TotalWeightSum) == false)
            {
                // TotalWeightSum = Convert.ToDouble(reader["TotalWeightSum"]);
                Memo += reader["TotalWeightSum"].ToString();
            }

            if (double.TryParse(reader["NeededForestry"].ToString(), out NeededVolume) == false)
            {
                // NeededVolume = Convert.ToDouble(reader["NeededForestry"]);
                Memo += reader["NeededForestry"].ToString();
            }
        }
    }
}