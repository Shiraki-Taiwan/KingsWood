using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class FaxService_PrintOfSelect2 : System.Web.UI.Page
{
    protected Dictionary<string, DbSummaryData> dbSummaryData = new Dictionary<string, DbSummaryData>();

    protected double nTotalPiece = 0;

    protected double nTotalVolume = 0;

    protected double nTotalWeight = 0;

    protected double nTotalBoard = 0;

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
            string szSNs = Request.QueryString["ids"];

            if (String.IsNullOrEmpty(szVesselListID))
                throw new HttpException(404, "Not find");
            else if (String.IsNullOrEmpty(szVesselLine))
                throw new HttpException(404, "Not find");
            else if (String.IsNullOrEmpty(szSNs))
                throw new HttpException(404, "Not find");

            szGroupType = Request.QueryString["t"];

            szVesselListID = szVesselListID.Replace("-", "").Replace("'", "");
            szSNs = szSNs.Replace("-", "").Replace("'", "");

            string[] rgIds = szSNs.Split(',');

            for (int i = 0; i < rgIds.Length; i++)
            {
                //rgIds[i] = string.Format("'{0}'", rgIds[i]);
                rgIds[i] = string.Format("{0}", rgIds[i]);
            }

            string dbPath = System.IO.File.ReadAllLines(Server.MapPath("../App_Data/connection.txt"))[0];

            szSNs = string.Join(",", rgIds);

            using (OleDbConnection conn = new OleDbConnection(String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", Server.MapPath("../App_Data/" + dbPath))))
            {
                using (OleDbCommand cmd = conn.CreateCommand())
                {
                    try
                    {
                        conn.Open();

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
                            "select ID, SN, PageNo, Volume, Board, IsPL, Length, Height, Width, Weight, Storehouse, Piece, PackageStyleID, TotalWeight, NeededForestry " +
                            "from FreightForm " +
                            "where SN in (" + szSNs + ") " +
                            "  and VesselID = ? " +
                            "order by ID, SN, PageNo ";

                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("?", szVesselListID);

                        using (OleDbDataReader rd = cmd.ExecuteReader())
                        {
                            while (rd.Read())
                            {
                                DbData data = new DbData(rd);

                                if (!dbSummaryData.ContainsKey(data.ID))
                                    dbSummaryData.Add(data.ID, new DbSummaryData());

                                dbSummaryData[data.ID].Items.Add(data);
                                dbSummaryData[data.ID].nTotalBoard += data.Board;
                                dbSummaryData[data.ID].nTotalPiece += data.Piece;
                                dbSummaryData[data.ID].nTotalVolume += data.Volume;
                                dbSummaryData[data.ID].nTotalWeight += data.Weight;

                                nTotalWeight += data.Width;
                                nTotalBoard += data.Board;
                                nTotalPiece += data.Piece;
                                nTotalVolume += data.Volume;
                            }
                            rd.Close();
                        }

                        foreach (KeyValuePair<string, DbSummaryData> summary in dbSummaryData)
                        {
                            foreach (DbData item in summary.Value.Items)
                            {
                                cmd.CommandText =
                                    "select Name " +
                                    "from PackageStyle " +
                                    "where ID = ? ";

                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("?", item.PackageStyleID);

                                using (OleDbDataReader rd = cmd.ExecuteReader())
                                {
                                    if (rd.Read())
                                        item.PackageStyle = rd["Name"].ToString();
                                    rd.Close();
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {

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

    public class DbSummaryData
    {
        public readonly List<DbData> Items = new List<DbData>();

        public double nTotalPiece = 0;

        public double nTotalVolume = 0;

        public double nTotalWeight = 0;

        public double nTotalBoard = 0;
    }

    public class DbData
    {
        public bool IsPL { get; set; }

        public string ID { get; set; }
        public string SN { get; set; }
        public string IsPLText { get { return IsPL ? "堆量" : ""; } }
        public string Storehouse { get; set; }
        public string PackageStyleID { get; set; }
        public string PackageStyle { get; set; }

        public double PageNo { get; set; }
        public double Piece { get; set; }
        public double Volume { get; set; }
        public double Board { get; set; }
        public double Width { get; set; }
        public double Weight { get; set; }
        public double Height { get; set; }
        public double Length { get; set; }
        public double NeededForestry { get; set; }
        public double TotalWeight { get; set; }

        public DbData(OleDbDataReader reader)
        {
            IsPL = Convert.ToBoolean(reader["IsPL"]);
            ID = reader["ID"].ToString();
            SN = reader["SN"].ToString();
            Storehouse = reader["Storehouse"].ToString();
            PackageStyleID = reader["PackageStyleID"].ToString();

            PageNo = Convert.ToDouble(reader["PageNo"]);
            Piece = Convert.ToDouble(reader["Piece"]);
            Volume = Convert.ToDouble(reader["Volume"]);
            Board = Convert.ToDouble(reader["Board"]);
            Width = Convert.ToDouble(reader["Width"]);
            Weight = Convert.ToDouble(reader["Weight"]);
            Height = Convert.ToDouble(reader["Height"]);
            Length = Convert.ToDouble(reader["Length"]);
            NeededForestry = Convert.ToDouble(reader["NeededForestry"]);
            TotalWeight = Convert.ToDouble(reader["TotalWeight"]);
        }
    }
}