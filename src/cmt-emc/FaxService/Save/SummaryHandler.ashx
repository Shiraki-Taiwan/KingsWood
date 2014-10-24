<%@ WebHandler Language="C#" Class="SummaryHandler" %>

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;
using System.IO;
using System.Web;

public class SummaryHandler : IHttpHandler
{
	public class Html
	{
		public string Head;
		public string Body;
		public string Floor;
	}

	public static readonly Html csvGroup1 = new Html();
	public static readonly Html csvGroup2 = new Html();

	public static readonly Html txtGroup1 = new Html();
	public static readonly Html txtGroup2 = new Html();

	static SummaryHandler()
	{
		csvGroup1.Head
			=
@"上 林 公 證 有 限 公 司,,,,,,,
KINGSWOOD SURVEY & MEASURER LTD,,,,,,,
版權所有 2014. All Rights Reserved.,,,,,,,
總表查詢,,,,,,,
船名：,{0},航次：,{1},結關日期：,{2},,
,,,,,,,
攬貨商編號,攬貨商,核對,單號S/O,件數,體積,板數,總重
";
		csvGroup1.Body
            = "{0},{1},{2},{3},{4},{5},{6},{7},(error:{8})\n";
		csvGroup1.Floor
			= "合計,,,,{0},{1},{2},{3}\n";

		csvGroup2.Head
			=
@"上 林 公 證 有 限 公 司,,,,,,,
KINGSWOOD SURVEY & MEASURER LTD,,,,,,,
版權所有 2014. All Rights Reserved.,,,,,,,
總表查詢,,,,,,,
船名：,{0},航次：,{1},結關日期：,{2},,
,,,,,,,
攬貨商,核對,單號S/O,件數,體積,總重
";
		csvGroup2.Body
			= "{0},{1},{2},{3},{4},{5}\n";
		csvGroup2.Floor
			= "合計,,,{0},{1},{2}\n";

		// {0:14}, {1:14}, {2:Y/M/D}
		txtGroup1.Head
			=
@"                                上 林 公 證 有 限 公 司
                            KINGSWOOD SURVEY & MEASURER LTD 
                          版權所有 2014. All Rights Reserved.
                                       總表查詢
--------------------------------------------------------------------------------------
船名：    {0}    航次：    {1}    結關日期：    {2}
--------------------------------------------------------------------------------------
攬貨商編號        攬貨商      核對   單號S/O      件數       體積      板數       總重
";
		// {0:4}, {1:12}, {2:6}, {3:6}, {4:6}, {5:4.2}, {6}, {7:4.2}
		txtGroup1.Body
            = "      {0}  {1}    {2}    {3}    {4}    {5}    {6}    {7}    (error:{8})";
		// {0:6}, {1:4.2}, {2:6}, {3:4.2}
		txtGroup1.Floor
			=
@"                                      ================================================
                                      總計      {0}    {1}    {2}    {3}";

		txtGroup2.Head
			=
@"                     上 林 公 證 有 限 公 司
                 KINGSWOOD SURVEY & MEASURER LTD 
               版權所有 2014. All Rights Reserved.
                            總表查詢
----------------------------------------------------------------
船名：{0}  航次：{1}  結關日期：{2}
----------------------------------------------------------------
      攬貨商      核對   單號S/O      件數       體積       總重
";
		txtGroup2.Body
			= "{1}    {2}    {3}    {4}    {5}    {7}";
		txtGroup2.Floor
			=
@"                          ======================================
                          總計      {0}    {1}    {3}";
	}

	public void ProcessRequest(HttpContext context)
	{
		string fileName = String.Empty;
		string format = context.Request.QueryString["f"];
		string vesselListID = context.Request.QueryString["vlid"];
		string vesselLine = context.Request.QueryString["line"];
		string ids = context.Request.QueryString["ids"];

		if (String.IsNullOrEmpty(vesselListID))
			throw new HttpException(404, "Not find");
		else if (String.IsNullOrEmpty(vesselLine))
			throw new HttpException(404, "Not find");
		else if (String.IsNullOrEmpty(ids))
			throw new HttpException(404, "Not find");

		szGroupType
			= context.Request.QueryString["t"];
		vesselListID
			= vesselListID.Replace("-", "").Replace("'", "");
		ids
			= ids.Replace("-", "").Replace("'", "");

		string[] rgIds = ids.Split(',');

		for (int i = 0; i < rgIds.Length; i++)
		{
			rgIds[i] = string.Format("'{0}'", rgIds[i]);
		}

		ids = string.Join(",", rgIds);

		string dbPath = File.ReadAllLines(context.Server.MapPath("../../App_Data/connection.txt"))[0];

		using (OleDbConnection conn =
			new OleDbConnection(String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", context.Server.MapPath("../../App_Data/" + dbPath))))
		{
			using (OleDbCommand cmd = conn.CreateCommand())
			{
				try
				{
					conn.Open();

					cmd.CommandText = "select top 1 * from VesselList where ID = ?";
					cmd.Parameters.Clear();
					cmd.Parameters.AddWithValue("?", vesselListID);

					using (OleDbDataReader rd = cmd.ExecuteReader())
					{
						if (rd.HasRows)
						{
							rd.Read();
							szVesselName = rd["VesselName"].ToString();
							szVesselList = rd["VesselNo"].ToString();
							szVesselOwner = rd["Owner"].ToString();
							szVesselDate = Convert.ToDateTime(rd["Date"]).ToString("yyyy/MM/dd");
						}
						rd.Close();
					}

					cmd.CommandText =
						"select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum, sum(Board) as TotalBoard " +
						"from FreightForm " +
						"where ID in (" + ids + ") and VesselID = ?" +
						"group by FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry " +
						"order by FreightForm.ID";
					cmd.Parameters.Clear();
					cmd.Parameters.AddWithValue("?", vesselListID);

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
						cmd.Parameters.AddWithValue("?", vesselLine);

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
                    throw;
				}
				finally
				{
					try { conn.Close(); }
					catch { }
				}
			}
		}

		StringBuilder sb = new StringBuilder();

		fileName
			= String.Format("{0}_{1}_{2}", szVesselName, szVesselList, szVesselDate.Replace("/", ""));

		switch (format)
		{
			case "csv":
				context.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + ".csv");
				context.Response.ContentType = "text/csv";

				Html csv = csvGroup2;
				if (szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
					csv = csvGroup1;

				sb.AppendFormat(csv.Head, szVesselName, szVesselList, szVesselDate);

				foreach (DbData d in dbData)
				{
					sb.AppendFormat(csv.Body
						, d.FO_ID
						, d.FO_Name
						, d.IsCheckedText
						, d.ID
						, d.TotalPiece
						, d.TotalVolume
						, d.TotalBoard
						, d.TotalWeightSum
                        , d.Memo);
				}
				sb.AppendFormat(csv.Floor
					, nTotalPiece
					, nTotalVolume
					, nTotalBoard
					, nTotalWeight);

				break;

			default:
				context.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + ".txt");
				context.Response.ContentType = "text/plain";

				Html txt = txtGroup2;
				if (szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
					txt = txtGroup1;

				// {0:14}, {1:14}, {2:Y/M/D}
				sb.AppendFormat(txt.Head, szVesselName.PadLeft(14), szVesselList.PadLeft(14), szVesselDate);

				foreach (DbData d in dbData)
				{
					d.FO_Name = String.Format("{0}", d.FO_Name);

					int padLen = 12;
					int charLen = Encoding.Default.GetBytes(d.FO_Name).Length;

					if (charLen > d.FO_Name.Length)
						padLen = padLen - charLen + d.FO_Name.Length;

					int padLen2 = 6;
					int charLen2 = Encoding.Default.GetBytes(d.IsCheckedText).Length;
					
					if (charLen2 > d.IsCheckedText.Length)
						padLen2 = padLen2 - charLen2 + d.IsCheckedText.Length;

					// {0:4}, {1:12}, {2:6}, {3:6}, {4:6}, {5:4.2}, {6:6}, {7:4.2}
					sb.AppendFormat(txt.Body
						// {0:4}
						, String.Format("{0}", d.FO_ID).PadLeft(4)
						// {1:12}
						, String.Format("{0}", d.FO_Name).PadLeft(padLen)
						// {2:6}
						, d.IsCheckedText.PadLeft(padLen2)
						// {3:6}
						, d.ID.PadLeft(6)
						// {4:6}
						, String.Format("{0}", d.TotalPiece).PadLeft(6)
						// {5:4.2}
						, String.Format("{0:0.00}", d.TotalVolume).PadLeft(7)
						// {6:6}
						, String.Format("{0}", d.TotalBoard).PadLeft(6)
						// {7:4.2}
						, String.Format("{0:0.00}", d.TotalWeightSum).PadLeft(7)
                        , d.Memo.PadLeft(7)).AppendLine();
				}

				// {0:6}, {1:4.2}, {2:6}, {3:4.2}
				sb.AppendFormat(txt.Floor
					, nTotalPiece.ToString().PadLeft(6)
					, String.Format("{0:0.00}", nTotalVolume).PadLeft(7)
					, nTotalBoard.ToString().PadLeft(6)
					, String.Format("{0:0.00}", nTotalWeight).PadLeft(7));

				break;
		}

		StringWriter sw = new StringWriter(sb);
		sw.Close();

		context.Response.Clear();
		context.Response.Charset = "big5";
		context.Response.HeaderEncoding = Encoding.Default;
		context.Response.ContentEncoding = Encoding.Default;

		//context.Response.BinaryWrite(new byte[] { 0xEF, 0xBB, 0xBF });
		context.Response.Write(sw);
		//StreamWriter sw = new StreamWriter(context.Response.OutputStream, Encoding.UTF8);
		//sw.Write(sb.ToString());
		//sw.Close();

		context.Response.Flush();
		context.Response.End();
	}

	public bool IsReusable
	{
		get
		{
			return false;
		}
	}

	double nTotalPiece = 0;
	double nTotalVolume = 0;
	double nTotalWeight = 0;
	double nTotalBoard = 0;
	string szVesselName = String.Empty;
	string szVesselList = String.Empty;
	string szVesselOwner = String.Empty;
	string szVesselDate = String.Empty;
	string szGroupType = String.Empty;
	List<DbData> dbData = new List<DbData>();

	public class DbData
	{
		public bool IsChecked;
		public double TotalPiece;
		public double TotalVolume;
		public double TotalBoard;
		public double TotalWeightSum;
		public double NeededVolume;
		public string ID;
		public string FO_ID;
		public string FO_Name;
		public string IsCheckedText
		{
			get
			{
				if (IsChecked) return ""; else return "未核對";
			}
		}

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

            if (double.TryParse(reader["TotalVolume"].ToString(), out TotalVolume)==false)
            {
                // TotalVolume = Convert.ToDouble(reader["TotalVolume"]);
                Memo += reader["TotalVolume"].ToString();
            }

            if (double.TryParse(reader["TotalBoard"].ToString(), out TotalBoard)==false)
            {
                // TotalBoard = Convert.ToDouble(reader["TotalBoard"]);
                Memo += reader["TotalBoard"].ToString();
            }

            if (double.TryParse(reader["TotalWeightSum"].ToString(), out TotalWeightSum)==false)
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