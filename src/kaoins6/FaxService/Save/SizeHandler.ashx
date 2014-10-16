<%@ WebHandler Language="C#" Class="SizeHandler" %>

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;
using System.IO;
using System.Web;

public class SizeHandler : IHttpHandler
{
	public class Html
	{
		public string Head;
		public string Body;
		public string FloorSubtotal;
		public string FloorTotal;
	}

	public static readonly Html csv = new Html();
	public static readonly Html txt = new Html();

	static SizeHandler()
	{
		csv.Head
			=
@"上 林 公 證 有 限 公 司,,,,,,,,,,,,
KINGSWOOD SURVEY & MEASURER LTD,,,,,,,,,,,,
版權所有 2014. All Rights Reserved.,,,,,,,,,,,,
尺寸資料查詢,,,,,,,,,,,,
船名：,{0},航次：,{1},結關日期：,{2},,,,,,,
頁次,單號,倉位,板數,堆量,件數,包裝,長,寬,高,體積,單重,總重
";
		csv.Body
			= "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}";
		csv.FloorSubtotal
			= "小計,,,,,{0},,,,,{1},,";
		csv.FloorTotal
			= "總計,,,,,{0},,,,,{1},,";

		// {0:14}, {1:14}, {2:Y/M/D}
		txt.Head
			=
@"                                           上 林 公 證 有 限 公 司
                                       KINGSWOOD SURVEY & MEASURER LTD 
                                     版權所有 2013. All Rights Reserved.
                                                尺寸資料查詢
------------------------------------------------------------------------------------------------------------
船名：    {0}    航次：    {1}    結關日期：    {2}
------------------------------------------------------------------------------------------------------------
  頁次    單號    倉位    板數    堆量    件數    包裝     長度     寬度     高度     體積     單重     總重
";
		// {0:6}, {1:6}, {2:6}, {3:6}, {4:6}, {5:6}, {6:6}, {7:4.2}, {8:4.2}, {9:4.2}, {10:4.2}, {11:4.2}, {12:4.2}
		txt.Body
			= "{0}  {1}  {2}  {3}  {4}  {5}  {6}  {7}  {8}  {9}  {10}  {11}  {12}";
		// {0:6}, {1:4.2}
		txt.FloorSubtotal
			=
@"                                ----------------------------------------------------------------------------
                                  小計  {0}                                     {1}";
		// {0:6}, {1:4.2}
		txt.FloorTotal
			=
@"                        ====================================================================================
                          總計          {0}                                     {1}";
	}

	public void ProcessRequest(HttpContext context)
	{
		string fileName = String.Empty;
		string format = context.Request.QueryString["f"];
		string vesselListID = context.Request.QueryString["vlid"];
		string vesselLine = context.Request.QueryString["line"];
		string sns = context.Request.QueryString["ids"];

		if (String.IsNullOrEmpty(vesselListID))
			throw new HttpException(404, "Not find");
		else if (String.IsNullOrEmpty(vesselLine))
			throw new HttpException(404, "Not find");
		else if (String.IsNullOrEmpty(sns))
			throw new HttpException(404, "Not find");

		szGroupType = context.Request.QueryString["t"];

		vesselListID = vesselListID.Replace("-", "").Replace("'", "");
		sns = sns.Replace("-", "").Replace("'", "");

		string[] rgIds = sns.Split(',');

		for (int i = 0; i < rgIds.Length; i++)
		{
			rgIds[i] = string.Format("{0}", rgIds[i]);
		}

		string dbPath = System.IO.File.ReadAllLines(context.Server.MapPath("../../App_Data/connection.txt"))[0];

		sns = string.Join(",", rgIds);

		using (OleDbConnection conn = new OleDbConnection(String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", context.Server.MapPath("../../App_Data/" + dbPath))))
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
							szVesselDate = Convert.ToDateTime(rd["Date"]).ToShortDateString();
						}
						rd.Close();
					}

					cmd.CommandText =
						"select ID, SN, PageNo, Volume, Board, IsPL, Length, Height, Width, Weight, Storehouse, Piece, PackageStyleID, TotalWeight, NeededForestry " +
						"from FreightForm " +
						"where SN in (" + sns + ") " +
						"  and VesselID = ? " +
						"order by ID, SN, PageNo ";

					cmd.Parameters.Clear();
					cmd.Parameters.AddWithValue("?", vesselListID);

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

		fileName
			= String.Format("{0}_{1}_{2}", szVesselName, szVesselList, szVesselDate.Replace("/", ""));

		StringBuilder sb = new StringBuilder();

		switch (format)
		{
			case "csv":
				context.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + ".csv");
				context.Response.ContentType = "text/csv";

				//Html csv = csvGroup2;
				//if (szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
				//	csv = csvGroup1;

				sb.AppendFormat(csv.Head, szVesselName, szVesselList, szVesselDate);

				foreach (KeyValuePair<string, DbSummaryData> summary in dbSummaryData)
				{
					foreach (DbData d in summary.Value.Items)
					{
						sb.AppendFormat(csv.Body
							, d.PageNo
							, d.ID
							, d.Storehouse
							, d.Board
							, d.IsPLText
							, d.Piece
							, d.PackageStyle
							, d.Length
							, d.Width
							, d.Height
							, d.Volume
							, d.Weight
							, d.TotalWeight).AppendLine();
					}
					sb.AppendFormat(csv.FloorSubtotal
						, summary.Value.nTotalPiece
						, summary.Value.nTotalVolume).AppendLine();
				}

				sb.AppendFormat(csv.FloorTotal
					, nTotalPiece
					, nTotalVolume);

				break;

			default:
				context.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + ".txt");
				context.Response.ContentType = "text/plain";

				//Html txt = txtGroup2;
				//if (szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
				//	txt = txtGroup1;

				// {0:14}, {1:14}, {2:Y/M/D}
				sb.AppendFormat(txt.Head, szVesselName.PadLeft(14), szVesselList.PadLeft(14), szVesselDate);

				foreach (KeyValuePair<string, DbSummaryData> summary in dbSummaryData)
				{
					sb.Append("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~").AppendLine();
					foreach (DbData d in summary.Value.Items)
					{
						int padLen = 6;
						int charLen = Encoding.Default.GetBytes(d.IsPLText).Length;

						if (charLen > d.IsPLText.Length)
							padLen = padLen - charLen + d.IsPLText.Length;

						d.PackageStyle = String.Format("{0}", d.PackageStyle);
						
						int padLen2 = 6;
						int charLen2 = Encoding.Default.GetBytes(d.PackageStyle).Length;

						if (charLen2 > d.PackageStyle.Length)
							padLen2 = padLen2 - charLen2 + d.PackageStyle.Length;

						// {0:6}, {1:6}, {2:6}, {3:6}, {4:6}, {5:6}, {6:6}, {7:4.2}, {8:4.2}, {9:4.2}, {10:4.2}, {11:4.2}, {12:4.2}
						sb.AppendFormat(txt.Body
							, String.Format("{0}", d.PageNo).PadLeft(6)
							, String.Format("{0}", d.ID).PadLeft(6)
							, String.Format("{0}", d.Storehouse).PadLeft(6)
							, String.Format("{0}", d.Board).PadLeft(6)
							, String.Format("{0}", d.IsPLText).PadLeft(padLen)
							, String.Format("{0}", d.Piece).PadLeft(6)
							, String.Format("{0}", d.PackageStyle).PadLeft(padLen2)
							, String.Format("{0:0.00}", d.Length).PadLeft(7)
							, String.Format("{0:0.00}", d.Width).PadLeft(7)
							, String.Format("{0:0.00}", d.Height).PadLeft(7)
							, String.Format("{0:0.00}", d.Volume).PadLeft(7)
							, String.Format("{0:0.00}", d.Weight).PadLeft(7)
							, String.Format("{0:0.00}", d.TotalWeight).PadLeft(7)).AppendLine();
					}
					// {0:6}, {1:4.2}
					sb.AppendFormat(txt.FloorSubtotal
						, summary.Value.nTotalPiece.ToString().PadLeft(6)
						, String.Format("{0:0.00}", summary.Value.nTotalVolume).PadLeft(7)).AppendLine();
				}

				// {0:6}, {1:4.2}
				sb.AppendFormat(txt.FloorTotal
					, nTotalPiece.ToString().PadLeft(6)
					, String.Format("{0:0.00}", nTotalVolume).PadLeft(7));

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