using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;


using T_COLINFO = HelperExcel.COLINFO;


namespace HelperExcel
{
	class COLINFO
	{
		string _strColName;
		string _strFormula;
		string _strFormat;

		public string ColName { get => _strColName; private set => _strColName = value; }
		public string Formula { get => _strFormula; private set => _strFormula = value; }
		public string Format { get => _strFormat; private set => _strFormat = value; }

		public COLINFO( string colName, string formula, string format = "" )
		{
			ColName = colName;
			Formula = formula;
			Format = format;
		}

	}

	class ExcelRemover
	{

		/*
		//Func<> f;
		//var member;
		 

		ExcelRemover(var rhs)
		{
			member = rhs;
			f = GetFuncFromTyp(member);
		}

		~ExcelRemover(){
		}

		public void Remove()
		{
			f(); 
			Marschal.ReleaseComObj(member);
		}


		*/
	}

	class ExcelCleaner : IDisposable
	{
		private void Put( Excel.Application app )
		{
			//_cont.push( new ExcelRemover( app, =>( app ){ app.close } ) )
		}

		public T Push<T>(T val)
		{
			//Put( val ); 

			return val;
		}


		public void Dispose()
		{
			/*
			 for cur in _cont
				cur.Remove();
			 */

			throw new NotImplementedException();
		}

		private Stack<ExcelRemover> _cont;
	}

	class ExcelProg
	{
		private const string _G_WWS_Daten_Sheet_Name = "Daten";
		private const string _G_Daten_Table_Name = "Tbl_WWS_Daten";

		public static string G_WWS_Daten_Sheet_Name => _G_WWS_Daten_Sheet_Name;
		public static string G_Daten_Table_Name => _G_Daten_Table_Name;

		private const string _connstring = @"ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=";

		public static string Connstring => _connstring;




		private static string _Q( string str )
		{
			const string quote = "\"";
			return quote + str + quote;
		}

		string strQuoteString( string str )
		{
			return "";
		}


		private static T_COLINFO[] ColInfo = {
			new T_COLINFO   ( "ZEDatum"         , "=[@Netto]+10" )
			, new T_COLINFO ( "ZEKW" 
				, "=" + _Q("KW") + "& TEXT( WEEKNUM([@ZEDatum], 21), "
					+ _Q("00") + " ) & "+ _Q("/") + "& YEAR([@ZEDatum]) ")
			, new T_COLINFO ( "Dauer ins Monat"
				, "=MAX( "
						+"MONTH([@ZEDatum])+ ( ( YEAR([@ZEDatum]) - YEAR([@BelegDat]) ) * 12 ) - MONTH([@BelegDat])"
						+", 0"
						+")"
						)
			, new T_COLINFO ( "BelegDatum" , "=DATEVALUE( [@BelegDat] )", "t/M/jjjj" )
			, new T_COLINFO ( "BelegJahrMon" 
				, "=CONCATENATE( YEAR( [@BelegDat] ), "+ _Q("/") +" ,TEXT( MONTH([@BelegDat]), "+ _Q("00") +" ) )" )
					
			, new T_COLINFO ( "Buch_Typ" , "=IF( [@[Betrag (€)]] > 0, "+ _Q("SOLL") +" , "+ _Q("HABEN") +" )" )
			, new T_COLINFO ( "ZEJahrMon" , "=TEXT( [@ZEDatum], "+ _Q("jjjj/MM") +" )" )
			//, new T_COLINFO ( "" , "" )
			//, new T_COLINFO ( "" , "" )
		};

		public void AddQueryResult( ref Excel.Workbook WB, string sqlQry, string sheetName, string tblName )
		{
			Excel.Worksheet WS = WB.Worksheets.Add();
			WS.Name = sheetName;

			Excel.Range rngDest = WS.get_Range("A1");

			Excel.ListObject lstObj = WS.ListObjects.AddEx(
				SourceType: 0
				, Source: Connstring
				, Destination: rngDest
				);

			Excel.QueryTable qryTbl = lstObj.QueryTable;
			{
				var p = qryTbl;

				p.CommandText = sqlQry;
				p.RowNumbers = false;
				p.FillAdjacentFormulas = false;
				p.PreserveFormatting = true;
				p.RefreshOnFileOpen = false;
				p.BackgroundQuery = true;
				p.RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
				p.SavePassword = false;
				p.SaveData = true;
				p.AdjustColumnWidth = true;
				p.RefreshPeriod = 0;
				p.PreserveColumnInfo = true;
				p.ListObject.DisplayName = tblName;
				p.ListObject.TableStyle = "TableStyleLight1";
				p.Refresh(BackgroundQuery: false);
			}

		}

		public void AddSubsets( ref Excel.Workbook WB )
		{
			/*
				OleDBConnect dbCon;

				sql = SELECT DISTINCT [KD-GRP] FROM dbo.vw_liquidi...

				using( dbCon(connString) )
				using( dbAdap( sql ) )
				{
					DataSet = dbAdap.GetDataset
				}

				string[] subvsetStrs = dataSet.ToArr()
			 */
			string[] subsetStrs = new string[] { "Edeka", "Markant", "Rewe", "Dritte" };

			string sqlQryFmt = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquiditätsplanung_v2] WHERE [RG-NR] <= ( SELECT MAX(RLRENR) FROM WWS_MIR.dbo.Tbl_WWS_HRELEP WHERE RLREDA < 20180500 ) AND [KD-GRP] = N'{0}'  ORDER BY [RG-NR]";
				//[dbo].[vw_Liquiditätsplanung_v2]
			
			foreach( var cur in subsetStrs )
			{
				AddQueryResult(ref WB
					, string.Format(sqlQryFmt, cur)
					, cur
					, "Tbl_WWS_" + cur);
			}
		}

		public void AddColumns(ref Excel.Worksheet WS )
		{
			Excel.ListObject lstObj;
			Excel.ListColumn lstColumn;
			Excel.Range rngColumn;

			foreach( var cur in ColInfo )
			{
				lstObj = WS.ListObjects[G_Daten_Table_Name];
				lstColumn = lstObj.ListColumns.Add();
				lstColumn.Name = cur.ColName;

				rngColumn = WS.get_Range( G_Daten_Table_Name + "["+ cur.ColName +"]" );
				rngColumn.FormulaR1C1 = cur.Formula;
				if( 0 < cur.Format.Length )
					rngColumn.NumberFormatLocal = cur.Format;
			}
		}

		public void GenerateNewPlan()
		{
			// using ( ExcelCleaner cleaner = new ExcelCleaner )
			// Excel.Application app  = cleaner.push( new Excel.Application() );
			Excel.Application app = new Excel.Application();

			Excel.Workbook WB = app.Workbooks.Add();
			Excel.Worksheet WS = WB.Worksheets[1];

			WS.Name = G_WWS_Daten_Sheet_Name;

			Excel.Range rngDBData = WS.get_Range("A1");

			string sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquiditätsplanung_v2]" 
				+" WHERE [RG-NR] <= ( SELECT MAX(RLRENR) FROM WWS_MIR.dbo.Tbl_WWS_HRELEP WHERE RLREDA < 20180500 )"
				+" ORDER BY [RG-NR]";
			//string connstring = @"ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=";
			//string odbcInfo = "LOGISTIK-8;ServerSPN=WWS_MIR;";

			

			Excel.ListObject lsObj = WS.ListObjects.AddEx(SourceType: 0, Source: Connstring
				, Destination: rngDBData);

			Excel.QueryTable qryTbl = lsObj.QueryTable;
			{
				var p = qryTbl;
				p.CommandText = sqlQry;
				p.RowNumbers = false;
				p.FillAdjacentFormulas = false;
				p.PreserveFormatting = true;
				p.RefreshOnFileOpen = false;
				p.BackgroundQuery = true;
				p.RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
				p.SavePassword = false;
				p.SaveData = true;
				p.AdjustColumnWidth = true;
				p.RefreshPeriod = 0;
				p.PreserveColumnInfo = true;
				p.ListObject.DisplayName = G_Daten_Table_Name;
				p.ListObject.TableStyle = "TableStyleLight1";
				p.Refresh(false);
			}

			AddColumns( ref WS );
			AddSubsets(ref WB);

			app.UserControl = true;
			app.Visible = true;

		}

	}

	class Program
	{


		static void Main(string[] args)
		{
			ExcelProg prog = new ExcelProg();
			prog.GenerateNewPlan();						
		}
	}
}
