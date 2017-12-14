using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace YoutubeExcel
{
	public class ExcelManager
	{
		string path = "";
		int lastRowIdx = 0;
		const string youtubeAddressBasic = "https://youtu.be/";
		const string youtubeCaptionBasic = "https://www.youtube.com/timedtext_video?v=";

		_Application excel = new _Excel.Application();
		Workbook wb;
		Worksheet ws;

		YoutubeManager youtubeApi;

		public ExcelManager(string path, int lastLowIdx, int sheet)
		{
			this.path = path;
			this.lastRowIdx = lastLowIdx;

			wb = excel.Workbooks.Open(this.path);
			ws = wb.Worksheets[sheet];

			youtubeApi = new YoutubeManager();

			Console.WriteLine($"ExcelManager initialized. Path({path}), Sheet({sheet})");
		}

		public void Close()
		{
			wb.Close();
		}

		public async Task Process()
		{
			Console.WriteLine($"Process started. Last Row idx : {lastRowIdx}");

			for (var curRowIdx = 2; curRowIdx <= lastRowIdx; ++curRowIdx)
			{
				Console.WriteLine($"Processing... Current Row Idx : {curRowIdx}");

				var youtubeAddress = ReadCell(curRowIdx, 4);

				if (youtubeAddress == "")
					continue;

				var youtubeId = youtubeAddress.Substring(youtubeAddressBasic.Length);

				var youtubeTitle = await youtubeApi.GetVideoTitle(youtubeId);

				// Youtube Title 쓰기.
				WriteToCell(curRowIdx, 3, youtubeTitle);

				// 번역툴 주소 쓰기.
				WriteToCell(curRowIdx, 5, youtubeCaptionBasic + youtubeId);

				var isCaptionValid = await youtubeApi.GetVideoCaptionValid(youtubeId);

				Console.WriteLine($"Id({youtubeId}), Title({youtubeTitle}), IsValid({isCaptionValid})");

				if (isCaptionValid)
				{
					WriteToCell(curRowIdx, 6, "N");
				}
				else
				{
					WriteToCell(curRowIdx, 6, "Y");
				}
			}

			SaveFile();
			Console.WriteLine("Process Ended.");
		}

		// 엑셀에서의 row, column은 1부터 시작한다.
		public string ReadCell(int row, int column)
		{
			if (row <= 0 || column <= 0)
			{
				Console.WriteLine($"Invalid cell address input Row({row}), Column({column})");
				return null;
			}

			if (ws.Cells[row, column].Value2 != null)
			{
				return ws.Cells[row, column].Value2;
			}
			else
			{
				return "";
			}
		}

		// 마찬가지로 row, column은 1부터 시작한다.
		public void WriteToCell(int row, int column, string s)
		{
			if (row <= 0 || column <= 0)
			{
				Console.WriteLine($"Invalid cell address input Row({row}), Column({column})");
				return;
			}

			ws.Cells[row, column].Value2 = s;
		}

		public void SaveFile()
		{
			wb.Save();
		}

		public void SaveFileAs(string path)
		{
			wb.SaveAs(path);
		}

		private int FindLastFillingRow()
		{
			var usedRange = ws.UsedRange;
			return usedRange.Rows.Count;
		}
	}
}
