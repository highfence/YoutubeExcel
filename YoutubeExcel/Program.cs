using Google.Apis.Services;
using Google.Apis.YouTube.v3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YoutubeExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			ExcelManager manager = new ExcelManager(args[0], Convert.ToInt32(args[1]), Convent.ToInt32(args[2]));
			manager.Process();

			Console.ReadLine();
		}
	}
}
