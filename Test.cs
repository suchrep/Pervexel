using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Pervexel {

	class Test {

		public static void Main (string[] args) {

			if (args.Length == 0) {
				Console.WriteLine("I need xls file path.");
				return;
			}
			
			int worksheet = 0;
			if (args.Length >= 2) {
				int.TryParse(args[1], out worksheet);
			}

			using (var file = File.OpenRead(args[0])) {

				var compound = new CompoundStream(file);
				if (!compound.Open("Workbook"))
					throw new Exception("Can't find Workbook entry");

				var workbook = new WorkbookReader(compound);
				if (!workbook.Open(worksheet))
					throw new Exception("Can't find worksheet!");

				var row = 0;
				var col = 0;
				while (workbook.Read()) {
					while (workbook.Row > row) {
						Console.WriteLine();
						row++;
						col = 0;
					}
					while (workbook.Column > col) {
						Console.Write("\t");
						col++;
					}
					Console.Write(Convert.ToString(workbook.Value));
				}
			}
		}
	}
}
