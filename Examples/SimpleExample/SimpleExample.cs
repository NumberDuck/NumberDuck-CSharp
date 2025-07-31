using NumberDuck;

class Application
{
	static int Main(string[] args)
	{
		System.Console.Write("Simple Example\n");
		
		Workbook pWorkbook = new Workbook(Workbook.License.AGPL);
		Worksheet pWorksheet = pWorkbook.GetWorksheetByIndex(0);

		Cell pCell = pWorksheet.GetCellByAddress("A1");
		pCell.SetString("Totally cool spreadsheet!");

		pWorksheet.GetCell(1,1).SetFloat(3.1417f);

		pWorkbook.Save("SimpleExample.xls", Workbook.FileType.XLS);
		
		Workbook pWorkbookIn = new Workbook(Workbook.License.AGPL);
		if (pWorkbookIn.Load("SimpleExample.xls"))
		{
			Worksheet pWorksheetIn = pWorkbookIn.GetWorksheetByIndex(0);
			Cell pCellIn = pWorksheetIn.GetCell(0,0);
			System.Console.Write("Cell Contents: " + pCellIn.GetString() + "\n");
		}

		return 0;
	}
}
