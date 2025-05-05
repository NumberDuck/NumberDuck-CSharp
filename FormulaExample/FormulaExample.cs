using NumberDuck;

class Application
{
	static void Main(string[] args)
	{
		System.Console.Write("Formula Example\n");
		System.Console.Write("Create a spreadsheet with formulas!\n\n");
		
		Workbook pWorkbook = new Workbook();
		Worksheet pWorksheet = pWorkbook.GetWorksheetByIndex(0);
			
		for (ushort i = 0; i < 5; i++)
		{
			Cell pCell = pWorksheet.GetCell(0, i);
			pCell.SetFloat(i * 2.34f);
		}

		pWorksheet.GetCell(1,0).SetString("=SUM(A1:A5)");
		pWorksheet.GetCell(1,1).SetString("=AVERAGE(A1:A5)");

		pWorksheet.GetCell(2,0).SetFormula("=SUM(A1:A5)");
		pWorksheet.GetCell(2,1).SetFormula("=AVERAGE(A1:A5)");
			
		pWorkbook.Save("FormulaExample.xls", Workbook.FileType.FILE_TYPE_XLS);



		Workbook pWorkbookIn = new Workbook();
		if (pWorkbookIn.Load("FormulaExample.xls"))
		{
			Worksheet pWorksheetIn = pWorkbookIn.GetWorksheetByIndex(0);
			Cell pCellIn = pWorksheetIn.GetCell(2,1);
			System.Console.Write("Formula: " + pCellIn.GetFormula() + "\n");
		}
	}
}
