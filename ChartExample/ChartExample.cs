using NumberDuck;

class Application
{
	static int Main(string[] args)
	{
		System.Console.Write("Chart Example\n");
		System.Console.Write("Embed a chart in a spreadsheet!\n\n");
		
		Workbook pWorkbook = new Workbook(Workbook.License.AGPL);
		Worksheet pWorksheet = pWorkbook.GetWorksheetByIndex(0);

		Cell pCell = pWorksheet.GetCellByAddress("A1");
		pCell.SetString("Duck Productivity");
		
		// filling out the table of data
			pWorksheet.GetCellByAddress("B2").SetString("Quacks");
			pWorksheet.GetCellByAddress("C2").SetString("Waddles");
			
			System.Random pRandom = new System.Random();
			int nQuacks = 90;
			int nWaddles = 70;	
			for (int i = 0; i <  9; i++)
			{
				string sTemp = "Day " + (i+1);
				nQuacks += pRandom.Next(50);
				nWaddles += pRandom.Next(50);
				
				pWorksheet.GetCell(0, (ushort)(i+2)).SetString(sTemp);
				pWorksheet.GetCell(1, (ushort)(i+2)).SetFloat(nQuacks);
				pWorksheet.GetCell(2, (ushort)(i+2)).SetFloat(nWaddles);
			}
			
		// create the chart
		Chart pChart = pWorksheet.CreateChart(Chart.Type.TYPE_COLUMN);
		
		// set cell position for the chart
		pChart.SetX(3);
		pChart.SetY(1);
		
		// center the chart within the cell
		pChart.SetSubX(Worksheet.DEFAULT_COLUMN_WIDTH / 2);
		pChart.SetSubY(Worksheet.DEFAULT_ROW_HEIGHT / 2);
		
		pChart.SetWidth(Worksheet.DEFAULT_COLUMN_WIDTH*5);
		pChart.SetHeight(Worksheet.DEFAULT_ROW_HEIGHT*9);
		
		pChart.SetCategories("=A3:A11");
		
		// Create the 'quacks' series
		Series pSeries = pChart.CreateSeries("=B3:B11");
		pSeries.SetName("=B2");
		
		// Create the 'waddles' series
		pSeries = pChart.CreateSeries("=C3:C11");
		pSeries.SetName("=C2");
		
		pWorkbook.Save("ChartExample.xls", Workbook.FileType.XLS);

		return 0;
	}
}