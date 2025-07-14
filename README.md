# Number Duck
v3.0.10 [J193]

Copyright (C) 2012-2025, File Scribe

https://numberduck.com

## 🦆 About
Number Duck is a programming library for developers to read and write Microsoft Excel compatible spreadsheets from a variety of languages.

See NumberDuck.html for API details, or check the example folders.

## 🚧 Installation
Since Number Duck is delivered as source, you can just drop `NumberDuck.cs` into your project.

## 💾 Saving and Loading a Spreadsheet
Here is a simple example of writing to a spreadsheet, and then reading it back in.

```csharp
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
```

More how to code in the example folders above, or at https://numberduck.com/docs

## 👮 License
Number Duck is dual licensed as Open Source (AGPLv3) and commercial closed source.

Closed source licenses may be purchased from https://numberduck.com