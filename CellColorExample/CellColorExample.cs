using NumberDuck;

class Application
{
	static int Main(string[] args)
	{
		System.Console.Write("Cell Color Example\n");
		System.Console.Write("Set cell size and background color!\n");

		// Load an image to get the colors we should assign to cells
		System.Drawing.Bitmap pImage = new System.Drawing.Bitmap("Duck.png");
		
		int nWidth = pImage.Width;
		int nHeight = pImage.Height;
		
		// the image could be larger than the maximum worksheet bounds, check for that to be safe
		if (nWidth > Worksheet.MAX_COLUMN || nHeight > Worksheet.MAX_ROW)
		{
			System.Console.Write("Sorry, Duck.png is too large. Max dimensions: %dx%d\n", Worksheet.MAX_COLUMN, Worksheet.MAX_ROW);
			return 1;
		}
		
		// Now we construct the workbook, and grab the default worksheet. Ready to start messing with cells
		Workbook pWorkbook = new Workbook();
		Worksheet pWorksheet = pWorkbook.GetWorksheetByIndex(0);

		// Now that we have our image and our worksheet we'll setup a nice square grid
		// Here we are setting the cell size to 7px by 7px.
		for (ushort nX = 0; nX < nWidth; nX++)
			pWorksheet.SetColumnWidth(nX, 7);

		for (ushort nY = 0; nY < nHeight; nY++)
			pWorksheet.SetRowHeight(nY, 7);

		for (ushort nY = 0; nY < nHeight; nY++)
		{
			for (ushort nX = 0; nX < nWidth; nX++)
			{
				// Get the data offset into the packed image data and read out the cell color there
				System.Drawing.Color pImageColor = pImage.GetPixel(nX, nY);
		
				byte nRed = pImageColor.R;
				byte nGreen = pImageColor.G;
				byte nBlue = pImageColor.B;
				byte nAlpha = pImageColor.A;

				// Skip transparent pixels
				if (nAlpha > 0)
				{
					Color pColor = new Color(nRed, nGreen, nBlue);

					Style pStyle = null;
					// reuse the same style if it has the same background color
					for (uint i = 0; i < pWorkbook.GetNumStyle(); i++)
					{
						Style pTestStyle = pWorkbook.GetStyleByIndex(i);
						Color pTestColor = pTestStyle.GetBackgroundColor(false);
						if (pTestColor != null && pTestColor.Equals(pColor))
						{
							pStyle = pTestStyle;
							break;
						}
					}

					// reusable style not found, create a new one
					if (pStyle == null)
					{
						pStyle = pWorkbook.CreateStyle();
						pStyle.GetBackgroundColor(true).SetFromColor(pColor);
					}

					// set the cell style (color)
					Cell pCell = pWorksheet.GetCell(nX, nY);
					pCell.SetStyle(pStyle);
				}
			}
		}

		pWorkbook.Save("CellColorExample.xls", Workbook.FileType.FILE_TYPE_XLS);
		return 0;
	}
}
