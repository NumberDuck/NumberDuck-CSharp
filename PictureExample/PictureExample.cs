using NumberDuck;

class Application
{
	static int Main(string[] args)
	{
		System.Console.Write("Picture Example\n");

		// Dino images by Dave Catchpole
		// http://www.flickr.com/photos/yaketyyakyak/sets/72157629976688365/
		// http://creativecommons.org/licenses/by/2.0/deed.en

		System.Console.Write("Embedding picture!\n");
		// Construct our workbook, and grab the default worksheet
		Workbook pWorkbook = new Workbook();
		Worksheet pWorksheet = pWorkbook.GetWorksheetByIndex(0);

		// Create the picture from our source image
		// Number Duck can create pictures from JPG and PNG files
		Picture pPicture = pWorksheet.CreatePicture("Dino.jpg");

		// By default, the picture is located in the top left corner of the worksheet
		pPicture.SetX(2);
		pPicture.SetY(2);

		// X and Y refer to the cell, to position within the cell, set SubX and SubY
		pPicture.SetSubX(Worksheet.DEFAULT_COLUMN_WIDTH / 2);
		pPicture.SetSubY(Worksheet.DEFAULT_ROW_HEIGHT / 2);

		pWorkbook.Save("PictureExample.xls", Workbook.FileType.FILE_TYPE_XLS);

		

		System.Console.Write("Extracting picture!\n");
		// Load the excel file with the image we want to extract
		pWorkbook = new Workbook();
		pWorkbook.Load("PictureExample.xls");
		pWorksheet = pWorkbook.GetWorksheetByIndex(0);

		pPicture = pWorksheet.GetPictureByIndex(0);

		// Work out the correct file name to save as based on the picture format
		// Note that while Number Duck can only create PNG and JPEG, it can extract any format
		string sFileName = "DinoOut.";
		switch (pPicture.GetFormat())
		{
			case Picture.Format.FORMAT_DIB: sFileName += "bmp";
				break;
			case Picture.Format.FORMAT_EMF: sFileName += "emf";
				break;
			case Picture.Format.FORMAT_JPEG: sFileName += "jpg";
				break;
			case Picture.Format.FORMAT_PICT: sFileName += "pict";
				break;
			case Picture.Format.FORMAT_PNG: sFileName += "png";
				break;
			case Picture.Format.FORMAT_TIFF: sFileName += "tiff";
				break;
			case Picture.Format.FORMAT_WMF: sFileName += "wmf";
				break;
		}

		// now write to disk
		pPicture.GetBlob().Save(sFileName);

		return 0;
	}
}