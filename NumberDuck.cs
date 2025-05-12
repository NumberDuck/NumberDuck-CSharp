/*
Number Duck
Copyright (c) 2012-2025 File Scribe

Closed source licenses may be purchased from https://numberduck.com

Otherwise:

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

namespace NumberDuck
{
	class Value
	{
		public enum Type
		{
			TYPE_EMPTY,
			TYPE_STRING,
			TYPE_FLOAT,
			TYPE_BOOLEAN,
			TYPE_FORMULA,
			TYPE_AREA,
			TYPE_AREA_3D,
			TYPE_ERROR,
		}

		public Secret.ValueImplementation m_pImpl;
		public Value()
		{
			m_pImpl = new Secret.ValueImplementation();
			m_pImpl.m_eType = Type.TYPE_EMPTY;
		}

		public bool Equals(Value pValue)
		{
			if (pValue == this)
				return true;
			if (m_pImpl.m_eType == pValue.m_pImpl.m_eType)
			{
				switch (m_pImpl.m_eType)
				{
					case Type.TYPE_EMPTY:
					{
						return true;
					}

					case Type.TYPE_STRING:
					{
						return Secret.ExternalString.Equal(m_pImpl.m_sString.GetExternalString(), pValue.m_pImpl.m_sString.GetExternalString());
					}

					case Type.TYPE_FLOAT:
					{
						double fDiff = m_pImpl.m_fFloat - pValue.m_pImpl.m_fFloat;
						return fDiff < 0.00001 && fDiff > -0.00001;
					}

					case Type.TYPE_BOOLEAN:
					{
						return m_pImpl.m_bBoolean == pValue.m_pImpl.m_bBoolean;
					}

					case Type.TYPE_FORMULA:
					{
						return Secret.ExternalString.Equal(GetFormula(), pValue.GetFormula());
					}

					case Type.TYPE_ERROR:
					{
						return true;
					}

				}
				Secret.nbAssert.Assert(false);
			}
			return false;
		}

		public new Value.Type GetType()
		{
			return m_pImpl.m_eType;
		}

		public string GetString()
		{
			if (m_pImpl.m_eType == Type.TYPE_STRING)
				return m_pImpl.m_sString.GetExternalString();
			return "";
		}

		public double GetFloat()
		{
			if (m_pImpl.m_eType == Type.TYPE_FLOAT)
				return m_pImpl.m_fFloat;
			return 0.0f;
		}

		public bool GetBoolean()
		{
			if (m_pImpl.m_eType == Type.TYPE_BOOLEAN)
				return m_pImpl.m_bBoolean;
			return false;
		}

		public string GetFormula()
		{
			if (m_pImpl.m_eType == Type.TYPE_FORMULA)
			{
				m_pImpl.m_sString.Set(m_pImpl.m_pFormula.ToString(m_pImpl.m_pWorksheet.m_pImpl));
				return m_pImpl.m_sString.GetExternalString();
			}
			return "=";
		}

		public Value EvaulateFormula()
		{
			if (m_pImpl.m_eType == Type.TYPE_FORMULA)
				return m_pImpl.m_pFormula.Evaluate(m_pImpl.m_pWorksheet.m_pImpl, 0);
			if (m_pImpl.m_pValue != null)
				{
				}
			m_pImpl.m_pValue = Secret.ValueImplementation.CreateErrorValue();
			return m_pImpl.m_pValue;
		}

		~Value()
		{
		}

	}
}

namespace NumberDuck
{
	class Cell
	{
		public Secret.CellImplementation m_pImpl;
		public Cell(Worksheet pWorksheet)
		{
			m_pImpl = new Secret.CellImplementation();
			m_pImpl.m_pValue = new Value();
			m_pImpl.m_pWorksheet = pWorksheet;
			Clear();
		}

		public bool Equals(Cell pCell)
		{
			return m_pImpl.m_pValue.Equals(pCell.m_pImpl.m_pValue) && m_pImpl.m_pStyle == pCell.m_pImpl.m_pStyle;
		}

		public Value GetValue()
		{
			return m_pImpl.m_pValue;
		}

		public new Value.Type GetType()
		{
			return m_pImpl.m_pValue.GetType();
		}

		public void Clear()
		{
			m_pImpl.m_pValue.m_pImpl.Clear();
			m_pImpl.m_pStyle = m_pImpl.m_pWorksheet.m_pImpl.m_pWorkbook.GetDefaultStyle();
		}

		public void SetString(string szString)
		{
			m_pImpl.m_pValue.m_pImpl.SetString(szString);
		}

		public string GetString()
		{
			return m_pImpl.m_pValue.GetString();
		}

		public void SetFloat(double fFloat)
		{
			m_pImpl.m_pValue.m_pImpl.SetFloat(fFloat);
		}

		public double GetFloat()
		{
			return m_pImpl.m_pValue.GetFloat();
		}

		public void SetBoolean(bool bBoolean)
		{
			m_pImpl.m_pValue.m_pImpl.SetBoolean(bBoolean);
		}

		public bool GetBoolean()
		{
			return m_pImpl.m_pValue.GetBoolean();
		}

		public void SetFormula(string szFormula)
		{
			m_pImpl.m_pValue.m_pImpl.SetFormulaFromString(szFormula, m_pImpl.m_pWorksheet);
		}

		public string GetFormula()
		{
			return m_pImpl.m_pValue.GetFormula();
		}

		public Value EvaulateFormula()
		{
			return m_pImpl.m_pValue.EvaulateFormula();
		}

		public bool SetStyle(Style pStyle)
		{
			Workbook pWorkbook = m_pImpl.GetWorkbook();
			ushort i;
			for (i = 0; i < pWorkbook.GetNumStyle(); i++)
				if (pWorkbook.GetStyleByIndex(i) == pStyle)
					break;
			if (i == pWorkbook.GetNumStyle())
				return false;
			m_pImpl.m_pStyle = pStyle;
			return true;
		}

		public Style GetStyle()
		{
			return m_pImpl.m_pStyle;
		}

		~Cell()
		{
		}

	}
}

namespace NumberDuck
{
	class Style
	{
		public Secret.StyleImplementation m_pImplementation;
		public Style()
		{
			m_pImplementation = new Secret.StyleImplementation();
		}

		public Font GetFont()
		{
			return m_pImplementation.m_pFont;
		}

		public enum HorizontalAlign
		{
			HORIZONTAL_ALIGN_GENERAL = 0,
			HORIZONTAL_ALIGN_LEFT,
			HORIZONTAL_ALIGN_CENTER,
			HORIZONTAL_ALIGN_RIGHT,
			HORIZONTAL_ALIGN_FILL,
			HORIZONTAL_ALIGN_JUSTIFY,
			HORIZONTAL_ALIGN_CENTER_ACROSS_SELECTION,
			HORIZONTAL_ALIGN_DISTRIBUTED,
			NUM_HORIZONTAL_ALIGN,
		}

		public Style.HorizontalAlign GetHorizontalAlign()
		{
			return m_pImplementation.m_eHorizontalAlign;
		}

		public void SetHorizontalAlign(Style.HorizontalAlign eHorizontalAlign)
		{
			if (eHorizontalAlign < HorizontalAlign.HORIZONTAL_ALIGN_GENERAL || eHorizontalAlign >= HorizontalAlign.NUM_HORIZONTAL_ALIGN)
				eHorizontalAlign = HorizontalAlign.HORIZONTAL_ALIGN_GENERAL;
			m_pImplementation.m_eHorizontalAlign = eHorizontalAlign;
		}

		public enum VerticalAlign
		{
			VERTICAL_ALIGN_TOP = 0,
			VERTICAL_ALIGN_CENTER,
			VERTICAL_ALIGN_BOTTOM,
			VERTICAL_ALIGN_JUSTIFY,
			VERTICAL_ALIGN_DISTRIBUTED,
			NUM_VERTICAL_ALIGN,
		}

		public Style.VerticalAlign GetVerticalAlign()
		{
			return m_pImplementation.m_eVerticalAlign;
		}

		public void SetVerticalAlign(Style.VerticalAlign eVerticalAlign)
		{
			if (eVerticalAlign < VerticalAlign.VERTICAL_ALIGN_TOP || eVerticalAlign >= VerticalAlign.NUM_VERTICAL_ALIGN)
				eVerticalAlign = VerticalAlign.VERTICAL_ALIGN_BOTTOM;
			m_pImplementation.m_eVerticalAlign = eVerticalAlign;
		}

		public Color GetBackgroundColor(bool bCreateIfMissing)
		{
			if (m_pImplementation.m_pBackgroundColor == null && bCreateIfMissing)
				m_pImplementation.m_pBackgroundColor = new Color(0xFF, 0xFF, 0xFF);
			return m_pImplementation.m_pBackgroundColor;
		}

		public void ClearBackgroundColor()
		{
			{
			}
			m_pImplementation.m_pBackgroundColor = null;
		}

		public enum FillPattern
		{
			FILL_PATTERN_NONE = 0,
			FILL_PATTERN_50,
			FILL_PATTERN_75,
			FILL_PATTERN_25,
			FILL_PATTERN_125,
			FILL_PATTERN_625,
			FILL_PATTERN_HORIZONTAL_STRIPE,
			FILL_PATTERN_VARTICAL_STRIPE,
			FILL_PATTERN_DIAGONAL_STRIPE,
			FILL_PATTERN_REVERSE_DIAGONAL_STRIPE,
			FILL_PATTERN_DIAGONAL_CROSSHATCH,
			FILL_PATTERN_THICK_DIAGONAL_CROSSHATCH,
			FILL_PATTERN_THIN_HORIZONTAL_STRIPE,
			FILL_PATTERN_THIN_VERTICAL_STRIPE,
			FILL_PATTERN_THIN_REVERSE_VERTICAL_STRIPE,
			FILL_PATTERN_THIN_DIAGONAL_STRIPE,
			FILL_PATTERN_THIN_HORIZONTAL_CROSSHATCH,
			FILL_PATTERN_THIN_DIAGONAL_CROSSHATCH,
			NUM_FILL_PATTERN,
		}

		public Style.FillPattern GetFillPattern()
		{
			return m_pImplementation.m_eFillPattern;
		}

		public void SetFillPattern(Style.FillPattern eFillPattern)
		{
			if (eFillPattern < FillPattern.FILL_PATTERN_NONE || eFillPattern >= FillPattern.NUM_FILL_PATTERN)
				eFillPattern = FillPattern.FILL_PATTERN_NONE;
			m_pImplementation.m_eFillPattern = eFillPattern;
		}

		public Color GetFillPatternColor(bool bCreateIfMissing)
		{
			if (m_pImplementation.m_pFillPatternColor == null && bCreateIfMissing)
				m_pImplementation.m_pFillPatternColor = new Color(0, 0, 0);
			return m_pImplementation.m_pFillPatternColor;
		}

		public void ClearFillPatternColor()
		{
			{
			}
			m_pImplementation.m_pFillPatternColor = null;
		}

		public Line GetTopBorderLine()
		{
			return m_pImplementation.m_pTopBorderLine;
		}

		public Line GetRightBorderLine()
		{
			return m_pImplementation.m_pRightBorderLine;
		}

		public Line GetBottomBorderLine()
		{
			return m_pImplementation.m_pBottomBorderLine;
		}

		public Line GetLeftBorderLine()
		{
			return m_pImplementation.m_pLeftBorderLine;
		}

		public string GetFormat()
		{
			return m_pImplementation.m_sFormat.GetExternalString();
		}

		public void SetFormat(string szFormat)
		{
			if (szFormat != null)
			{
				{
				}
				m_pImplementation.m_sFormat = new Secret.InternalString(szFormat);
			}
		}

		~Style()
		{
		}

	}
}

namespace NumberDuck
{
	class Line
	{
		protected Secret.LineImplementation m_pImpl;
		public enum Type
		{
			TYPE_NONE = 0,
			TYPE_THIN,
			TYPE_DASHED,
			TYPE_DOTTED,
			TYPE_DASH_DOT,
			TYPE_DASH_DOT_DOT,
			TYPE_MEDIUM,
			TYPE_MEDIUM_DASHED,
			TYPE_MEDIUM_DASH_DOT,
			TYPE_MEDIUM_DASH_DOT_DOT,
			TYPE_THICK,
			TYPE_DOUBLE,
			TYPE_HAIR,
			TYPE_SLANT_DASH_DOT_DOT,
			NUM_TYPE,
		}

		public Line()
		{
			m_pImpl = new Secret.LineImplementation();
		}

		public bool Equals(Line pLine)
		{
			if (pLine == null)
				return false;
			return m_pImpl.Equals(pLine.m_pImpl);
		}

		public new Line.Type GetType()
		{
			return m_pImpl.GetType();
		}

		public void SetType(Type eType)
		{
			m_pImpl.SetType(eType);
		}

		public Color GetColor()
		{
			return m_pImpl.m_pColor;
		}

		~Line()
		{
		}

	}
}

namespace NumberDuck
{
	class Marker
	{
		public Secret.MarkerImplementation m_pImpl;
		public enum Type
		{
			TYPE_NONE = 0,
			TYPE_SQUARE,
			TYPE_DIAMOND,
			TYPE_TRIANGLE,
			TYPE_X,
			TYPE_ASTERISK,
			TYPE_SHORT_BAR,
			TYPE_LONG_BAR,
			TYPE_CIRCULAR,
			TYPE_PLUS,
			NUM_TYPE,
		}

		public Marker()
		{
			m_pImpl = new Secret.MarkerImplementation();
		}

		public bool Equals(Marker pMarker)
		{
			if (pMarker == null)
				return false;
			return m_pImpl.Equals(pMarker.m_pImpl);
		}

		public new Marker.Type GetType()
		{
			return m_pImpl.GetType();
		}

		public void SetType(Marker.Type eType)
		{
			m_pImpl.SetType(eType);
		}

		public Color GetFillColor(bool bCreateIfMissing)
		{
			return m_pImpl.GetFillColor(bCreateIfMissing);
		}

		public void SetFillColor(Color pColor)
		{
			m_pImpl.SetFillColor(pColor);
		}

		public void ClearFillColor()
		{
			m_pImpl.ClearFillColor();
		}

		public Color GetBorderColor(bool bCreateIfMissing)
		{
			return m_pImpl.GetBorderColor(bCreateIfMissing);
		}

		public void SetBorderColor(Color pColor)
		{
			m_pImpl.SetBorderColor(pColor);
		}

		public void ClearBorderColor()
		{
			m_pImpl.ClearBorderColor();
		}

		public int GetSize()
		{
			return m_pImpl.GetSize();
		}

		public void SetSize(int nSize)
		{
			m_pImpl.SetSize(nSize);
		}

		~Marker()
		{
		}

	}
}

namespace NumberDuck
{
	class Fill
	{
		protected Secret.FillImplementation m_pImpl;
		public enum Type
		{
			TYPE_NONE = 0,
			TYPE_SOLID,
			NUM_TYPE,
		}

		public Fill()
		{
			m_pImpl = new Secret.FillImplementation();
		}

		public bool Equals(Fill pFill)
		{
			if (pFill == null)
				return false;
			return m_pImpl.Equals(pFill.m_pImpl);
		}

		public new Fill.Type GetType()
		{
			return m_pImpl.GetType();
		}

		public void SetType(Type eType)
		{
			m_pImpl.SetType(eType);
		}

		public Color GetForegroundColor()
		{
			return m_pImpl.m_pForegroundColor;
		}

		public Color GetBackgroundColor()
		{
			return m_pImpl.m_pBackgroundColor;
		}

		~Fill()
		{
		}

	}
}

namespace NumberDuck
{
	class Font
	{
		public Secret.FontImplementation m_pImpl;
		public enum Underline
		{
			UNDERLINE_NONE = 0,
			UNDERLINE_SINGLE,
			UNDERLINE_DOUBLE,
			UNDERLINE_SINGLE_ACCOUNTING,
			UNDERLINE_DOUBLE_ACCOUNTING,
			NUM_UNDERLINE,
		}

		public Font()
		{
			m_pImpl = new Secret.FontImplementation();
		}

		public void SetName(string szName)
		{
			m_pImpl.m_sName.Set(szName);
		}

		public string GetName()
		{
			return m_pImpl.m_sName.GetExternalString();
		}

		public void SetSize(byte nSize)
		{
			if (nSize < 10)
				nSize = 10;
			else if (nSize > 96)
				nSize = 96;
			m_pImpl.m_nSizeTwips = nSize * 15;
		}

		public byte GetSize()
		{
			return (byte)(m_pImpl.m_nSizeTwips / 15);
		}

		public Color GetColor(bool bCreateIfMissing)
		{
			if (m_pImpl.m_pColor == null && bCreateIfMissing)
				m_pImpl.m_pColor = new Color(0, 0, 0);
			return m_pImpl.m_pColor;
		}

		public void ClearColor()
		{
			{
			}
			m_pImpl.m_pColor = null;
		}

		public bool GetBold()
		{
			return m_pImpl.m_bBold;
		}

		public void SetBold(bool bBold)
		{
			m_pImpl.m_bBold = bBold;
		}

		public bool GetItalic()
		{
			return m_pImpl.m_bItalic;
		}

		public void SetItalic(bool bItalic)
		{
			m_pImpl.m_bItalic = bItalic;
		}

		public Font.Underline GetUnderline()
		{
			return m_pImpl.m_eUnderline;
		}

		public void SetUnderline(Font.Underline eUnderline)
		{
			if (eUnderline < Font.Underline.UNDERLINE_NONE || eUnderline >= Font.Underline.NUM_UNDERLINE)
				eUnderline = Font.Underline.UNDERLINE_NONE;
			m_pImpl.m_eUnderline = eUnderline;
		}

		~Font()
		{
		}

	}
}

namespace NumberDuck
{
	class Chart
	{
		public enum Type
		{
			TYPE_COLUMN = 0,
			TYPE_COLUMN_STACKED,
			TYPE_COLUMN_STACKED_100,
			TYPE_BAR,
			TYPE_BAR_STACKED,
			TYPE_BAR_STACKED_100,
			TYPE_LINE,
			TYPE_LINE_STACKED,
			TYPE_LINE_STACKED_100,
			TYPE_AREA,
			TYPE_AREA_STACKED,
			TYPE_AREA_STACKED_100,
			TYPE_SCATTER,
			NUM_TYPE,
		}

		public Secret.ChartImplementation m_pImpl;
		public Chart(Worksheet pWorksheet, Type eType)
		{
			m_pImpl = new Secret.ChartImplementation(pWorksheet, eType);
		}

		public uint GetX()
		{
			return m_pImpl.m_nX;
		}

		public bool SetX(uint nX)
		{
			m_pImpl.m_nX = nX;
			return true;
		}

		public uint GetY()
		{
			return m_pImpl.m_nY;
		}

		public bool SetY(uint nY)
		{
			m_pImpl.m_nY = nY;
			return true;
		}

		public uint GetSubX()
		{
			return m_pImpl.m_nSubX;
		}

		public void SetSubX(uint nSubX)
		{
			m_pImpl.m_nSubX = nSubX;
		}

		public uint GetSubY()
		{
			return m_pImpl.m_nSubY;
		}

		public void SetSubY(uint nSubY)
		{
			m_pImpl.m_nSubY = nSubY;
		}

		public uint GetWidth()
		{
			return m_pImpl.m_nWidth;
		}

		public void SetWidth(uint nWidth)
		{
			m_pImpl.m_nWidth = nWidth;
		}

		public uint GetHeight()
		{
			return m_pImpl.m_nHeight;
		}

		public void SetHeight(uint nHeight)
		{
			m_pImpl.m_nHeight = nHeight;
		}

		public new Chart.Type GetType()
		{
			return m_pImpl.m_eType;
		}

		public void SetType(Type eType)
		{
			if (eType >= Type.TYPE_COLUMN && eType < Type.NUM_TYPE)
				m_pImpl.m_eType = eType;
		}

		public uint GetNumSeries()
		{
			return (uint)(m_pImpl.m_pSeriesVector.GetSize());
		}

		public Series GetSeriesByIndex(uint nIndex)
		{
			if (nIndex >= GetNumSeries())
				return null;
			return m_pImpl.m_pSeriesVector.Get((int)(nIndex));
		}

		public Series CreateSeries(string szValues)
		{
			Secret.Formula pFormula = new Secret.Formula(szValues, m_pImpl.m_pWorksheet.m_pImpl);
			if (pFormula.ValidateForChart(m_pImpl.m_pWorksheet.m_pImpl))
			{
				NumberDuck.Secret.Formula __879619620 = pFormula;
				pFormula = null;
				{
					return m_pImpl.CreateSeries(__879619620);
				}
			}
			{
				pFormula = null;
			}
			{
				return null;
			}
		}

		public void PurgeSeries(uint nIndex)
		{
			m_pImpl.PurgeSeries((int)(nIndex));
		}

		public string GetCategories()
		{
			if (m_pImpl.m_pCategoriesFormula != null)
				return m_pImpl.m_pCategoriesFormula.ToChartString(m_pImpl.m_pWorksheet.m_pImpl);
			return "=";
		}

		public bool SetCategories(string szCategories)
		{
			Secret.Formula pFormula = new Secret.Formula(szCategories, m_pImpl.m_pWorksheet.m_pImpl);
			if (pFormula.ValidateForChart(m_pImpl.m_pWorksheet.m_pImpl))
			{
				{
				}
				{
					NumberDuck.Secret.Formula __879619620 = pFormula;
					pFormula = null;
					m_pImpl.m_pCategoriesFormula = __879619620;
				}
				{
					return true;
				}
			}
			{
				return false;
			}
		}

		public string GetTitle()
		{
			return m_pImpl.m_sTitle.GetExternalString();
		}

		public void SetTitle(string szTitle)
		{
			m_pImpl.m_sTitle.Set(szTitle);
		}

		public string GetHorizontalAxisLabel()
		{
			return m_pImpl.m_sHorizontalAxisLabel.GetExternalString();
		}

		public void SetHorizontalAxisLabel(string szHorizontalAxisLabel)
		{
			m_pImpl.m_sHorizontalAxisLabel.Set(szHorizontalAxisLabel);
		}

		public string GetVerticalAxisLabel()
		{
			return m_pImpl.m_sVerticalAxisLabel.GetExternalString();
		}

		public void SetVerticalAxisLabel(string szVerticalAxisLabel)
		{
			m_pImpl.m_sVerticalAxisLabel.Set(szVerticalAxisLabel);
		}

		public Legend GetLegend()
		{
			return m_pImpl.m_pLegend;
		}

		public Line GetFrameBorderLine()
		{
			return m_pImpl.m_pFrameBorderLine;
		}

		public Fill GetFrameFill()
		{
			return m_pImpl.m_pFrameFill;
		}

		public Line GetPlotBorderLine()
		{
			return m_pImpl.m_pPlotBorderLine;
		}

		public Fill GetPlotFill()
		{
			return m_pImpl.m_pPlotFill;
		}

		public Line GetHorizontalAxisLine()
		{
			return m_pImpl.m_pHorizontalAxisLine;
		}

		public Line GetHorizontalGridLine()
		{
			return m_pImpl.m_pHorizontalGridLine;
		}

		public Line GetVerticalAxisLine()
		{
			return m_pImpl.m_pVerticalAxisLine;
		}

		public Line GetVerticalGridLine()
		{
			return m_pImpl.m_pVerticalGridLine;
		}

		~Chart()
		{
		}

	}
}

namespace NumberDuck
{
	class Picture
	{
		public Secret.PictureImplementation m_pImplementation;
		public enum Format
		{
			FORMAT_PNG,
			FORMAT_JPEG,
			FORMAT_EMF,
			FORMAT_WMF,
			FORMAT_PICT,
			FORMAT_DIB,
			FORMAT_TIFF,
		}

		public Picture(Blob pBlob, Picture.Format eFormat)
		{
			m_pImplementation = new Secret.PictureImplementation(pBlob, eFormat);
		}

		public uint GetX()
		{
			return m_pImplementation.m_nX;
		}

		public bool SetX(uint nX)
		{
			m_pImplementation.m_nX = nX;
			return true;
		}

		public uint GetY()
		{
			return m_pImplementation.m_nY;
		}

		public bool SetY(uint nY)
		{
			m_pImplementation.m_nY = nY;
			return true;
		}

		public uint GetSubX()
		{
			return m_pImplementation.m_nSubX;
		}

		public void SetSubX(uint nSubX)
		{
			m_pImplementation.m_nSubX = nSubX;
		}

		public uint GetSubY()
		{
			return m_pImplementation.m_nSubY;
		}

		public void SetSubY(uint nSubY)
		{
			m_pImplementation.m_nSubY = nSubY;
		}

		public uint GetWidth()
		{
			return m_pImplementation.m_nWidth;
		}

		public void SetWidth(uint nWidth)
		{
			m_pImplementation.m_nWidth = nWidth;
		}

		public uint GetHeight()
		{
			return m_pImplementation.m_nHeight;
		}

		public void SetHeight(uint nHeight)
		{
			m_pImplementation.m_nHeight = nHeight;
		}

		public string GetUrl()
		{
			return m_pImplementation.m_sUrl.GetExternalString();
		}

		public void SetUrl(string szUrl)
		{
			m_pImplementation.m_sUrl.Set(szUrl);
		}

		public Blob GetBlob()
		{
			return m_pImplementation.m_pBlob;
		}

		public Picture.Format GetFormat()
		{
			return m_pImplementation.m_eFormat;
		}

		~Picture()
		{
		}

	}
}

namespace NumberDuck
{
	class Worksheet
	{
		public Secret.WorksheetImplementation m_pImpl;
		public const ushort MAX_COLUMN = 255;
		public const ushort MAX_ROW = 65535;
		public const ushort DEFAULT_COLUMN_WIDTH = 64;
		public const ushort DEFAULT_ROW_HEIGHT = 20;
		public enum Orientation
		{
			ORIENTATION_PORTRAIT,
			ORIENTATION_LANDSCAPE,
		}

		public Worksheet(Workbook pWorkbook)
		{
			m_pImpl = new Secret.WorksheetImplementation(pWorkbook, this);
		}

		~Worksheet()
		{
			{
				m_pImpl = null;
			}
		}

		public string GetName()
		{
			return m_pImpl.m_sName.GetExternalString();
		}

		public bool SetName(string szName)
		{
			for (ushort i = 0; i < m_pImpl.m_pWorkbook.GetNumWorksheet(); i++)
			{
				Worksheet pWorksheet = m_pImpl.m_pWorkbook.GetWorksheetByIndex(i);
				if (pWorksheet != this && Secret.ExternalString.Equal(pWorksheet.GetName(), szName))
					return false;
			}
			m_pImpl.m_sName.Set(szName);
			return true;
		}

		public Worksheet.Orientation GetOrientation()
		{
			return m_pImpl.m_eOrientation;
		}

		public void SetOrientation(Worksheet.Orientation eOrientation)
		{
			m_pImpl.m_eOrientation = eOrientation;
		}

		public bool GetPrintGridlines()
		{
			return m_pImpl.m_bPrintGridlines;
		}

		public void SetPrintGridlines(bool bPrintGridlines)
		{
			m_pImpl.m_bPrintGridlines = bPrintGridlines;
		}

		public bool GetShowGridlines()
		{
			return m_pImpl.m_bShowGridlines;
		}

		public void SetShowGridlines(bool bShowGridlines)
		{
			m_pImpl.m_bShowGridlines = bShowGridlines;
		}

		public ushort GetColumnWidth(ushort nColumn)
		{
			Secret.ColumnInfo pColumnInfo = m_pImpl.GetColumnInfo(nColumn);
			if (pColumnInfo != null)
				return pColumnInfo.m_nWidth;
			return DEFAULT_COLUMN_WIDTH;
		}

		public void SetColumnWidth(ushort nColumn, ushort nWidth)
		{
			if (nColumn > MAX_COLUMN)
				return;
			Secret.ColumnInfo pColumnInfo = m_pImpl.GetOrCreateColumnInfo(nColumn);
			pColumnInfo.m_nWidth = nWidth;
		}

		public bool GetColumnHidden(ushort nColumn)
		{
			Secret.ColumnInfo pColumnInfo = m_pImpl.GetColumnInfo(nColumn);
			if (pColumnInfo != null)
				return pColumnInfo.m_bHidden;
			return false;
		}

		public void SetColumnHidden(ushort nColumn, bool bHidden)
		{
			if (nColumn > MAX_COLUMN)
				return;
			Secret.ColumnInfo pColumnInfo = m_pImpl.GetOrCreateColumnInfo(nColumn);
			pColumnInfo.m_bHidden = bHidden;
		}

		public void InsertColumn(ushort nColumn)
		{
			if (nColumn > MAX_COLUMN)
				return;
			if (m_pImpl.m_pColumnInfoTable.GetSize() > 0)
			{
				int i = m_pImpl.m_pColumnInfoTable.GetSize() - 1;
				while (true)
				{
					Secret.TableElement<Secret.ColumnInfo> pElement = m_pImpl.m_pColumnInfoTable.GetByIndex(i);
					if (pElement.m_nColumn == MAX_COLUMN)
					{
						m_pImpl.m_pColumnInfoTable.Erase(i);
					}
					else if (pElement.m_nColumn >= nColumn)
					{
						{
							NumberDuck.Secret.ColumnInfo __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							m_pImpl.m_pColumnInfoTable.Set(pElement.m_nColumn + 1, pElement.m_nRow, __3920382863);
						}
						m_pImpl.m_pColumnInfoTable.Erase(i);
					}
					if (i == 0)
						break;
					i--;
				}
			}
			if (m_pImpl.m_pCellTable.GetSize() > 0)
			{
				int i = m_pImpl.m_pCellTable.GetSize() - 1;
				while (true)
				{
					Secret.TableElement<Cell> pElement = m_pImpl.m_pCellTable.GetByIndex(i);
					if (pElement.m_nColumn == MAX_COLUMN)
					{
						m_pImpl.m_pCellTable.Erase(i);
					}
					else if (pElement.m_nColumn >= nColumn)
					{
						{
							NumberDuck.Cell __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							m_pImpl.m_pCellTable.Set(pElement.m_nColumn + 1, pElement.m_nRow, __3920382863);
						}
						m_pImpl.m_pCellTable.Erase(i);
					}
					if (i == 0)
						break;
					i--;
				}
			}
			for (int i = 0; i < m_pImpl.m_pPictureVector.GetSize(); )
			{
				Picture pPicture = m_pImpl.m_pPictureVector.Get(i);
				if (pPicture.GetX() == MAX_COLUMN)
				{
					PurgePicture((ushort)(i));
					continue;
				}
				if (pPicture.GetX() >= nColumn)
					pPicture.SetX(pPicture.GetX() + 1);
				i++;
			}
			ushort nWorksheetIndex = 0;
			for (ushort i = 0; i < m_pImpl.m_pWorkbook.GetNumWorksheet(); i++)
			{
				if (this == m_pImpl.m_pWorkbook.GetWorksheetByIndex(i))
				{
					nWorksheetIndex = i;
					break;
				}
			}
			for (int i = 0; i < m_pImpl.m_pChartVector.GetSize(); )
			{
				Chart pChart = m_pImpl.m_pChartVector.Get(i);
				if (pChart.GetX() == MAX_COLUMN)
				{
					PurgeChart((ushort)(i));
					continue;
				}
				if (pChart.GetX() >= nColumn)
					pChart.SetX(pChart.GetX() + 1);
				pChart.m_pImpl.InsertColumn(nWorksheetIndex, nColumn);
				i++;
			}
			for (int i = 0; i < m_pImpl.m_pMergedCellVector.GetSize(); )
			{
				MergedCell pMergedCell = m_pImpl.m_pMergedCellVector.Get(i);
				if (pMergedCell.GetX() == MAX_COLUMN)
				{
					PurgeMergedCell((ushort)(i));
					continue;
				}
				if (pMergedCell.GetX() >= nColumn)
				{
					pMergedCell.SetX(pMergedCell.GetX() + 1);
					pMergedCell.SetWidth(pMergedCell.GetWidth());
				}
				else if (pMergedCell.GetX() + pMergedCell.GetWidth() >= nColumn)
					pMergedCell.SetWidth(pMergedCell.GetWidth() + 1);
				i++;
			}
		}

		public void DeleteColumn(ushort nColumn)
		{
			if (nColumn > MAX_COLUMN)
				return;
			{
				int i = 0;
				while (i < m_pImpl.m_pColumnInfoTable.GetSize())
				{
					Secret.TableElement<Secret.ColumnInfo> pElement = m_pImpl.m_pColumnInfoTable.GetByIndex(i);
					if (pElement.m_nColumn == nColumn)
					{
						m_pImpl.m_pColumnInfoTable.Erase(i);
					}
					else if (pElement.m_nColumn > nColumn)
					{
						int nTempColumn = pElement.m_nColumn - 1;
						int nTempRow = pElement.m_nRow;
						Secret.ColumnInfo pColumnInfo;
						{
							NumberDuck.Secret.ColumnInfo __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							pColumnInfo = __3920382863;
						}
						m_pImpl.m_pColumnInfoTable.Erase(i++);
						{
							NumberDuck.Secret.ColumnInfo __1173438266 = pColumnInfo;
							pColumnInfo = null;
							m_pImpl.m_pColumnInfoTable.Set(nTempColumn, nTempRow, __1173438266);
						}
					}
					else
					{
						i++;
					}
				}
			}
			{
				int i = 0;
				while (i < m_pImpl.m_pCellTable.GetSize())
				{
					Secret.TableElement<Cell> pElement = m_pImpl.m_pCellTable.GetByIndex(i);
					if (pElement.m_nColumn == nColumn)
					{
						m_pImpl.m_pCellTable.Erase(i);
					}
					else if (pElement.m_nColumn > nColumn)
					{
						int nTempColumn = pElement.m_nColumn - 1;
						int nTempRow = pElement.m_nRow;
						Cell pCell;
						{
							NumberDuck.Cell __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							pCell = __3920382863;
						}
						m_pImpl.m_pCellTable.Erase(i++);
						{
							NumberDuck.Cell __2223188566 = pCell;
							pCell = null;
							m_pImpl.m_pCellTable.Set(nTempColumn, nTempRow, __2223188566);
						}
					}
					else
					{
						i++;
					}
				}
			}
			for (int i = 0; i < m_pImpl.m_pPictureVector.GetSize(); )
			{
				Picture pPicture = m_pImpl.m_pPictureVector.Get(i);
				if (pPicture.GetX() == nColumn)
				{
					PurgePicture((ushort)(i));
					continue;
				}
				if (pPicture.GetX() > nColumn)
					pPicture.SetX(pPicture.GetX() - 1);
				i++;
			}
			ushort nWorksheetIndex = 0;
			for (ushort i = 0; i < m_pImpl.m_pWorkbook.GetNumWorksheet(); i++)
			{
				if (this == m_pImpl.m_pWorkbook.GetWorksheetByIndex(i))
				{
					nWorksheetIndex = i;
					break;
				}
			}
			for (int i = 0; i < m_pImpl.m_pChartVector.GetSize(); )
			{
				Chart pChart = m_pImpl.m_pChartVector.Get(i);
				if (pChart.GetX() == nColumn)
				{
					PurgeChart((ushort)(i));
					continue;
				}
				if (pChart.GetX() > nColumn)
					pChart.SetX(pChart.GetX() - 1);
				pChart.m_pImpl.DeleteColumn(nWorksheetIndex, nColumn);
				i++;
			}
			for (int i = 0; i < m_pImpl.m_pMergedCellVector.GetSize(); )
			{
				MergedCell pMergedCell = m_pImpl.m_pMergedCellVector.Get(i);
				if (pMergedCell.GetX() == nColumn)
				{
					PurgeMergedCell((ushort)(i));
					continue;
				}
				else if (pMergedCell.GetX() > nColumn)
					pMergedCell.SetX(pMergedCell.GetX() - 1);
				else if (pMergedCell.GetX() + pMergedCell.GetWidth() > nColumn)
					pMergedCell.SetWidth(pMergedCell.GetWidth() - 1);
				i++;
			}
		}

		public ushort GetRowHeight(ushort nRow)
		{
			Secret.RowInfo pRowInfo = m_pImpl.GetRowInfo(nRow);
			if (pRowInfo != null)
				return pRowInfo.m_nHeight;
			return DEFAULT_ROW_HEIGHT;
		}

		public void SetRowHeight(ushort nRow, ushort nHeight)
		{
			if (nRow > MAX_ROW)
				return;
			Secret.RowInfo pRowInfo = m_pImpl.GetOrCreateRowInfo(nRow);
			pRowInfo.m_nHeight = nHeight;
		}

		public void InsertRow(ushort nRow)
		{
			if (nRow > MAX_ROW)
				return;
			if (m_pImpl.m_pRowInfoTable.GetSize() > 0)
			{
				int i = m_pImpl.m_pRowInfoTable.GetSize() - 1;
				while (true)
				{
					Secret.TableElement<Secret.RowInfo> pElement = m_pImpl.m_pRowInfoTable.GetByIndex(i);
					if (pElement.m_nRow == MAX_ROW)
					{
						m_pImpl.m_pRowInfoTable.Erase(i);
					}
					else if (pElement.m_nRow >= nRow)
					{
						{
							NumberDuck.Secret.RowInfo __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							m_pImpl.m_pRowInfoTable.Set(pElement.m_nColumn, pElement.m_nRow + 1, __3920382863);
						}
						m_pImpl.m_pRowInfoTable.Erase(i);
					}
					if (i == 0)
						break;
					i--;
				}
			}
			if (m_pImpl.m_pCellTable.GetSize() > 0)
			{
				int i = m_pImpl.m_pCellTable.GetSize() - 1;
				while (true)
				{
					Secret.TableElement<Cell> pElement = m_pImpl.m_pCellTable.GetByIndex(i);
					if (pElement.m_nRow == MAX_ROW)
					{
						m_pImpl.m_pCellTable.Erase(i);
					}
					else if (pElement.m_nRow >= nRow)
					{
						{
							NumberDuck.Cell __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							m_pImpl.m_pCellTable.Set(pElement.m_nColumn, pElement.m_nRow + 1, __3920382863);
						}
						m_pImpl.m_pCellTable.Erase(i);
					}
					if (i == 0)
						break;
					i--;
				}
			}
			for (int i = 0; i < m_pImpl.m_pPictureVector.GetSize(); )
			{
				Picture pPicture = m_pImpl.m_pPictureVector.Get(i);
				if (pPicture.GetY() == MAX_ROW)
				{
					PurgePicture((ushort)(i));
					continue;
				}
				if (pPicture.GetY() >= nRow)
					pPicture.SetY(pPicture.GetY() + 1);
				i++;
			}
			ushort nWorksheetIndex = 0;
			for (ushort i = 0; i < m_pImpl.m_pWorkbook.GetNumWorksheet(); i++)
			{
				if (this == m_pImpl.m_pWorkbook.GetWorksheetByIndex(i))
				{
					nWorksheetIndex = i;
					break;
				}
			}
			for (int i = 0; i < m_pImpl.m_pChartVector.GetSize(); )
			{
				Chart pChart = m_pImpl.m_pChartVector.Get(i);
				if (pChart.GetY() == MAX_ROW)
				{
					PurgeChart((ushort)(i));
					continue;
				}
				if (pChart.GetY() >= nRow)
					pChart.SetY(pChart.GetY() + 1);
				pChart.m_pImpl.InsertRow(nWorksheetIndex, nRow);
				i++;
			}
			for (int i = 0; i < m_pImpl.m_pMergedCellVector.GetSize(); )
			{
				MergedCell pMergedCell = m_pImpl.m_pMergedCellVector.Get(i);
				if (pMergedCell.GetY() == MAX_ROW)
				{
					PurgeMergedCell((ushort)(i));
					continue;
				}
				if (pMergedCell.GetY() >= nRow)
				{
					pMergedCell.SetY(pMergedCell.GetY() + 1);
					pMergedCell.SetHeight(pMergedCell.GetHeight());
				}
				else if (pMergedCell.GetY() + pMergedCell.GetHeight() >= nRow)
					pMergedCell.SetHeight(pMergedCell.GetHeight() + 1);
				i++;
			}
		}

		public void DeleteRow(ushort nRow)
		{
			if (nRow > MAX_ROW)
				return;
			{
				int i = 0;
				while (i < m_pImpl.m_pRowInfoTable.GetSize())
				{
					Secret.TableElement<Secret.RowInfo> pElement = m_pImpl.m_pRowInfoTable.GetByIndex(i);
					if (pElement.m_nRow == nRow)
					{
						m_pImpl.m_pRowInfoTable.Erase(i);
					}
					else if (pElement.m_nRow > nRow)
					{
						int nTempColumn = pElement.m_nColumn;
						int nTempRow = pElement.m_nRow - 1;
						Secret.RowInfo pRowInfo;
						{
							NumberDuck.Secret.RowInfo __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							pRowInfo = __3920382863;
						}
						m_pImpl.m_pRowInfoTable.Erase(i++);
						{
							NumberDuck.Secret.RowInfo __3798332131 = pRowInfo;
							pRowInfo = null;
							m_pImpl.m_pRowInfoTable.Set(nTempColumn, nTempRow, __3798332131);
						}
					}
					else
					{
						i++;
					}
				}
			}
			{
				int i = 0;
				while (i < m_pImpl.m_pCellTable.GetSize())
				{
					Secret.TableElement<Cell> pElement = m_pImpl.m_pCellTable.GetByIndex(i);
					if (pElement.m_nRow == nRow)
					{
						m_pImpl.m_pCellTable.Erase(i);
					}
					else if (pElement.m_nRow > nRow)
					{
						int nTempColumn = pElement.m_nColumn;
						int nTempRow = pElement.m_nRow - 1;
						Cell pCell;
						{
							NumberDuck.Cell __3920382863 = pElement.m_xObject;
							pElement.m_xObject = null;
							pCell = __3920382863;
						}
						m_pImpl.m_pCellTable.Erase(i++);
						{
							NumberDuck.Cell __2223188566 = pCell;
							pCell = null;
							m_pImpl.m_pCellTable.Set(nTempColumn, nTempRow, __2223188566);
						}
					}
					else
					{
						i++;
					}
				}
			}
			for (int i = 0; i < m_pImpl.m_pPictureVector.GetSize(); )
			{
				Picture pPicture = m_pImpl.m_pPictureVector.Get(i);
				if (pPicture.GetY() == nRow)
				{
					PurgePicture((ushort)(i));
					continue;
				}
				if (pPicture.GetY() > nRow)
					pPicture.SetY(pPicture.GetY() - 1);
				i++;
			}
			ushort nWorksheetIndex = 0;
			for (ushort i = 0; i < m_pImpl.m_pWorkbook.GetNumWorksheet(); i++)
			{
				if (this == m_pImpl.m_pWorkbook.GetWorksheetByIndex(i))
				{
					nWorksheetIndex = i;
					break;
				}
			}
			for (int i = 0; i < m_pImpl.m_pChartVector.GetSize(); )
			{
				Chart pChart = m_pImpl.m_pChartVector.Get(i);
				if (pChart.GetY() == nRow)
				{
					PurgeChart((ushort)(i));
					continue;
				}
				if (pChart.GetY() > nRow)
					pChart.SetY(pChart.GetY() - 1);
				pChart.m_pImpl.DeleteRow(nWorksheetIndex, nRow);
				i++;
			}
			for (int i = 0; i < m_pImpl.m_pMergedCellVector.GetSize(); )
			{
				MergedCell pMergedCell = m_pImpl.m_pMergedCellVector.Get(i);
				if (pMergedCell.GetY() == nRow)
				{
					PurgeMergedCell((ushort)(i));
					continue;
				}
				else if (pMergedCell.GetY() > nRow)
					pMergedCell.SetY(pMergedCell.GetY() - 1);
				else if (pMergedCell.GetY() + pMergedCell.GetHeight() > nRow)
					pMergedCell.SetHeight(pMergedCell.GetHeight() - 1);
				i++;
			}
		}

		public Cell GetCell(ushort nX, ushort nY)
		{
			if (nX > MAX_COLUMN || nY > MAX_ROW)
				return null;
			Secret.TableElement<Cell> pElement = m_pImpl.m_pCellTable.GetOrCreate(nX, nY);
			if (pElement.m_xObject != null)
				return pElement.m_xObject;
			pElement.m_xObject = new Cell(this);
			return pElement.m_xObject;
		}

		public Cell GetCellByRC(ushort nRow, ushort nColumn)
		{
			if (nRow == 0 || nColumn == 0)
				return null;
			return GetCell((ushort)(nColumn - 1), (ushort)(nRow - 1));
		}

		public Cell GetCellByAddress(string szAddress)
		{
			Secret.Coordinate pCoordinate = Secret.WorksheetImplementation.AddressToCoordinate(szAddress);
			if (pCoordinate == null)
			{
				return null;
			}
			Cell pCell = GetCell(pCoordinate.m_nX, pCoordinate.m_nY);
			{
				pCoordinate = null;
			}
			{
				return pCell;
			}
		}

		public int GetNumPicture()
		{
			return m_pImpl.m_pPictureVector.GetSize();
		}

		public Picture GetPictureByIndex(int nIndex)
		{
			if (nIndex >= GetNumPicture())
				return null;
			return m_pImpl.m_pPictureVector.Get(nIndex);
		}

		public Picture CreatePicture(string szFileName)
		{
			Picture pPicture = null;
			Blob pBlob = new Blob(false);
			if (pBlob.Load(szFileName))
			{
				if (pPicture == null)
				{
					Secret.JpegLoader pJpegLoader = new Secret.JpegLoader();
					Secret.JpegImageInfo pImageInfo = pJpegLoader.Load(pBlob);
					if (pImageInfo != null)
					{
						Picture pOwnedPicture = new Picture(pBlob, Picture.Format.FORMAT_JPEG);
						pOwnedPicture.SetWidth((uint)(pImageInfo.m_nWidth));
						pOwnedPicture.SetHeight((uint)(pImageInfo.m_nHeight));
						pPicture = pOwnedPicture;
						{
							NumberDuck.Picture __1257297764 = pOwnedPicture;
							pOwnedPicture = null;
							m_pImpl.m_pPictureVector.PushBack(__1257297764);
						}
					}
					{
						pJpegLoader = null;
					}
				}
				if (pPicture == null)
				{
					Secret.PngLoader pPngLoader = new Secret.PngLoader();
					Secret.PngImageInfo pImageInfo = pPngLoader.Load(pBlob);
					if (pImageInfo != null)
					{
						Picture pOwnedPicture = new Picture(pBlob, Picture.Format.FORMAT_PNG);
						pOwnedPicture.SetWidth((uint)(pImageInfo.m_nWidth));
						pOwnedPicture.SetHeight((uint)(pImageInfo.m_nHeight));
						pPicture = pOwnedPicture;
						{
							NumberDuck.Picture __1257297764 = pOwnedPicture;
							pOwnedPicture = null;
							m_pImpl.m_pPictureVector.PushBack(__1257297764);
						}
					}
					{
						pPngLoader = null;
					}
				}
			}
			{
				pBlob = null;
			}
			{
				return pPicture;
			}
		}

		public void PurgePicture(int nIndex)
		{
			Picture pPicture = GetPictureByIndex(nIndex);
			if (pPicture != null)
				m_pImpl.m_pPictureVector.Erase(nIndex);
		}

		public int GetNumChart()
		{
			return m_pImpl.m_pChartVector.GetSize();
		}

		public Chart GetChartByIndex(int nIndex)
		{
			if (nIndex >= GetNumChart())
				return null;
			return m_pImpl.m_pChartVector.Get(nIndex);
		}

		public Chart CreateChart(Chart.Type eType)
		{
			Chart pOwnedChart = new Chart(this, eType);
			Chart pChart = pOwnedChart;
			{
				NumberDuck.Chart __1362309574 = pOwnedChart;
				pOwnedChart = null;
				m_pImpl.m_pChartVector.PushBack(__1362309574);
			}
			{
				return pChart;
			}
		}

		public void PurgeChart(int nIndex)
		{
			Chart pChart = GetChartByIndex(nIndex);
			if (pChart != null)
				m_pImpl.m_pChartVector.Erase(nIndex);
		}

		public int GetNumMergedCell()
		{
			return m_pImpl.m_pMergedCellVector.GetSize();
		}

		public MergedCell GetMergedCellByIndex(int nIndex)
		{
			if (nIndex >= GetNumMergedCell())
				return null;
			return m_pImpl.m_pMergedCellVector.Get(nIndex);
		}

		public MergedCell CreateMergedCell(ushort nX, ushort nY, ushort nWidth, ushort nHeight)
		{
			MergedCell pOwnedMergedCell = new MergedCell(nX, nY, nWidth, nHeight);
			MergedCell pMergedCell = pOwnedMergedCell;
			{
				NumberDuck.MergedCell __3844553686 = pOwnedMergedCell;
				pOwnedMergedCell = null;
				m_pImpl.m_pMergedCellVector.PushBack(__3844553686);
			}
			{
				return pMergedCell;
			}
		}

		public void PurgeMergedCell(int nIndex)
		{
			MergedCell pMergedCell = GetMergedCellByIndex(nIndex);
			if (pMergedCell != null)
				m_pImpl.m_pMergedCellVector.Erase(nIndex);
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtRecord
		{
			public const ushort MAX_DATA_SIZE = 8224;
			public enum Type
			{
				TYPE_OFFICE_ART_DGG_CONTAINER = 0xF000,
				TYPE_OFFICE_ART_B_STORE_CONTAINER = 0xF001,
				TYPE_OFFICE_ART_DG_CONTAINER = 0xF002,
				TYPE_OFFICE_ART_SPGR_CONTAINER = 0xF003,
				TYPE_OFFICE_ART_SP_CONTAINER = 0xF004,
				TYPE_OFFICE_ART_FDGG_BLOCK = 0xF006,
				TYPE_OFFICE_ART_FBSE = 0xF007,
				TYPE_OFFICE_ART_FDG = 0xF008,
				TYPE_OFFICE_ART_FSPGR = 0xF009,
				TYPE_OFFICE_ART_FSP = 0xF00A,
				TYPE_OFFICE_ART_FOPT = 0xF00B,
				TYPE_OFFICE_ART_TERTIARY_FOPT = 0xF122,
				TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET = 0xF010,
				TYPE_OFFICE_ART_CLIENT_DATA = 0xF011,
				TYPE_OFFICE_ART_BLIP_EMF = 0xF01A,
				TYPE_OFFICE_ART_BLIP_WMF = 0xF01B,
				TYPE_OFFICE_ART_BLIP_PICT = 0xF01C,
				TYPE_OFFICE_ART_BLIP_JPEG = 0xF01D,
				TYPE_OFFICE_ART_BLIP_PNG = 0xF01E,
				TYPE_OFFICE_ART_BLIP_DIB = 0xF01F,
				TYPE_OFFICE_ART_BLIP_TIFF = 0xF029,
				TYPE_OFFICE_ART_BLIP_JPEG_CMYK = 0xF02A,
				TYPE_OFFICE_ART_FRIT_CONTAINER = 0xF118,
				TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER = 0xF11E,
			}

			public enum OPIDType
			{
				OPID_PROTECTION_BOOLEAN_PROPERTIES = 0x007F,
				OPID_TEXT_BOOLEAN_PROPERTIES = 0x00BF,
				OPID_PIB = 0x0104,
				OPID_PIB_NAME = 0x0105,
				OPID_BLIP_BOOLEAN_PROPERTIES = 0x013F,
				OPID_FILL_COLOR = 0x0181,
				OPID_FILL_OPACITY = 0x0182,
				OPID_FILL_BACK_COLOR = 0x0183,
				OPID_FILL_BACK_OPACITY = 0x0184,
				OPID_FILL_STYLE_BOOLEAN_PROPERTIES = 0x01BF,
				OPID_LINE_COLOR = 0x01C0,
				OPID_LINE_STYLE_BOOLEAN_PROPERTIES = 0x01FF,
				OPID_SHADOW_STYLE_BOOLEAN_PROPERTIES = 0x023F,
				OPID_SHAPE_BOOLEAN_PROPERTIES = 0x033F,
				OPID_WZ_NAME = 0x0380,
				OPID_HYPERLINK = 0x0382,
				OPID_GROUP_SHAPE_BOOLEAN_PROPERTIES = 0x03BF,
			}

			protected OfficeArtRecordHeaderStruct m_pHeader;
			protected bool m_bIsContainer;
			protected OwnedVector<OfficeArtRecord> m_pOfficeArtRecordVector;
			public string GetTypeName()
			{
				switch ((Type)(m_pHeader.m_recType))
				{
					case Type.TYPE_OFFICE_ART_DGG_CONTAINER:
					{
						return "TYPE_OFFICE_ART_DGG_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_B_STORE_CONTAINER:
					{
						return "TYPE_OFFICE_ART_B_STORE_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_DG_CONTAINER:
					{
						return "TYPE_OFFICE_ART_DG_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_SPGR_CONTAINER:
					{
						return "TYPE_OFFICE_ART_SPGR_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_SP_CONTAINER:
					{
						return "TYPE_OFFICE_ART_SP_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_FDGG_BLOCK:
					{
						return "TYPE_OFFICE_ART_FDGG_BLOCK";
					}

					case Type.TYPE_OFFICE_ART_FBSE:
					{
						return "TYPE_OFFICE_ART_FBSE";
					}

					case Type.TYPE_OFFICE_ART_FDG:
					{
						return "TYPE_OFFICE_ART_FDG";
					}

					case Type.TYPE_OFFICE_ART_FSPGR:
					{
						return "TYPE_OFFICE_ART_FSPGR";
					}

					case Type.TYPE_OFFICE_ART_FSP:
					{
						return "TYPE_OFFICE_ART_FSP";
					}

					case Type.TYPE_OFFICE_ART_FOPT:
					{
						return "TYPE_OFFICE_ART_FOPT";
					}

					case Type.TYPE_OFFICE_ART_TERTIARY_FOPT:
					{
						return "TYPE_OFFICE_ART_TERTIARY_FOPT";
					}

					case Type.TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET:
					{
						return "TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET";
					}

					case Type.TYPE_OFFICE_ART_CLIENT_DATA:
					{
						return "TYPE_OFFICE_ART_CLIENT_DATA";
					}

					case Type.TYPE_OFFICE_ART_BLIP_EMF:
					{
						return "TYPE_OFFICE_ART_BLIP_EMF";
					}

					case Type.TYPE_OFFICE_ART_BLIP_WMF:
					{
						return "TYPE_OFFICE_ART_BLIP_WMF";
					}

					case Type.TYPE_OFFICE_ART_BLIP_PICT:
					{
						return "TYPE_OFFICE_ART_BLIP_PICT";
					}

					case Type.TYPE_OFFICE_ART_BLIP_JPEG:
					{
						return "TYPE_OFFICE_ART_BLIP_JPEG";
					}

					case Type.TYPE_OFFICE_ART_BLIP_PNG:
					{
						return "TYPE_OFFICE_ART_BLIP_PNG";
					}

					case Type.TYPE_OFFICE_ART_BLIP_DIB:
					{
						return "TYPE_OFFICE_ART_BLIP_DIB";
					}

					case Type.TYPE_OFFICE_ART_BLIP_TIFF:
					{
						return "TYPE_OFFICE_ART_BLIP_TIFF";
					}

					case Type.TYPE_OFFICE_ART_BLIP_JPEG_CMYK:
					{
						return "TYPE_OFFICE_ART_BLIP_JPEG_CMYK";
					}

					case Type.TYPE_OFFICE_ART_FRIT_CONTAINER:
					{
						return "TYPE_OFFICE_ART_FRIT_CONTAINER";
					}

					case Type.TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER:
					{
						return "TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER";
					}

				}
				return "???";
			}

			public OfficeArtRecord(OfficeArtRecordHeaderStruct pHeader, bool bIsContainer, BlobView pBlobView)
			{
				m_pHeader = pHeader;
				m_bIsContainer = bIsContainer;
				m_pOfficeArtRecordVector = new OwnedVector<OfficeArtRecord>();
				if (m_bIsContainer)
				{
					nbAssert.Assert(m_pHeader.m_recVer == 0xF);
					while (pBlobView.GetOffset() < (int)(m_pHeader.m_recLen) && pBlobView.GetOffset() < pBlobView.GetSize())
					{
						OfficeArtRecord pOfficeArtRecord = CreateOfficeArtRecord(pBlobView);
						{
							NumberDuck.Secret.OfficeArtRecord __3533451309 = pOfficeArtRecord;
							pOfficeArtRecord = null;
							m_pOfficeArtRecordVector.PushBack(__3533451309);
						}
					}
				}
			}

			public OfficeArtRecord(Type nType, bool bIsContainer, uint nSize, bool bHax)
			{
				m_pHeader = new OfficeArtRecordHeaderStruct();
				m_pHeader.m_recType = (ushort)(nType);
				m_pHeader.m_recLen = nSize;
				m_bIsContainer = bIsContainer;
				m_pOfficeArtRecordVector = new OwnedVector<OfficeArtRecord>();
			}

			public ushort GetVersion()
			{
				return m_pHeader.m_recVer;
			}

			public ushort GetInstance()
			{
				return m_pHeader.m_recInstance;
			}

			public new OfficeArtRecord.Type GetType()
			{
				return (Type)(m_pHeader.m_recType);
			}

			public bool GetIsContainer()
			{
				return m_bIsContainer;
			}

			public uint GetSize()
			{
				return m_pHeader.m_recLen + OfficeArtRecordHeaderStruct.SIZE;
			}

			public virtual uint GetRecursiveSize()
			{
				uint nSize = OfficeArtRecordHeaderStruct.SIZE;
				if (m_bIsContainer)
				{
					for (int i = 0; i < m_pOfficeArtRecordVector.GetSize(); i++)
						nSize = nSize + m_pOfficeArtRecordVector.Get(i).GetRecursiveSize();
				}
				else
				{
					nSize = nSize + m_pHeader.m_recLen;
				}
				return nSize;
			}

			public virtual void RecursiveWrite(BlobView pBlobView)
			{
				int nBefore = pBlobView.GetOffset();
				m_pHeader.BlobWrite(pBlobView);
				nbAssert.Assert(pBlobView.GetOffset() - nBefore == OfficeArtRecordHeaderStruct.SIZE);
				if (m_bIsContainer)
				{
					for (int i = 0; i < m_pOfficeArtRecordVector.GetSize(); i++)
						m_pOfficeArtRecordVector.Get(i).RecursiveWrite(pBlobView);
				}
				else
				{
					BlobWrite(pBlobView);
				}
				nbAssert.Assert(pBlobView.GetOffset() - nBefore == (int)(OfficeArtRecordHeaderStruct.SIZE + m_pHeader.m_recLen));
			}

			public virtual void BlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public virtual void BlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public ushort GetNumOfficeArtRecord()
			{
				nbAssert.Assert(m_bIsContainer);
				return (ushort)(m_pOfficeArtRecordVector.GetSize());
			}

			public OfficeArtRecord GetOfficeArtRecordByIndex(ushort nIndex)
			{
				nbAssert.Assert(m_bIsContainer);
				nbAssert.Assert(nIndex < m_pOfficeArtRecordVector.GetSize());
				return m_pOfficeArtRecordVector.Get(nIndex);
			}

			public OfficeArtRecord FindOfficeArtRecordByType(Type eType)
			{
				nbAssert.Assert(m_bIsContainer);
				for (int i = 0; i < m_pOfficeArtRecordVector.GetSize(); i++)
				{
					OfficeArtRecord pOfficeArtRecord = m_pOfficeArtRecordVector.Get(i);
					if (pOfficeArtRecord.GetType() == eType)
						return pOfficeArtRecord;
					if (pOfficeArtRecord.GetIsContainer())
					{
						OfficeArtRecord pTemp = pOfficeArtRecord.FindOfficeArtRecordByType(eType);
						if (pTemp != null)
							return pTemp;
					}
				}
				return null;
			}

			public virtual void AddOfficeArtRecord(OfficeArtRecord pOfficeArtRecord)
			{
				nbAssert.Assert(m_bIsContainer);
				m_pHeader.m_recLen += pOfficeArtRecord.GetRecursiveSize();
				m_pOfficeArtRecordVector.PushBack(pOfficeArtRecord);
			}

			public static OfficeArtRecord CreateOfficeArtRecord(BlobView pBlobView)
			{
				nbAssert.Assert(pBlobView.GetOffset() < pBlobView.GetSize());
				if (pBlobView.GetSize() - pBlobView.GetOffset() < (int)(OfficeArtRecordHeaderStruct.SIZE))
					return null;
				OfficeArtRecordHeaderStruct pHeader = new OfficeArtRecordHeaderStruct();
				pHeader.BlobRead(pBlobView);
				if ((Type)(pHeader.m_recType) != Type.TYPE_OFFICE_ART_DG_CONTAINER && (Type)(pHeader.m_recType) != Type.TYPE_OFFICE_ART_SPGR_CONTAINER)
					if (pBlobView.GetOffset() + (int)(pHeader.m_recLen) > pBlobView.GetSize())
					{
						{
							return null;
						}
					}
				switch ((Type)(pHeader.m_recType))
				{
					case Type.TYPE_OFFICE_ART_DGG_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtDggContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_B_STORE_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtBStoreContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_DG_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtDgContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_SPGR_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtSpgrContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_SP_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtSpContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FDGG_BLOCK:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFDGGBlockRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FBSE:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFBSERecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FDG:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFDGRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FSPGR:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFSPGRRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FSP:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFSPRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FOPT:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFOPTRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_TERTIARY_FOPT:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtTertiaryFOPTRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtClientAnchorSheetRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_CLIENT_DATA:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtClientDataRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_BLIP_EMF:
					case Type.TYPE_OFFICE_ART_BLIP_WMF:
					case Type.TYPE_OFFICE_ART_BLIP_PICT:
					case Type.TYPE_OFFICE_ART_BLIP_JPEG:
					case Type.TYPE_OFFICE_ART_BLIP_PNG:
					case Type.TYPE_OFFICE_ART_BLIP_DIB:
					case Type.TYPE_OFFICE_ART_BLIP_TIFF:
					case Type.TYPE_OFFICE_ART_BLIP_JPEG_CMYK:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtBlipRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_FRIT_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtFRITContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					case Type.TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtSplitMenuColorContainerRecord(__1199093386, pBlobView);
							}
						}
					}

					default:
					{
						{
							NumberDuck.Secret.OfficeArtRecordHeaderStruct __1199093386 = pHeader;
							pHeader = null;
							{
								return new OfficeArtRecord(__1199093386, false, pBlobView);
							}
						}
					}

				}
			}

			~OfficeArtRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ParsedExpressionRecord
		{
			public enum Type
			{
				TYPE_PtgExp,
				TYPE_PtgAdd,
				TYPE_PtgSub,
				TYPE_PtgMul,
				TYPE_PtgDiv,
				TYPE_PtgPower,
				TYPE_PtgConcat,
				TYPE_PtgLt,
				TYPE_PtgLe,
				TYPE_PtgEq,
				TYPE_PtgGe,
				TYPE_PtgGt,
				TYPE_PtgNe,
				TYPE_PtgParen,
				TYPE_PtgMissArg,
				TYPE_PtgStr,
				TYPE_PtgAttrSemi,
				TYPE_PtgAttrIf,
				TYPE_PtgAttrGoto,
				TYPE_PtgAttrSum,
				TYPE_PtgAttrSpace,
				TYPE_PtgBool,
				TYPE_PtgInt,
				TYPE_PtgNum,
				TYPE_PtgArea,
				TYPE_PtgArea3d,
				TYPE_PtgFunc,
				TYPE_PtgFuncVar,
				TYPE_PtgRef,
				TYPE_PtgRef3d,
				TYPE_UNKNOWN,
			}

			protected Type m_eType;
			protected ushort m_nSize;
			protected byte m_nFirstByte;
			protected byte m_nSecondByte;
			public ParsedExpressionRecord(Type eType, ushort nSize, bool bHax)
			{
				m_eType = eType;
				m_nSize = nSize;
				m_nFirstByte = 0;
				m_nSecondByte = 0;
			}

			~ParsedExpressionRecord()
			{
			}

			public virtual Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return null;
			}

			public new ParsedExpressionRecord.Type GetType()
			{
				return m_eType;
			}

			public string GetTypeName()
			{
				switch ((Type)(m_eType))
				{
					case Type.TYPE_PtgExp:
					{
						return "TYPE_PtgExp";
					}

					case Type.TYPE_PtgAdd:
					{
						return "TYPE_PtgAdd";
					}

					case Type.TYPE_PtgSub:
					{
						return "TYPE_PtgSub";
					}

					case Type.TYPE_PtgMul:
					{
						return "TYPE_PtgMul";
					}

					case Type.TYPE_PtgDiv:
					{
						return "TYPE_PtgDiv";
					}

					case Type.TYPE_PtgPower:
					{
						return "TYPE_PtgPower";
					}

					case Type.TYPE_PtgConcat:
					{
						return "TYPE_PtgConcat";
					}

					case Type.TYPE_PtgLt:
					{
						return "TYPE_PtgLt";
					}

					case Type.TYPE_PtgLe:
					{
						return "TYPE_PtgLe";
					}

					case Type.TYPE_PtgEq:
					{
						return "TYPE_PtgEq";
					}

					case Type.TYPE_PtgGe:
					{
						return "TYPE_PtgGe";
					}

					case Type.TYPE_PtgGt:
					{
						return "TYPE_PtgGt";
					}

					case Type.TYPE_PtgNe:
					{
						return "TYPE_PtgNe";
					}

					case Type.TYPE_PtgParen:
					{
						return "TYPE_PtgParen";
					}

					case Type.TYPE_PtgMissArg:
					{
						return "TYPE_PtgMissArg";
					}

					case Type.TYPE_PtgStr:
					{
						return "TYPE_PtgStr";
					}

					case Type.TYPE_PtgAttrSemi:
					{
						return "TYPE_PtgAttrSemi";
					}

					case Type.TYPE_PtgAttrIf:
					{
						return "TYPE_PtgAttrIf";
					}

					case Type.TYPE_PtgAttrGoto:
					{
						return "TYPE_PtgAttrGoto";
					}

					case Type.TYPE_PtgAttrSum:
					{
						return "TYPE_PtgAttrSum";
					}

					case Type.TYPE_PtgAttrSpace:
					{
						return "TYPE_PtgAttrSpace";
					}

					case Type.TYPE_PtgArea:
					{
						return "TYPE_PtgArea";
					}

					case Type.TYPE_PtgBool:
					{
						return "TYPE_PtgBool";
					}

					case Type.TYPE_PtgInt:
					{
						return "TYPE_PtgInt";
					}

					case Type.TYPE_PtgNum:
					{
						return "TYPE_PtgNum";
					}

					case Type.TYPE_PtgArea3d:
					{
						return "TYPE_PtgArea3d";
					}

					case Type.TYPE_PtgFunc:
					{
						return "TYPE_PtgFunc";
					}

					case Type.TYPE_PtgFuncVar:
					{
						return "TYPE_PtgFuncVar";
					}

					case Type.TYPE_PtgRef:
					{
						return "TYPE_PtgRef";
					}

					case Type.TYPE_PtgRef3d:
					{
						return "TYPE_PtgRef3d";
					}

					case Type.TYPE_UNKNOWN:
					{
						return "TYPE_UNKNOWN";
					}

				}
				return "???";
			}

			public ushort GetDataSize()
			{
				return m_nSize;
			}

			public virtual void BlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public virtual void BlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public static ParsedExpressionRecord CreateParsedExpressionRecord(BlobView pBlobView)
			{
				int nInitialOffset = pBlobView.GetOffset();
				byte ptg = pBlobView.UnpackUint8();
				pBlobView.SetOffset(nInitialOffset);
				ParsedExpressionRecord pParsedExpressionRecord = null;
				switch (ptg)
				{
					case 0x03:
					case 0x04:
					case 0x05:
					case 0x06:
					case 0x07:
					case 0x08:
					case 0x09:
					case 0x0A:
					case 0x0B:
					case 0x0C:
					case 0x0D:
					case 0x0E:
					{
						pParsedExpressionRecord = new PtgOperatorRecord(pBlobView);
						break;
					}

					case 0x15:
					{
						pParsedExpressionRecord = new PtgParenRecord(pBlobView);
						break;
					}

					case 0x16:
					{
						pParsedExpressionRecord = new PtgMissArgRecord(pBlobView);
						break;
					}

					case 0x17:
					{
						pParsedExpressionRecord = new PtgStrRecord(pBlobView);
						break;
					}

					case 0x19:
					{
						pBlobView.SetOffset(nInitialOffset + 1);
						byte eptg = pBlobView.UnpackUint8();
						pBlobView.SetOffset(nInitialOffset);
						switch (eptg)
						{
							case 0x01:
							{
								pParsedExpressionRecord = new PtgAttrSemiRecord(pBlobView);
								break;
							}

							case 0x02:
							{
								pParsedExpressionRecord = new PtgAttrIfRecord(pBlobView);
								break;
							}

							case 0x08:
							{
								pParsedExpressionRecord = new PtgAttrGotoRecord(pBlobView);
								break;
							}

							case 0x10:
							{
								pParsedExpressionRecord = new PtgAttrSumRecord(pBlobView);
								break;
							}

							case 0x40:
							{
								pParsedExpressionRecord = new PtgAttrSpaceRecord(pBlobView);
								break;
							}

						}
						break;
					}

					case 0x1d:
					{
						pParsedExpressionRecord = new PtgBoolRecord(pBlobView);
						break;
					}

					case 0x1e:
					{
						pParsedExpressionRecord = new PtgIntRecord(pBlobView);
						break;
					}

					case 0x1f:
					{
						pParsedExpressionRecord = new PtgNumRecord(pBlobView);
						break;
					}

					case 0x22:
					{
						pParsedExpressionRecord = new PtgFuncVarRecord(pBlobView);
						break;
					}

					case 0x24:
					{
						pParsedExpressionRecord = new PtgRefRecord(pBlobView);
						break;
					}

					case 0x25:
					{
						pParsedExpressionRecord = new PtgAreaRecord(pBlobView);
						break;
					}

					case 0x3A:
					{
						pParsedExpressionRecord = new PtgRef3dRecord(pBlobView);
						break;
					}

					case 0x3B:
					{
						pParsedExpressionRecord = new PtgArea3dRecord(pBlobView);
						break;
					}

					case 0x41:
					{
						pParsedExpressionRecord = new PtgFuncRecord(pBlobView);
						break;
					}

					case 0x42:
					{
						pParsedExpressionRecord = new PtgFuncVarRecord(pBlobView);
						break;
					}

					case 0x44:
					{
						pParsedExpressionRecord = new PtgRefRecord(pBlobView);
						break;
					}

					case 0x5A:
					{
						pParsedExpressionRecord = new PtgRef3dRecord(pBlobView);
						break;
					}

					default:
					{
						pParsedExpressionRecord = new ParsedExpressionRecord(Type.TYPE_UNKNOWN, 0, true);
						break;
					}

				}
				{
					NumberDuck.Secret.ParsedExpressionRecord __3596419756 = pParsedExpressionRecord;
					pParsedExpressionRecord = null;
					return __3596419756;
				}
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffHeader
		{
			public ushort m_nType;
			public uint m_nSize;
		}
		class BiffRecord
		{
			public const ushort MAX_DATA_SIZE = 8224;
			public enum Type
			{
				TYPE_DIMENSIONS_B2 = 0x0000,
				TYPE_BLANK_B2 = 0x0001,
				TYPE_INTEGER_B2_ONLY = 0x0002,
				TYPE_NUMBER_B2 = 0x0003,
				TYPE_LABEL_B2 = 0x0004,
				TYPE_BOOLERR_B2 = 0x0005,
				TYPE_FORMULA = 0x0006,
				TYPE_STRING_B2 = 0x0007,
				TYPE_ROW_B2 = 0x0008,
				TYPE_BOF_B2 = 0x0009,
				TYPE_EOF = 0x000A,
				TYPE_INDEX_B2_ONLY = 0x000B,
				TYPE_CALCCOUNT = 0x000C,
				TYPE_CALCMODE = 0x000D,
				TYPE_CALC_PRECISION = 0x000E,
				TYPE_REFMODE = 0x000F,
				TYPE_DELTA = 0x0010,
				TYPE_ITERATION = 0x0011,
				TYPE_PROTECT = 0x0012,
				TYPE_PASSWORD = 0x0013,
				TYPE_HEADER = 0x0014,
				TYPE_FOOTER = 0x0015,
				TYPE_EXTERNCOUNT = 0x0016,
				TYPE_EXTERN_SHEET = 0x0017,
				TYPE_NAME_B2 = 0x0018,
				TYPE_WIN_PROTECT = 0x0019,
				TYPE_VERTICALPAGEBREAKS = 0x001A,
				TYPE_HORIZONTALPAGEBREAKS = 0x001B,
				TYPE_NOTE = 0x001C,
				TYPE_SELECTION = 0x001D,
				TYPE_FORMAT_B2 = 0x001E,
				TYPE_BUILTINFMTCOUNT_B2 = 0x001F,
				TYPE_COLUMNDEFAULT_B2_ONLY = 0x0020,
				TYPE_ARRAY_B2_ONLY = 0x0021,
				TYPE_DATE_1904 = 0x0022,
				TYPE_EXTERN_NAME = 0x0023,
				TYPE_COLWIDTH_B2_ONLY = 0x0024,
				TYPE_DEFAULTROWHEIGHT_B2_ONLY = 0x0025,
				TYPE_LEFT_MARGIN = 0x0026,
				TYPE_RIGHT_MARGIN = 0x0027,
				TYPE_TOP_MARGIN = 0x0028,
				TYPE_BOTTOM_MARGIN = 0x0029,
				TYPE_PrintRowCol = 0x002A,
				TYPE_PrintGrid = 0x002B,
				TYPE_FILEPASS = 0x002F,
				TYPE_FONT = 0x0031,
				TYPE_FONT2_B2_ONLY = 0x0032,
				TYPE_PRINT_SIZE = 0x0033,
				TYPE_TABLEOP_B2 = 0x0036,
				TYPE_TABLEOP2_B2 = 0x0037,
				TYPE_CONTINUE = 0x003C,
				TYPE_WINDOW1 = 0x003D,
				TYPE_WINDOW2_B2 = 0x003E,
				TYPE_BACKUP = 0x0040,
				TYPE_PANE = 0x0041,
				TYPE_CODE_PAGE = 0x0042,
				TYPE_XF_B2 = 0x0043,
				TYPE_IXFE_B2_ONLY = 0x0044,
				TYPE_EFONT_B2_ONLY = 0x0045,
				TYPE_PLS = 0x004D,
				TYPE_DCONREF = 0x0051,
				TYPE_DEFCOLWIDTH = 0x0055,
				TYPE_BUILTINFMTCOUNT_B3 = 0x0056,
				TYPE_XCT = 0x0059,
				TYPE_CRN = 0x005A,
				TYPE_FILESHARING = 0x005B,
				TYPE_WRITE_ACCESS = 0x005C,
				TYPE_OBJ = 0x005D,
				TYPE_UNCALCED = 0x005E,
				TYPE_SAVERECALC = 0x005F,
				TYPE_OBJECTPROTECT = 0x0063,
				TYPE_COLINFO = 0x007D,
				TYPE_RK2_mythical = 0x007E,
				TYPE_GUTS = 0x0080,
				TYPE_WSBOOL = 0x0081,
				TYPE_GRIDSET = 0x0082,
				TYPE_HCENTER = 0x0083,
				TYPE_VCENTER = 0x0084,
				TYPE_BOUND_SHEET_8 = 0x0085,
				TYPE_WRITEPROT = 0x0086,
				TYPE_COUNTRY = 0x008C,
				TYPE_HIDE_OBJ = 0x008D,
				TYPE_SHEETSOFFSET = 0x008E,
				TYPE_SHEETHDR = 0x008F,
				TYPE_SORT = 0x0090,
				TYPE_PALETTE = 0x0092,
				TYPE_STANDARDWIDTH = 0x0099,
				TYPE_FILTERMODE = 0x009B,
				TYPE_BUILT_IN_FN_GROUP_COUNT = 0x009C,
				TYPE_AUTOFILTERINFO = 0x009D,
				TYPE_AUTOFILTER = 0x009E,
				TYPE_SCL = 0x00A0,
				TYPE_SETUP = 0x00A1,
				TYPE_GCW = 0x00AB,
				TYPE_MULRK = 0x00BD,
				TYPE_MULBLANK = 0x00BE,
				TYPE_MMS = 0x00C1,
				TYPE_RSTRING = 0x00D6,
				TYPE_DBCELL = 0x00D7,
				TYPE_BOOK_BOOL = 0x00DA,
				TYPE_SCENPROTECT = 0x00DD,
				TYPE_XF = 0x00E0,
				TYPE_INTERFACE_HDR = 0x00E1,
				TYPE_INTERFACE_END = 0x00E2,
				TYPE_MergeCells = 0x00E5,
				TYPE_BITMAP = 0x00E9,
				TYPE_MSO_DRAWING_GROUP = 0x00EB,
				TYPE_MSO_DRAWING = 0x00EC,
				TYPE_MSO_DRAWING_SELECTION = 0x00ED,
				TYPE_PHONETIC = 0x00EF,
				TYPE_SST = 0x00FC,
				TYPE_LABELSST = 0x00FD,
				TYPE_EXT_SST = 0x00FF,
				TYPE_RR_TAB_ID = 0x013D,
				TYPE_LABELRANGES = 0x015F,
				TYPE_USESELFS = 0x0160,
				TYPE_DSF = 0x0161,
				TYPE_SUP_BOOK = 0x01AE,
				TYPE_PROT_4_REV = 0x01AF,
				TYPE_CONDFMT = 0x01B0,
				TYPE_CF = 0x01B1,
				TYPE_DVAL = 0x01B2,
				TYPE_TXO = 0x01B6,
				TYPE_REFRESH_ALL = 0x01B7,
				TYPE_HLINK = 0x01B8,
				TYPE_PROT_4_REV_PASS = 0x01BC,
				TYPE_DV = 0x01BE,
				TYPE_EXCEL9_FILE = 0x01C0,
				TYPE_RECALCID = 0x01C1,
				TYPE_DIMENSION = 0x0200,
				TYPE_BLANK = 0x0201,
				TYPE_NUMBER = 0x0203,
				TYPE_LABEL = 0x0204,
				TYPE_BOOLERR = 0x0205,
				TYPE_FORMULA_B3 = 0x0206,
				TYPE_STRING = 0x0207,
				TYPE_ROW = 0x0208,
				TYPE_INDEX_B3 = 0x020B,
				TYPE_NAME = 0x0218,
				TYPE_ARRAY = 0x0221,
				TYPE_EXTERNNAME_B3 = 0x0223,
				TYPE_DEFAULTROWHEIGHT = 0x0225,
				TYPE_FONT_B3B4 = 0x0231,
				TYPE_TABLEOP = 0x0236,
				TYPE_WINDOW2 = 0x023E,
				TYPE_XF_B3 = 0x0243,
				TYPE_RK = 0x027E,
				TYPE_STYLE = 0x0293,
				TYPE_STYLE_EXT = 0x0892,
				TYPE_FORMULA_B4 = 0x0406,
				TYPE_FORMAT = 0x041E,
				TYPE_XF_B4 = 0x0443,
				TYPE_SHRFMLA = 0x04BC,
				TYPE_QUICKTIP = 0x0800,
				TYPE_BOF = 0x0809,
				TYPE_SHEETLAYOUT = 0x0862,
				TYPE_SHEETPROTECTION = 0x0867,
				TYPE_RANGEPROTECTION = 0x0868,
				TYPE_XF_CRC = 0x087C,
				TYPE_XF_EXT = 0x087D,
				TYPE_BOOK_EXT = 0x0863,
				TYPE_THEME = 0x0896,
				TYPE_COMPRESS_PICTURES = 0x089B,
				TYPE_PLV = 0x088B,
				TYPE_COMPAT12 = 0x088C,
				TYPE_MTR_SETTINGS = 0x089A,
				TYPE_FORCE_FULL_CALCULATION = 0x08A3,
				TYPE_TABLE_STYLES = 0x088E,
				TYPE_ChartColors = 684,
				TYPE_ChartFrtInfo = 2128,
				TYPE_StartBlock = 2130,
				TYPE_EndBlock = 2131,
				TYPE_StartObject = 2132,
				TYPE_EndObject = 2133,
				TYPE_CatLab = 2134,
				TYPE_YMult = 2135,
				TYPE_FrtFontList = 2138,
				TYPE_DataLabExt = 2154,
				TYPE_DataLabExtContents = 2155,
				TYPE_NamePublish = 2195,
				TYPE_HeaderFooter = 2204,
				TYPE_CrtMlFrt = 2206,
				TYPE_ShapePropsStream = 2212,
				TYPE_TextPropsStream = 2213,
				TYPE_CrtLayout12A = 2215,
				TYPE_Units = 4097,
				TYPE_Chart = 4098,
				TYPE_Series = 4099,
				TYPE_DataFormat = 4102,
				TYPE_LineFormat = 4103,
				TYPE_MarkerFormat = 4105,
				TYPE_AreaFormat = 4106,
				TYPE_PieFormat = 4107,
				TYPE_AttachedLabel = 4108,
				TYPE_SeriesText = 4109,
				TYPE_ChartFormat = 4116,
				TYPE_Legend = 4117,
				TYPE_SeriesList = 4118,
				TYPE_Bar = 4119,
				TYPE_Line = 4120,
				TYPE_Pie = 4121,
				TYPE_Area = 4122,
				TYPE_Scatter = 4123,
				TYPE_CrtLine = 4124,
				TYPE_Axis = 4125,
				TYPE_Tick = 4126,
				TYPE_ValueRange = 4127,
				TYPE_CatSerRange = 4128,
				TYPE_AxisLine = 4129,
				TYPE_CrtLink = 4130,
				TYPE_DefaultText = 4132,
				TYPE_Text = 4133,
				TYPE_FontX = 4134,
				TYPE_ObjectLink = 4135,
				TYPE_Frame = 4146,
				TYPE_Begin = 4147,
				TYPE_End = 4148,
				TYPE_PlotArea = 4149,
				TYPE_Chart3d = 4154,
				TYPE_PicF = 4156,
				TYPE_DropBar = 4157,
				TYPE_Radar = 4158,
				TYPE_Surf = 4159,
				TYPE_RadarArea = 4160,
				TYPE_AxisParent = 4161,
				TYPE_LegendException = 4163,
				TYPE_ShtProps = 4164,
				TYPE_SerToCrt = 4165,
				TYPE_AxesUsed = 4166,
				TYPE_SerParent = 4170,
				TYPE_SerAuxTrend = 4171,
				TYPE_IFmtRecord = 4174,
				TYPE_Pos = 4175,
				TYPE_AlRuns = 4176,
				TYPE_BRAI = 4177,
				TYPE_BOFDatasheet = 4178,
				TYPE_ExcludeRows = 4179,
				TYPE_ExcludeColumns = 4180,
				TYPE_Orient = 4181,
				TYPE_WinDoc = 4183,
				TYPE_MaxStatus = 4184,
				TYPE_MainWindow = 4185,
				TYPE_SerAuxErrBar = 4187,
				TYPE_ClrtClient = 4188,
				TYPE_SerFmt = 4189,
				TYPE_LinkedSelection = 4190,
				TYPE_Chart3DBarShape = 4191,
				TYPE_Fbi = 4192,
				TYPE_BopPop = 4193,
				TYPE_AxcExt = 4194,
				TYPE_Dat = 4195,
				TYPE_PlotGrowth = 4196,
				TYPE_SIIndex = 4197,
				TYPE_GelFrame = 4198,
				TYPE_BopPopCustom = 4199,
				TYPE_Fbi2 = 4200,
			}

			public string GetTypeName()
			{
				switch ((Type)(m_pHeader.m_nType))
				{
					case Type.TYPE_DIMENSIONS_B2:
					{
						return "TYPE_DIMENSIONS_B2";
					}

					case Type.TYPE_BLANK_B2:
					{
						return "TYPE_BLANK_B2";
					}

					case Type.TYPE_INTEGER_B2_ONLY:
					{
						return "TYPE_INTEGER_B2_ONLY";
					}

					case Type.TYPE_NUMBER_B2:
					{
						return "TYPE_NUMBER_B2";
					}

					case Type.TYPE_LABEL_B2:
					{
						return "TYPE_LABEL_B2";
					}

					case Type.TYPE_BOOLERR_B2:
					{
						return "TYPE_BOOLERR_B2";
					}

					case Type.TYPE_FORMULA:
					{
						return "TYPE_FORMULA";
					}

					case Type.TYPE_STRING_B2:
					{
						return "TYPE_STRING_B2";
					}

					case Type.TYPE_ROW_B2:
					{
						return "TYPE_ROW_B2";
					}

					case Type.TYPE_BOF_B2:
					{
						return "TYPE_BOF_B2";
					}

					case Type.TYPE_EOF:
					{
						return "TYPE_EOF";
					}

					case Type.TYPE_INDEX_B2_ONLY:
					{
						return "TYPE_INDEX_B2_ONLY";
					}

					case Type.TYPE_CALCCOUNT:
					{
						return "TYPE_CALCCOUNT";
					}

					case Type.TYPE_CALCMODE:
					{
						return "TYPE_CALCMODE";
					}

					case Type.TYPE_CALC_PRECISION:
					{
						return "TYPE_CALC_PRECISION";
					}

					case Type.TYPE_REFMODE:
					{
						return "TYPE_REFMODE";
					}

					case Type.TYPE_DELTA:
					{
						return "TYPE_DELTA";
					}

					case Type.TYPE_ITERATION:
					{
						return "TYPE_ITERATION";
					}

					case Type.TYPE_PROTECT:
					{
						return "TYPE_PROTECT";
					}

					case Type.TYPE_PASSWORD:
					{
						return "TYPE_PASSWORD";
					}

					case Type.TYPE_HEADER:
					{
						return "TYPE_HEADER";
					}

					case Type.TYPE_FOOTER:
					{
						return "TYPE_FOOTER";
					}

					case Type.TYPE_EXTERNCOUNT:
					{
						return "TYPE_EXTERNCOUNT";
					}

					case Type.TYPE_EXTERN_SHEET:
					{
						return "TYPE_EXTERN_SHEET";
					}

					case Type.TYPE_NAME_B2:
					{
						return "TYPE_NAME_B2";
					}

					case Type.TYPE_WIN_PROTECT:
					{
						return "TYPE_WIN_PROTECT";
					}

					case Type.TYPE_VERTICALPAGEBREAKS:
					{
						return "TYPE_VERTICALPAGEBREAKS";
					}

					case Type.TYPE_HORIZONTALPAGEBREAKS:
					{
						return "TYPE_HORIZONTALPAGEBREAKS";
					}

					case Type.TYPE_NOTE:
					{
						return "TYPE_NOTE";
					}

					case Type.TYPE_SELECTION:
					{
						return "TYPE_SELECTION";
					}

					case Type.TYPE_FORMAT_B2:
					{
						return "TYPE_FORMAT_B2";
					}

					case Type.TYPE_BUILTINFMTCOUNT_B2:
					{
						return "TYPE_BUILTINFMTCOUNT_B2";
					}

					case Type.TYPE_COLUMNDEFAULT_B2_ONLY:
					{
						return "TYPE_COLUMNDEFAULT_B2_ONLY";
					}

					case Type.TYPE_ARRAY_B2_ONLY:
					{
						return "TYPE_ARRAY_B2_ONLY";
					}

					case Type.TYPE_DATE_1904:
					{
						return "TYPE_DATE_1904";
					}

					case Type.TYPE_EXTERN_NAME:
					{
						return "TYPE_EXTERN_NAME";
					}

					case Type.TYPE_COLWIDTH_B2_ONLY:
					{
						return "TYPE_COLWIDTH_B2_ONLY";
					}

					case Type.TYPE_DEFAULTROWHEIGHT_B2_ONLY:
					{
						return "TYPE_DEFAULTROWHEIGHT_B2_ONLY";
					}

					case Type.TYPE_LEFT_MARGIN:
					{
						return "TYPE_LEFT_MARGIN";
					}

					case Type.TYPE_RIGHT_MARGIN:
					{
						return "TYPE_RIGHT_MARGIN";
					}

					case Type.TYPE_TOP_MARGIN:
					{
						return "TYPE_TOP_MARGIN";
					}

					case Type.TYPE_BOTTOM_MARGIN:
					{
						return "TYPE_BOTTOM_MARGIN";
					}

					case Type.TYPE_PrintRowCol:
					{
						return "TYPE_PrintRowCol";
					}

					case Type.TYPE_PrintGrid:
					{
						return "TYPE_PrintGrid";
					}

					case Type.TYPE_FILEPASS:
					{
						return "TYPE_FILEPASS";
					}

					case Type.TYPE_FONT:
					{
						return "TYPE_FONT";
					}

					case Type.TYPE_FONT2_B2_ONLY:
					{
						return "TYPE_FONT2_B2_ONLY";
					}

					case Type.TYPE_PRINT_SIZE:
					{
						return "TYPE_PRINT_SIZE";
					}

					case Type.TYPE_TABLEOP_B2:
					{
						return "TYPE_TABLEOP_B2";
					}

					case Type.TYPE_TABLEOP2_B2:
					{
						return "TYPE_TABLEOP2_B2";
					}

					case Type.TYPE_CONTINUE:
					{
						return "TYPE_CONTINUE";
					}

					case Type.TYPE_WINDOW1:
					{
						return "TYPE_WINDOW1";
					}

					case Type.TYPE_WINDOW2_B2:
					{
						return "TYPE_WINDOW2_B2";
					}

					case Type.TYPE_BACKUP:
					{
						return "TYPE_BACKUP";
					}

					case Type.TYPE_PANE:
					{
						return "TYPE_PANE";
					}

					case Type.TYPE_CODE_PAGE:
					{
						return "TYPE_CODE_PAGE";
					}

					case Type.TYPE_XF_B2:
					{
						return "TYPE_XF_B2";
					}

					case Type.TYPE_IXFE_B2_ONLY:
					{
						return "TYPE_IXFE_B2_ONLY";
					}

					case Type.TYPE_EFONT_B2_ONLY:
					{
						return "TYPE_EFONT_B2_ONLY";
					}

					case Type.TYPE_PLS:
					{
						return "TYPE_PLS";
					}

					case Type.TYPE_DCONREF:
					{
						return "TYPE_DCONREF";
					}

					case Type.TYPE_DEFCOLWIDTH:
					{
						return "TYPE_DEFCOLWIDTH";
					}

					case Type.TYPE_BUILTINFMTCOUNT_B3:
					{
						return "TYPE_BUILTINFMTCOUNT_B3";
					}

					case Type.TYPE_XCT:
					{
						return "TYPE_XCT";
					}

					case Type.TYPE_CRN:
					{
						return "TYPE_CRN";
					}

					case Type.TYPE_FILESHARING:
					{
						return "TYPE_FILESHARING";
					}

					case Type.TYPE_WRITE_ACCESS:
					{
						return "TYPE_WRITE_ACCESS";
					}

					case Type.TYPE_OBJ:
					{
						return "TYPE_OBJ";
					}

					case Type.TYPE_UNCALCED:
					{
						return "TYPE_UNCALCED";
					}

					case Type.TYPE_SAVERECALC:
					{
						return "TYPE_SAVERECALC";
					}

					case Type.TYPE_OBJECTPROTECT:
					{
						return "TYPE_OBJECTPROTECT";
					}

					case Type.TYPE_COLINFO:
					{
						return "TYPE_COLINFO";
					}

					case Type.TYPE_RK2_mythical:
					{
						return "TYPE_RK2_mythical_";
					}

					case Type.TYPE_GUTS:
					{
						return "TYPE_GUTS";
					}

					case Type.TYPE_WSBOOL:
					{
						return "TYPE_WSBOOL";
					}

					case Type.TYPE_GRIDSET:
					{
						return "TYPE_GRIDSET";
					}

					case Type.TYPE_HCENTER:
					{
						return "TYPE_HCENTER";
					}

					case Type.TYPE_VCENTER:
					{
						return "TYPE_VCENTER";
					}

					case Type.TYPE_BOUND_SHEET_8:
					{
						return "TYPE_BOUND_SHEET_8";
					}

					case Type.TYPE_WRITEPROT:
					{
						return "TYPE_WRITEPROT";
					}

					case Type.TYPE_COUNTRY:
					{
						return "TYPE_COUNTRY";
					}

					case Type.TYPE_HIDE_OBJ:
					{
						return "TYPE_HIDE_OBJ";
					}

					case Type.TYPE_SHEETSOFFSET:
					{
						return "TYPE_SHEETSOFFSET";
					}

					case Type.TYPE_SHEETHDR:
					{
						return "TYPE_SHEETHDR";
					}

					case Type.TYPE_SORT:
					{
						return "TYPE_SORT";
					}

					case Type.TYPE_PALETTE:
					{
						return "TYPE_PALETTE";
					}

					case Type.TYPE_STANDARDWIDTH:
					{
						return "TYPE_STANDARDWIDTH";
					}

					case Type.TYPE_FILTERMODE:
					{
						return "TYPE_FILTERMODE";
					}

					case Type.TYPE_BUILT_IN_FN_GROUP_COUNT:
					{
						return "TYPE_BUILT_IN_FN_GROUP_COUNT";
					}

					case Type.TYPE_AUTOFILTERINFO:
					{
						return "TYPE_AUTOFILTERINFO";
					}

					case Type.TYPE_AUTOFILTER:
					{
						return "TYPE_AUTOFILTER";
					}

					case Type.TYPE_SCL:
					{
						return "TYPE_SCL";
					}

					case Type.TYPE_SETUP:
					{
						return "TYPE_SETUP";
					}

					case Type.TYPE_GCW:
					{
						return "TYPE_GCW";
					}

					case Type.TYPE_MULRK:
					{
						return "TYPE_MULRK";
					}

					case Type.TYPE_MULBLANK:
					{
						return "TYPE_MULBLANK";
					}

					case Type.TYPE_MMS:
					{
						return "TYPE_MMS";
					}

					case Type.TYPE_RSTRING:
					{
						return "TYPE_RSTRING";
					}

					case Type.TYPE_DBCELL:
					{
						return "TYPE_DBCELL";
					}

					case Type.TYPE_BOOK_BOOL:
					{
						return "TYPE_BOOK_BOOL";
					}

					case Type.TYPE_SCENPROTECT:
					{
						return "TYPE_SCENPROTECT";
					}

					case Type.TYPE_XF:
					{
						return "TYPE_XF";
					}

					case Type.TYPE_INTERFACE_HDR:
					{
						return "TYPE_INTERFACE_HDR";
					}

					case Type.TYPE_INTERFACE_END:
					{
						return "TYPE_INTERFACE_END";
					}

					case Type.TYPE_MergeCells:
					{
						return "TYPE_MergeCells";
					}

					case Type.TYPE_BITMAP:
					{
						return "TYPE_BITMAP";
					}

					case Type.TYPE_MSO_DRAWING_GROUP:
					{
						return "TYPE_MSO_DRAWING_GROUP";
					}

					case Type.TYPE_MSO_DRAWING:
					{
						return "TYPE_MSO_DRAWING";
					}

					case Type.TYPE_MSO_DRAWING_SELECTION:
					{
						return "TYPE_MSO_DRAWING_SELECTION";
					}

					case Type.TYPE_PHONETIC:
					{
						return "TYPE_PHONETIC";
					}

					case Type.TYPE_SST:
					{
						return "TYPE_SST";
					}

					case Type.TYPE_LABELSST:
					{
						return "TYPE_LABELSST";
					}

					case Type.TYPE_EXT_SST:
					{
						return "TYPE_EXT_SST";
					}

					case Type.TYPE_RR_TAB_ID:
					{
						return "TYPE_RR_TAB_ID";
					}

					case Type.TYPE_LABELRANGES:
					{
						return "TYPE_LABELRANGES";
					}

					case Type.TYPE_USESELFS:
					{
						return "TYPE_USESELFS";
					}

					case Type.TYPE_DSF:
					{
						return "TYPE_DSF";
					}

					case Type.TYPE_SUP_BOOK:
					{
						return "TYPE_SUP_BOOK";
					}

					case Type.TYPE_PROT_4_REV:
					{
						return "TYPE_PROT_4_REV";
					}

					case Type.TYPE_CONDFMT:
					{
						return "TYPE_CONDFMT";
					}

					case Type.TYPE_CF:
					{
						return "TYPE_CF";
					}

					case Type.TYPE_DVAL:
					{
						return "TYPE_DVAL";
					}

					case Type.TYPE_TXO:
					{
						return "TYPE_TXO";
					}

					case Type.TYPE_REFRESH_ALL:
					{
						return "TYPE_REFRESH_ALL";
					}

					case Type.TYPE_HLINK:
					{
						return "TYPE_HLINK";
					}

					case Type.TYPE_PROT_4_REV_PASS:
					{
						return "TYPE_PROT_4_REV_PASS";
					}

					case Type.TYPE_DV:
					{
						return "TYPE_DV";
					}

					case Type.TYPE_EXCEL9_FILE:
					{
						return "TYPE_EXCEL9_FILE";
					}

					case Type.TYPE_RECALCID:
					{
						return "TYPE_RECALCID";
					}

					case Type.TYPE_DIMENSION:
					{
						return "TYPE_DIMENSION";
					}

					case Type.TYPE_BLANK:
					{
						return "TYPE_BLANK";
					}

					case Type.TYPE_NUMBER:
					{
						return "TYPE_NUMBER";
					}

					case Type.TYPE_LABEL:
					{
						return "TYPE_LABEL";
					}

					case Type.TYPE_BOOLERR:
					{
						return "TYPE_BOOLERR";
					}

					case Type.TYPE_FORMULA_B3:
					{
						return "TYPE_FORMULA_B3";
					}

					case Type.TYPE_STRING:
					{
						return "TYPE_STRING";
					}

					case Type.TYPE_ROW:
					{
						return "TYPE_ROW";
					}

					case Type.TYPE_INDEX_B3:
					{
						return "TYPE_INDEX_B3";
					}

					case Type.TYPE_NAME:
					{
						return "TYPE_NAME";
					}

					case Type.TYPE_ARRAY:
					{
						return "TYPE_ARRAY";
					}

					case Type.TYPE_EXTERNNAME_B3:
					{
						return "TYPE_EXTERNNAME_B3";
					}

					case Type.TYPE_DEFAULTROWHEIGHT:
					{
						return "TYPE_DEFAULTROWHEIGHT";
					}

					case Type.TYPE_FONT_B3B4:
					{
						return "TYPE_FONT_B3B4";
					}

					case Type.TYPE_TABLEOP:
					{
						return "TYPE_TABLEOP";
					}

					case Type.TYPE_WINDOW2:
					{
						return "TYPE_WINDOW2";
					}

					case Type.TYPE_XF_B3:
					{
						return "TYPE_XF_B3";
					}

					case Type.TYPE_RK:
					{
						return "TYPE_RK";
					}

					case Type.TYPE_STYLE:
					{
						return "TYPE_STYLE";
					}

					case Type.TYPE_STYLE_EXT:
					{
						return "TYPE_STYLE_EXT";
					}

					case Type.TYPE_FORMULA_B4:
					{
						return "TYPE_FORMULA_B4";
					}

					case Type.TYPE_FORMAT:
					{
						return "TYPE_FORMAT";
					}

					case Type.TYPE_XF_B4:
					{
						return "TYPE_XF_B4";
					}

					case Type.TYPE_SHRFMLA:
					{
						return "TYPE_SHRFMLA";
					}

					case Type.TYPE_QUICKTIP:
					{
						return "TYPE_QUICKTIP";
					}

					case Type.TYPE_BOF:
					{
						return "TYPE_BOF";
					}

					case Type.TYPE_SHEETLAYOUT:
					{
						return "TYPE_SHEETLAYOUT";
					}

					case Type.TYPE_SHEETPROTECTION:
					{
						return "TYPE_SHEETPROTECTION";
					}

					case Type.TYPE_RANGEPROTECTION:
					{
						return "TYPE_RANGEPROTECTION";
					}

					case Type.TYPE_XF_CRC:
					{
						return "TYPE_XF_CRC";
					}

					case Type.TYPE_XF_EXT:
					{
						return "TYPE_XF_EXT";
					}

					case Type.TYPE_BOOK_EXT:
					{
						return "TYPE_BOOK_EXT";
					}

					case Type.TYPE_THEME:
					{
						return "TYPE_THEME";
					}

					case Type.TYPE_COMPRESS_PICTURES:
					{
						return "TYPE_COMPRESS_PICTURES";
					}

					case Type.TYPE_PLV:
					{
						return "TYPE_PLV";
					}

					case Type.TYPE_COMPAT12:
					{
						return "TYPE_COMPAT12";
					}

					case Type.TYPE_MTR_SETTINGS:
					{
						return "TYPE_MTR_SETTINGS";
					}

					case Type.TYPE_FORCE_FULL_CALCULATION:
					{
						return "TYPE_FORCE_FULL_CALCULATION";
					}

					case Type.TYPE_TABLE_STYLES:
					{
						return "TYPE_TABLE_STYLES";
					}

					case Type.TYPE_ChartColors:
					{
						return "TYPE_ChartColors";
					}

					case Type.TYPE_ChartFrtInfo:
					{
						return "TYPE_ChartFrtInfo";
					}

					case Type.TYPE_StartBlock:
					{
						return "TYPE_StartBlock";
					}

					case Type.TYPE_EndBlock:
					{
						return "TYPE_EndBlock";
					}

					case Type.TYPE_StartObject:
					{
						return "TYPE_StartObject";
					}

					case Type.TYPE_EndObject:
					{
						return "TYPE_EndObject";
					}

					case Type.TYPE_CatLab:
					{
						return "TYPE_CatLab";
					}

					case Type.TYPE_YMult:
					{
						return "TYPE_YMult";
					}

					case Type.TYPE_FrtFontList:
					{
						return "TYPE_FrtFontList";
					}

					case Type.TYPE_DataLabExt:
					{
						return "TYPE_DataLabExt";
					}

					case Type.TYPE_DataLabExtContents:
					{
						return "TYPE_DataLabExtContents";
					}

					case Type.TYPE_NamePublish:
					{
						return "TYPE_NamePublish";
					}

					case Type.TYPE_HeaderFooter:
					{
						return "TYPE_HeaderFooter";
					}

					case Type.TYPE_CrtMlFrt:
					{
						return "TYPE_CrtMlFrt";
					}

					case Type.TYPE_ShapePropsStream:
					{
						return "TYPE_ShapePropsStream";
					}

					case Type.TYPE_TextPropsStream:
					{
						return "TYPE_TextPropsStream";
					}

					case Type.TYPE_CrtLayout12A:
					{
						return "TYPE_CrtLayout12A";
					}

					case Type.TYPE_Units:
					{
						return "TYPE_Units";
					}

					case Type.TYPE_Chart:
					{
						return "TYPE_Chart";
					}

					case Type.TYPE_Series:
					{
						return "TYPE_Series";
					}

					case Type.TYPE_DataFormat:
					{
						return "TYPE_DataFormat";
					}

					case Type.TYPE_LineFormat:
					{
						return "TYPE_LineFormat";
					}

					case Type.TYPE_MarkerFormat:
					{
						return "TYPE_MarkerFormat";
					}

					case Type.TYPE_AreaFormat:
					{
						return "TYPE_AreaFormat";
					}

					case Type.TYPE_PieFormat:
					{
						return "TYPE_PieFormat";
					}

					case Type.TYPE_AttachedLabel:
					{
						return "TYPE_AttachedLabel";
					}

					case Type.TYPE_SeriesText:
					{
						return "TYPE_SeriesText";
					}

					case Type.TYPE_ChartFormat:
					{
						return "TYPE_ChartFormat";
					}

					case Type.TYPE_Legend:
					{
						return "TYPE_Legend";
					}

					case Type.TYPE_SeriesList:
					{
						return "TYPE_SeriesList";
					}

					case Type.TYPE_Bar:
					{
						return "TYPE_Bar";
					}

					case Type.TYPE_Line:
					{
						return "TYPE_Line";
					}

					case Type.TYPE_Pie:
					{
						return "TYPE_Pie";
					}

					case Type.TYPE_Area:
					{
						return "TYPE_Area";
					}

					case Type.TYPE_Scatter:
					{
						return "TYPE_Scatter";
					}

					case Type.TYPE_CrtLine:
					{
						return "TYPE_CrtLine";
					}

					case Type.TYPE_Axis:
					{
						return "TYPE_Axis";
					}

					case Type.TYPE_Tick:
					{
						return "TYPE_Tick";
					}

					case Type.TYPE_ValueRange:
					{
						return "TYPE_ValueRange";
					}

					case Type.TYPE_CatSerRange:
					{
						return "TYPE_CatSerRange";
					}

					case Type.TYPE_AxisLine:
					{
						return "TYPE_AxisLine";
					}

					case Type.TYPE_CrtLink:
					{
						return "TYPE_CrtLink";
					}

					case Type.TYPE_DefaultText:
					{
						return "TYPE_DefaultText";
					}

					case Type.TYPE_Text:
					{
						return "TYPE_Text";
					}

					case Type.TYPE_FontX:
					{
						return "TYPE_FontX";
					}

					case Type.TYPE_ObjectLink:
					{
						return "TYPE_ObjectLink";
					}

					case Type.TYPE_Frame:
					{
						return "TYPE_Frame";
					}

					case Type.TYPE_Begin:
					{
						return "TYPE_Begin";
					}

					case Type.TYPE_End:
					{
						return "TYPE_End";
					}

					case Type.TYPE_PlotArea:
					{
						return "TYPE_PlotArea";
					}

					case Type.TYPE_Chart3d:
					{
						return "TYPE_Chart3d";
					}

					case Type.TYPE_PicF:
					{
						return "TYPE_PicF";
					}

					case Type.TYPE_DropBar:
					{
						return "TYPE_DropBar";
					}

					case Type.TYPE_Radar:
					{
						return "TYPE_Radar";
					}

					case Type.TYPE_Surf:
					{
						return "TYPE_Surf";
					}

					case Type.TYPE_RadarArea:
					{
						return "TYPE_RadarArea";
					}

					case Type.TYPE_AxisParent:
					{
						return "TYPE_AxisParent";
					}

					case Type.TYPE_LegendException:
					{
						return "TYPE_LegendException";
					}

					case Type.TYPE_ShtProps:
					{
						return "TYPE_ShtProps";
					}

					case Type.TYPE_SerToCrt:
					{
						return "TYPE_SerToCrt";
					}

					case Type.TYPE_AxesUsed:
					{
						return "TYPE_AxesUsed";
					}

					case Type.TYPE_SerParent:
					{
						return "TYPE_SerParent";
					}

					case Type.TYPE_SerAuxTrend:
					{
						return "TYPE_SerAuxTrend";
					}

					case Type.TYPE_IFmtRecord:
					{
						return "TYPE_IFmtRecord";
					}

					case Type.TYPE_Pos:
					{
						return "TYPE_Pos";
					}

					case Type.TYPE_AlRuns:
					{
						return "TYPE_AlRuns";
					}

					case Type.TYPE_BRAI:
					{
						return "TYPE_BRAI";
					}

					case Type.TYPE_BOFDatasheet:
					{
						return "TYPE_BOFDatasheet";
					}

					case Type.TYPE_ExcludeRows:
					{
						return "TYPE_ExcludeRows";
					}

					case Type.TYPE_ExcludeColumns:
					{
						return "TYPE_ExcludeColumns";
					}

					case Type.TYPE_Orient:
					{
						return "TYPE_Orient";
					}

					case Type.TYPE_WinDoc:
					{
						return "TYPE_WinDoc";
					}

					case Type.TYPE_MaxStatus:
					{
						return "TYPE_MaxStatus";
					}

					case Type.TYPE_MainWindow:
					{
						return "TYPE_MainWindow";
					}

					case Type.TYPE_SerAuxErrBar:
					{
						return "TYPE_SerAuxErrBar";
					}

					case Type.TYPE_ClrtClient:
					{
						return "TYPE_ClrtClient";
					}

					case Type.TYPE_SerFmt:
					{
						return "TYPE_SerFmt";
					}

					case Type.TYPE_LinkedSelection:
					{
						return "TYPE_LinkedSelection";
					}

					case Type.TYPE_Chart3DBarShape:
					{
						return "TYPE_Chart3DBarShape";
					}

					case Type.TYPE_Fbi:
					{
						return "TYPE_Fbi";
					}

					case Type.TYPE_BopPop:
					{
						return "TYPE_BopPop";
					}

					case Type.TYPE_AxcExt:
					{
						return "TYPE_AxcExt";
					}

					case Type.TYPE_Dat:
					{
						return "TYPE_Dat";
					}

					case Type.TYPE_PlotGrowth:
					{
						return "TYPE_PlotGrowth";
					}

					case Type.TYPE_SIIndex:
					{
						return "TYPE_SIIndex";
					}

					case Type.TYPE_GelFrame:
					{
						return "TYPE_GelFrame";
					}

					case Type.TYPE_BopPopCustom:
					{
						return "TYPE_BopPopCustom";
					}

					case Type.TYPE_Fbi2:
					{
						return "TYPE_Fbi2";
					}

				}
				return "???";
			}

			public const int SIZEOF_HEADER = 2 + 2;
			protected BiffHeader m_pHeader;
			protected Blob m_pContinueBlob;
			protected BlobView m_pBlobView;
			protected OwnedVector<BiffRecord_ContinueInfo> m_pContinueInfoVector;
			public BiffRecord(BiffHeader pHeader, Stream pStream)
			{
				m_pHeader = pHeader;
				m_pContinueBlob = null;
				BlobView pBlobView = pStream.GetSectorChain().GetBlobView();
				m_pBlobView = new BlobView(pBlobView.GetBlob(), pBlobView.GetStart() + pBlobView.GetOffset(), (int)(pBlobView.GetStart() + pBlobView.GetOffset() + m_pHeader.m_nSize));
				pBlobView.SetOffset((int)(pBlobView.GetOffset() + m_pHeader.m_nSize));
				m_pContinueInfoVector = new OwnedVector<BiffRecord_ContinueInfo>();
				if (m_pHeader.m_nType != (ushort)(Type.TYPE_CONTINUE))
				{
					while (pStream.GetStreamSize() - pStream.GetOffset() >= SIZEOF_HEADER)
					{
						int nOffset = pStream.GetOffset();
						BiffHeader pSubHeader = new BiffHeader();
						pSubHeader.m_nType = pBlobView.UnpackUint16();
						pSubHeader.m_nSize = pBlobView.UnpackUint16();
						if (!(m_pHeader.m_nType == (ushort)(Type.TYPE_MSO_DRAWING_GROUP) && pSubHeader.m_nType == (ushort)(Type.TYPE_MSO_DRAWING_GROUP) || pSubHeader.m_nType == (ushort)(Type.TYPE_CONTINUE)))
						{
							pStream.SetOffset(nOffset);
							{
								pSubHeader = null;
							}
							{
								break;
							}
						}
						pSubHeader.m_nType = (ushort)(Type.TYPE_CONTINUE);
						if (m_pContinueBlob == null)
						{
							m_pContinueBlob = new Blob(true);
							m_pContinueBlob.GetBlobView().Pack(m_pBlobView, m_pBlobView.GetSize());
							{
								m_pBlobView = null;
							}
							m_pBlobView = new BlobView(m_pContinueBlob, 0, m_pContinueBlob.GetSize());
						}
						ushort nTempType = pSubHeader.m_nType;
						BiffRecord pBiffRecord;
						{
							NumberDuck.Secret.BiffHeader __4098344237 = pSubHeader;
							pSubHeader = null;
							pBiffRecord = new BiffRecord(__4098344237, pStream);
						}
						Extend(pBiffRecord);
						m_pContinueInfoVector.Get(m_pContinueInfoVector.GetSize() - 1).m_nType = nTempType;
						m_pContinueBlob.GetBlobView().Pack(pBiffRecord.m_pBlobView, pBiffRecord.m_pBlobView.GetSize());
						{
							m_pBlobView = null;
						}
						m_pBlobView = new BlobView(m_pContinueBlob, 0, m_pContinueBlob.GetSize());
						{
							pBiffRecord = null;
						}
					}
				}
			}

			public BiffRecord(Type nType, uint nSize)
			{
				m_pHeader = new BiffHeader();
				m_pHeader.m_nType = (ushort)(nType);
				m_pHeader.m_nSize = nSize;
				m_pContinueBlob = null;
				m_pBlobView = null;
				m_pContinueInfoVector = new OwnedVector<BiffRecord_ContinueInfo>();
			}

			public new BiffRecord.Type GetType()
			{
				return (Type)(m_pHeader.m_nType);
			}

			public uint GetSize()
			{
				return (uint)(m_pHeader.m_nSize + SIZEOF_HEADER + m_pContinueInfoVector.GetSize() * SIZEOF_HEADER);
			}

			public static BiffRecord CreateBiffRecord(Stream pStream)
			{
				BlobView pBlobView = pStream.GetSectorChain().GetBlobView();
				BiffHeader pHeader = new BiffHeader();
				pHeader.m_nType = pBlobView.UnpackUint16();
				pHeader.m_nSize = pBlobView.UnpackUint16();
				switch ((Type)(pHeader.m_nType))
				{
					case Type.TYPE_BOF:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BofRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_INTERFACE_HDR:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new InterfaceHdr(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MMS:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Mms(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_INTERFACE_END:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new InterfaceEnd(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_WRITE_ACCESS:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new WriteAccess(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CODE_PAGE:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CodePage(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DSF:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DSF(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_EXCEL9_FILE:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Excel9File(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_RR_TAB_ID:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new RRTabId(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BUILT_IN_FN_GROUP_COUNT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BuiltInFnGroupCount(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_WIN_PROTECT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new WinProtect(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PROTECT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ProtectRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PASSWORD:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PasswordRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PROT_4_REV:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Prot4RevRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PROT_4_REV_PASS:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Prot4RevPassRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_WINDOW1:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Window1(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BACKUP:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Backup(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_HIDE_OBJ:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new HideObj(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DATE_1904:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Date1904(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CALC_PRECISION:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CalcPrecision(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_REFRESH_ALL:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new RefreshAllRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BOOK_BOOL:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BookBool(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_FONT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new FontRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_FORMAT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Format(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_XF:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new XF(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_STYLE:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new StyleRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MergeCells:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MergeCellsRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SST:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SstRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BOUND_SHEET_8:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BoundSheet8Record(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PALETTE:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PaletteRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PrintRowCol:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PrintRowColRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PrintGrid:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PrintGridRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CALCCOUNT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CalcCountRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_WSBOOL:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new WsBoolRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DIMENSION:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DimensionRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DEFCOLWIDTH:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DefColWidthRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DEFAULTROWHEIGHT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DefaultRowHeight(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_COLINFO:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ColInfoRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ROW:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new RowRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_LABELSST:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new LabelSstRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_RK:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new RkRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MULRK:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MulRkRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BLANK:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Blank(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MULBLANK:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MulBlank(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_NUMBER:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new NumberRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BOOLERR:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BoolErrRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_FORMULA:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new FormulaRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_WINDOW2:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Window2Record(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SELECTION:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SelectionRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_XF_CRC:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new XFCRC(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_XF_EXT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new XFExt(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BOOK_EXT:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BookExtRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_THEME:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Theme(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SUP_BOOK:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SupBookRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_EXTERN_SHEET:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ExternSheetRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CrtLink:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CrtLinkRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Units:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new UnitsRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Chart:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ChartRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Begin:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BeginRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Frame:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new FrameRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_LineFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new LineFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_AreaFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AreaFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Series:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SeriesRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BRAI:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BraiRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DataFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DataFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MarkerFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MarkerFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SerToCrt:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SerToCrtRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ShtProps:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ShtPropsRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_AxesUsed:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AxesUsedRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_AxisParent:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AxisParentRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Pos:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PosRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Axis:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AxisRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CatSerRange:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CatSerRangeRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_AxcExt:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AxcExtRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Tick:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new TickRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_FontX:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new FontXRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_AxisLine:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AxisLineRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ValueRange:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ValueRangeRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PlotArea:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PlotAreaRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ChartFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ChartFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Line:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new LineRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Bar:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BarRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Area:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new AreaRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Scatter:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ScatterRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Legend:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new LegendRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Text:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new TextRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SeriesText:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SeriesTextRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ObjectLink:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ObjectLinkRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_HCENTER:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new HCenterRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_VCENTER:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new VCenterRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_LEFT_MARGIN:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new LeftMarginRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_RIGHT_MARGIN:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new RightMarginRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_TOP_MARGIN:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new TopMarginRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_BOTTOM_MARGIN:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new BottomMarginRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SETUP:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SetupRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PRINT_SIZE:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PrintSizeRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SCL:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SclRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PlotGrowth:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PlotGrowthRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_SIIndex:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new SIIndexRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ChartFrtInfo:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ChartFrtInfoRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_StartBlock:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new StartBlockRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_EndBlock:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new EndBlockRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_ShapePropsStream:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ShapePropsStreamRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CrtLayout12A:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CrtLayout12ARecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CrtMlFrt:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CrtMlFrtRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_HeaderFooter:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new HeaderFooterRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_Chart3DBarShape:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new Chart3DBarShapeRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_PieFormat:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new PieFormatRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_CatLab:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new CatLabRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_DefaultText:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new DefaultTextRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_GelFrame:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new GelFrameRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_OBJ:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new ObjRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MSO_DRAWING_GROUP:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MsoDrawingGroupRecord(__1199093386, pStream);
							}
						}
					}

					case Type.TYPE_MSO_DRAWING:
					{
						{
							NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
							pHeader = null;
							{
								return new MsoDrawingRecord(__1199093386, pStream);
							}
						}
					}

				}
				{
					NumberDuck.Secret.BiffHeader __1199093386 = pHeader;
					pHeader = null;
					{
						return new BiffRecord(__1199093386, pStream);
					}
				}
			}

			public virtual void Write(Stream pStream, BlobView pTempBlobView)
			{
				BlobWrite(pTempBlobView);
				pTempBlobView.SetOffset(0);
				nbAssert.Assert(pTempBlobView.GetSize() == (int)(m_pHeader.m_nSize));
				m_pHeader.m_nSize = (uint)(pTempBlobView.GetSize());
				if (m_pContinueInfoVector.GetSize() == 0)
				{
					nbAssert.Assert(pTempBlobView.GetSize() <= MAX_DATA_SIZE);
					pStream.SizeToFit(SIZEOF_HEADER + pTempBlobView.GetSize());
					BlobView pStreamBlobView = pStream.GetSectorChain().GetBlobView();
					pStreamBlobView.PackUint16(m_pHeader.m_nType);
					pStreamBlobView.PackUint16((ushort)(m_pHeader.m_nSize));
					pStreamBlobView.Pack(pTempBlobView, pTempBlobView.GetSize());
				}
				else
				{
					ushort nIndex = 0;
					uint nOffset = 0;
					uint nSize = m_pHeader.m_nSize;
					ushort nRecordSize;
					{
						BiffRecord_ContinueInfo pContinueInfo = m_pContinueInfoVector.Get(nIndex);
						nbAssert.Assert(pContinueInfo.m_nOffset <= MAX_DATA_SIZE);
						nRecordSize = (ushort)(pContinueInfo.m_nOffset);
					}
					{
						pStream.SizeToFit(SIZEOF_HEADER + nRecordSize);
						BlobView pStreamBlobView = pStream.GetSectorChain().GetBlobView();
						pStreamBlobView.PackUint16(m_pHeader.m_nType);
						pStreamBlobView.PackUint16(nRecordSize);
						pStreamBlobView.Pack(pTempBlobView, nRecordSize);
					}
					nOffset += nRecordSize;
					nIndex++;
					while (nOffset != nSize)
					{
						Type eType = Type.TYPE_CONTINUE;
						if (m_pHeader.m_nType == (ushort)(Type.TYPE_MSO_DRAWING_GROUP) && nIndex == 1)
							eType = Type.TYPE_MSO_DRAWING_GROUP;
						if (nIndex < m_pContinueInfoVector.GetSize())
						{
							BiffRecord_ContinueInfo pContinueInfo = m_pContinueInfoVector.Get(nIndex);
							BiffRecord_ContinueInfo pPreviousContinueInfo = m_pContinueInfoVector.Get(nIndex - 1);
							nbAssert.Assert(pContinueInfo.m_nOffset - pPreviousContinueInfo.m_nOffset <= MAX_DATA_SIZE);
							nRecordSize = (ushort)(pContinueInfo.m_nOffset - pPreviousContinueInfo.m_nOffset);
						}
						else
						{
							nbAssert.Assert(nSize - nOffset <= MAX_DATA_SIZE);
							nRecordSize = (ushort)(nSize - nOffset);
						}
						{
							pStream.SizeToFit(SIZEOF_HEADER + nRecordSize);
							BlobView pStreamBlobView = pStream.GetSectorChain().GetBlobView();
							pStreamBlobView.PackUint16((ushort)(eType));
							pStreamBlobView.PackUint16(nRecordSize);
							pStreamBlobView.Pack(pTempBlobView, nRecordSize);
						}
						nOffset += nRecordSize;
						nIndex++;
					}
				}
			}

			public virtual void BlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public virtual void BlobWrite(BlobView pBlobView)
			{
			}

			protected void Extend(BiffRecord pBiffRecord)
			{
				nbAssert.Assert(pBiffRecord.GetType() == Type.TYPE_CONTINUE || pBiffRecord.GetType() == Type.TYPE_MSO_DRAWING_GROUP);
				m_pContinueInfoVector.PushBack(new BiffRecord_ContinueInfo((int)(m_pHeader.m_nSize), (int)(pBiffRecord.GetType())));
				uint nNewSize = m_pHeader.m_nSize + pBiffRecord.m_pHeader.m_nSize;
				m_pHeader.m_nSize = nNewSize;
			}

			~BiffRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffWorksheetStreamSize
		{
			public uint m_nSize;
		}
		class WorkbookGlobals
		{
			public SharedStringContainer m_pSharedStringContainer;
			public Vector<Picture> m_pSharedPictureVector;
			public OwnedVector<WorksheetRange> m_pWorksheetRangeVector;
			public OwnedVector<BiffWorksheetStreamSize> m_pBiffWorksheetStreamSizeVector;
			public OwnedVector<Style> m_pStyleVector;
			public Font m_pHeaderFont;
			public ushort m_Window1_nTabRatio;
			public WorkbookGlobals()
			{
				m_pSharedStringContainer = new SharedStringContainer();
				m_pSharedPictureVector = new Vector<Picture>();
				m_pWorksheetRangeVector = new OwnedVector<WorksheetRange>();
				m_pBiffWorksheetStreamSizeVector = new OwnedVector<BiffWorksheetStreamSize>();
				m_pStyleVector = new OwnedVector<Style>();
				m_pHeaderFont = new Font();
				m_pHeaderFont.SetName("Sans");
				m_pHeaderFont.SetSize(13);
				m_pHeaderFont.SetBold(false);
				m_pHeaderFont.SetItalic(false);
				m_pHeaderFont.SetUnderline(Font.Underline.UNDERLINE_NONE);
				CreateStyle();
				m_Window1_nTabRatio = 600;
			}

			~WorkbookGlobals()
			{
				Clear();
				{
					m_pSharedStringContainer = null;
				}
				{
					m_pSharedPictureVector = null;
				}
				{
					m_pWorksheetRangeVector = null;
				}
				{
					m_pBiffWorksheetStreamSizeVector = null;
				}
				{
					m_pStyleVector = null;
				}
				{
					m_pHeaderFont = null;
				}
			}

			public void Clear()
			{
				m_pSharedStringContainer.Clear();
				m_pSharedPictureVector.Clear();
				while (m_pWorksheetRangeVector.GetSize() > 0)
				{
					WorksheetRange pWorksheetRange = m_pWorksheetRangeVector.PopBack();
					{
						pWorksheetRange = null;
					}
				}
				while (m_pBiffWorksheetStreamSizeVector.GetSize() > 0)
				{
					BiffWorksheetStreamSize pBiffWorksheetStreamSize = m_pBiffWorksheetStreamSizeVector.PopBack();
					{
						pBiffWorksheetStreamSize = null;
					}
				}
			}

			public void PushBiffWorksheetStreamSize(uint nStreamSize)
			{
				BiffWorksheetStreamSize pStreamSize = new BiffWorksheetStreamSize();
				pStreamSize.m_nSize = nStreamSize;
				{
					NumberDuck.Secret.BiffWorksheetStreamSize __521035195 = pStreamSize;
					pStreamSize = null;
					m_pBiffWorksheetStreamSizeVector.PushBack(__521035195);
				}
			}

			public string GetSharedStringByIndex(uint nIndex)
			{
				return m_pSharedStringContainer.Get((int)(nIndex));
			}

			public uint GetSharedStringIndex(string szString)
			{
				return (uint)(m_pSharedStringContainer.GetIndex(szString));
			}

			public uint PushSharedString(string szString)
			{
				return (uint)(m_pSharedStringContainer.Push(szString));
			}

			public uint PushPicture(Picture pPicture)
			{
				m_pSharedPictureVector.PushBack(pPicture);
				return (uint)(m_pSharedPictureVector.GetSize());
			}

			public WorksheetRange GetWorksheetRangeByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex <= (ushort)(m_pWorksheetRangeVector.GetSize()));
				return m_pWorksheetRangeVector.Get((int)(nIndex));
			}

			public ushort GetWorksheetRangeIndex(ushort nFirst, ushort nLast)
			{
				for (int i = 0; i < m_pWorksheetRangeVector.GetSize(); i++)
				{
					WorksheetRange pWorksheetRange = m_pWorksheetRangeVector.Get(i);
					if (nFirst == pWorksheetRange.m_nFirst)
						if (nLast == pWorksheetRange.m_nLast)
							return (ushort)(i);
				}
				{
					WorksheetRange pWorksheetRange = new WorksheetRange(nFirst, nLast);
					{
						NumberDuck.Secret.WorksheetRange __4286285562 = pWorksheetRange;
						pWorksheetRange = null;
						m_pWorksheetRangeVector.PushBack(__4286285562);
					}
				}
				return (ushort)(m_pWorksheetRangeVector.GetSize() - 1);
			}

			public Style GetStyleByIndex(ushort nIndex)
			{
				if (nIndex >= GetNumStyle())
					return null;
				return m_pStyleVector.Get(nIndex);
			}

			public ushort GetStyleIndex(Style pStyle)
			{
				for (int i = 0; i < m_pStyleVector.GetSize(); i++)
					if (m_pStyleVector.Get(i) == pStyle)
						return (ushort)(i + 15);
				nbAssert.Assert(false);
				return 0;
			}

			public ushort GetNumStyle()
			{
				return (ushort)(m_pStyleVector.GetSize());
			}

			public Style CreateStyle()
			{
				Style pStyle = new Style();
				Style pTempStyle = pStyle;
				{
					NumberDuck.Style __2188486757 = pStyle;
					pStyle = null;
					m_pStyleVector.PushBack(__2188486757);
				}
				{
					return pTempStyle;
				}
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffWorkbookGlobals : WorkbookGlobals
		{
			public const ushort NUM_DEFAULT_PALETTE_ENTRY = 8;
			public const ushort NUM_CUSTOM_PALETTE_ENTRY = 56;
			public const ushort PALETTE_INDEX_ERROR = 0xFF;
			public const ushort PALETTE_INDEX_DEFAULT = 0xFE;
			public const byte PALETTE_INDEX_DEFAULT_FOREGROUND = 0x0040;
			public const byte PALETTE_INDEX_DEFAULT_BACKGROUND = 0x0041;
			public const ushort PALETTE_INDEX_DEFAULT_CHART_FOREGROUND = 0x004D;
			public const ushort PALETTE_INDEX_DEFAULT_CHART_BACKGROUND = 0x004E;
			public const ushort PALETTE_INDEX_DEFAULT_NEUTRAL = 0x004F;
			public const ushort PALETTE_INDEX_DEFAULT_TOOL_TIP_TEXT = 0x0051;
			public const ushort PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC = 0x7FFF;
			public static readonly uint[] DEFAULT_COLOR = {0x000000, 0xFFFFFF, 0x0000FF, 0x00FF00, 0xFF0000, 0x00FFFF, 0xFF00FF, 0xFFFF00};
			public static readonly uint[] DEFAULT_CUSTOM_COLOR = {((uint)(0)) << 0 | ((uint)(0)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(255)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(0)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(255)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(0)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(255)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(0)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(255)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(0)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(128)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(0)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(128)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(0)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(128)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(192)) << 0 | ((uint)(192)) << 8 | ((uint)(192)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(128)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(153)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(51)) << 8 | ((uint)(102)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(255)) << 8 | ((uint)(204)) << 16 | ((uint)(255)) << 24, ((uint)(204)) << 0 | ((uint)(255)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(102)) << 0 | ((uint)(0)) << 8 | ((uint)(102)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(128)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(102)) << 8 | ((uint)(204)) << 16 | ((uint)(255)) << 24, ((uint)(204)) << 0 | ((uint)(204)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(0)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(0)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(255)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(255)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(0)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(128)) << 0 | ((uint)(0)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(128)) << 8 | ((uint)(128)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(0)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(204)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(204)) << 0 | ((uint)(255)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(204)) << 0 | ((uint)(255)) << 8 | ((uint)(204)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(255)) << 8 | ((uint)(153)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(204)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(153)) << 8 | ((uint)(204)) << 16 | ((uint)(255)) << 24, ((uint)(204)) << 0 | ((uint)(153)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(204)) << 8 | ((uint)(153)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(102)) << 8 | ((uint)(255)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(204)) << 8 | ((uint)(204)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(204)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(204)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(153)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(255)) << 0 | ((uint)(102)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(102)) << 0 | ((uint)(102)) << 8 | ((uint)(153)) << 16 | ((uint)(255)) << 24, ((uint)(150)) << 0 | ((uint)(150)) << 8 | ((uint)(150)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(51)) << 8 | ((uint)(102)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(153)) << 8 | ((uint)(102)) << 16 | ((uint)(255)) << 24, ((uint)(0)) << 0 | ((uint)(51)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(51)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(51)) << 8 | ((uint)(0)) << 16 | ((uint)(255)) << 24, ((uint)(153)) << 0 | ((uint)(51)) << 8 | ((uint)(102)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(51)) << 8 | ((uint)(153)) << 16 | ((uint)(255)) << 24, ((uint)(51)) << 0 | ((uint)(51)) << 8 | ((uint)(51)) << 16 | ((uint)(255)) << 24};
			protected BiffRecordContainer m_pBiffRecordContainer;
			protected OwnedVector<InternalString> m_sWorkheetNameVector;
			protected InternalString m_sTempWorksheetName;
			protected MsoDrawingGroupRecord m_MsoDrawingGroupRecord;
			protected PaletteRecord m_pPaletteRecord;
			public BiffWorkbookGlobals(BiffRecord pInitialBiffRecord, Stream pStream)
			{
				m_pBiffRecordContainer = new BiffRecordContainer(pInitialBiffRecord, pStream);
				m_sWorkheetNameVector = new OwnedVector<InternalString>();
				m_sTempWorksheetName = null;
				m_MsoDrawingGroupRecord = null;
				m_pPaletteRecord = null;
				Vector<FontRecord> pFontRecordVector = new Vector<FontRecord>();
				Vector<Format> pFormatRecordVector = new Vector<Format>();
				Vector<XF> pXFVector = new Vector<XF>();
				XFCRC pXFCRC = null;
				Vector<XFExt> pXFExtVector = new Vector<XFExt>();
				Theme pTheme = null;
				for (int i = 0; i < m_pBiffRecordContainer.m_pBiffRecordVector.GetSize(); i++)
				{
					BiffRecord pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_SST)
					{
						SstRecord pSstRecord = (SstRecord)(pBiffRecord);
						for (int j = 0; j < pSstRecord.GetNumString(); j++)
							m_pSharedStringContainer.Push(pSstRecord.GetString(j));
					}
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_BOUND_SHEET_8)
					{
						BoundSheet8Record pBoundSheet8Record = (BoundSheet8Record)(pBiffRecord);
						m_sWorkheetNameVector.PushBack(new InternalString(pBoundSheet8Record.GetName()));
					}
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_FONT)
						pFontRecordVector.PushBack((FontRecord)(pBiffRecord));
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_FORMAT)
						pFormatRecordVector.PushBack((Format)(pBiffRecord));
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_XF)
						pXFVector.PushBack((XF)(pBiffRecord));
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_XF_CRC)
						pXFCRC = (XFCRC)(pBiffRecord);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_XF_EXT)
						pXFExtVector.PushBack((XFExt)(pBiffRecord));
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_PALETTE)
						m_pPaletteRecord = (PaletteRecord)(pBiffRecord);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_MSO_DRAWING_GROUP)
						m_MsoDrawingGroupRecord = (MsoDrawingGroupRecord)(pBiffRecord);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_EXTERN_SHEET)
					{
						while (m_pWorksheetRangeVector.GetSize() > 0)
						{
							WorksheetRange pWorksheetRange = m_pWorksheetRangeVector.PopBack();
							{
								pWorksheetRange = null;
							}
						}
						ExternSheetRecord pExternSheetRecord = (ExternSheetRecord)(pBiffRecord);
						for (ushort j = 0; j < pExternSheetRecord.GetNumXTI(); j++)
						{
							XTIStruct pXTI = pExternSheetRecord.GetXTIByIndex(j);
							WorksheetRange pWorksheetRange = new WorksheetRange((ushort)(pXTI.m_itabFirst), (ushort)(pXTI.m_itabLast));
							{
								NumberDuck.Secret.WorksheetRange __4286285562 = pWorksheetRange;
								pWorksheetRange = null;
								m_pWorksheetRangeVector.PushBack(__4286285562);
							}
						}
					}
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_THEME)
						pTheme = (Theme)(pBiffRecord);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_WINDOW1)
					{
						m_Window1_nTabRatio = ((Window1)(pBiffRecord)).GetTabRatio();
					}
				}
				bool bXFExtValid = false;
				if (pXFCRC != null)
					bXFExtValid = pXFCRC.ValidateCrc(pXFVector);
				for (int i = 15; i < pXFVector.GetSize(); i++)
				{
					XF pXF = pXFVector.Get(i);
					{
						Style pStyle = null;
						if (i == 15)
							pStyle = this.GetStyleByIndex(0);
						else
							pStyle = this.CreateStyle();
						ushort nFontIndex = pXF.GetFontIndex();
						if (nFontIndex > 4)
							nFontIndex -= 1;
						FontRecord pFontRecord = pFontRecordVector.Get(nFontIndex);
						pStyle.GetFont().SetName(pFontRecord.GetName());
						pStyle.GetFont().m_pImpl.m_nSizeTwips = pFontRecord.GetSizeTwips();
						pStyle.GetFont().SetBold(pFontRecord.GetBold());
						pStyle.GetFont().SetItalic(pFontRecord.GetItalic());
						pStyle.GetFont().SetUnderline(pFontRecord.GetUnderline());
						ushort nFormatIndex = pXF.GetFormatIndex();
						switch (nFormatIndex)
						{
							case 0:
							{
								pStyle.SetFormat("General");
								break;
							}

							case 1:
							{
								pStyle.SetFormat("0");
								break;
							}

							case 2:
							{
								pStyle.SetFormat("0.00");
								break;
							}

							case 3:
							{
								pStyle.SetFormat("#,##0");
								break;
							}

							case 4:
							{
								pStyle.SetFormat("#,##0.00");
								break;
							}

							case 9:
							{
								pStyle.SetFormat("0%");
								break;
							}

							case 10:
							{
								pStyle.SetFormat("0.00%");
								break;
							}

							case 11:
							{
								pStyle.SetFormat("0.00E+00");
								break;
							}

							case 12:
							{
								pStyle.SetFormat("# \\?/\\?");
								break;
							}

							case 13:
							{
								pStyle.SetFormat("# \\?\\?/\\?\\?");
								break;
							}

							case 14:
							{
								pStyle.SetFormat("mm-dd-yy");
								break;
							}

							case 15:
							{
								pStyle.SetFormat("d-mmm-yy");
								break;
							}

							case 16:
							{
								pStyle.SetFormat("d-mmm");
								break;
							}

							case 17:
							{
								pStyle.SetFormat("mmm-yy");
								break;
							}

							case 18:
							{
								pStyle.SetFormat("h:mm AM/PM");
								break;
							}

							case 19:
							{
								pStyle.SetFormat("h:mm:ss AM/PM");
								break;
							}

							case 20:
							{
								pStyle.SetFormat("h:mm");
								break;
							}

							case 21:
							{
								pStyle.SetFormat("h:mm:ss");
								break;
							}

							case 22:
							{
								pStyle.SetFormat("m/d/yy h:mm");
								break;
							}

							case 37:
							{
								pStyle.SetFormat("#,##0 ;(#,##0)");
								break;
							}

							case 38:
							{
								pStyle.SetFormat("#,##0 ;[Red](#,##0)");
								break;
							}

							case 39:
							{
								pStyle.SetFormat("#,##0.00;-#,##0.00");
								break;
							}

							case 40:
							{
								pStyle.SetFormat("#,##0.00;[Red](#,##0.00)");
								break;
							}

							case 45:
							{
								pStyle.SetFormat("mm:ss");
								break;
							}

							case 46:
							{
								pStyle.SetFormat("[h]:mm:ss");
								break;
							}

							case 47:
							{
								pStyle.SetFormat("mmss.0");
								break;
							}

							case 48:
							{
								pStyle.SetFormat("##0.0E+0");
								break;
							}

							case 49:
							{
								pStyle.SetFormat("@");
								break;
							}

						}
						for (int j = 0; j < pFormatRecordVector.GetSize(); j++)
						{
							Format pFormatRecord = pFormatRecordVector.Get(j);
							if (pFormatRecord.GetFormatIndex() == nFormatIndex)
							{
								pStyle.SetFormat(pFormatRecord.GetFormat());
							}
						}
						ushort nBackgroundPaletteIndex = pXF.GetBackgroundPaletteIndex();
						if (nBackgroundPaletteIndex != PALETTE_INDEX_DEFAULT && nBackgroundPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
							pStyle.GetBackgroundColor(true).SetFromRgba(GetPaletteColorByIndex(nBackgroundPaletteIndex));
						pStyle.SetFillPattern(pXF.GetFillPattern());
						ushort nForegroundPaletteIndex = pXF.GetForegroundPaletteIndex();
						if (nForegroundPaletteIndex != PALETTE_INDEX_DEFAULT && nForegroundPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
							pStyle.GetFillPatternColor(true).SetFromRgba(GetPaletteColorByIndex(nForegroundPaletteIndex));
						ushort nPaletteIndex = pFontRecord.GetColourIndex();
						if (nPaletteIndex != PALETTE_INDEX_DEFAULT && nPaletteIndex != PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC)
							pStyle.GetFont().GetColor(true).SetFromRgba(GetPaletteColorByIndex(nPaletteIndex));
						pStyle.SetHorizontalAlign((Style.HorizontalAlign)(pXF.GetHorizontalAlign()));
						pStyle.SetVerticalAlign((Style.VerticalAlign)(pXF.GetVerticalAlign()));
						pStyle.GetTopBorderLine().SetType(pXF.GetTopBorderType());
						nPaletteIndex = pXF.GetTopBorderPaletteIndex();
						if (nPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
						{
							Line pLine = pStyle.GetTopBorderLine();
							pLine.GetColor().SetFromRgba(GetPaletteColorByIndex(nPaletteIndex));
						}
						if (nPaletteIndex == PALETTE_INDEX_DEFAULT_FOREGROUND)
						{
							Line pLine = pStyle.GetTopBorderLine();
							pLine.GetColor().Set(0x00, 0x00, 0x00);
						}
						pStyle.GetRightBorderLine().SetType(pXF.GetRightBorderType());
						nPaletteIndex = pXF.GetRightBorderPaletteIndex();
						if (nPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
						{
							Line pLine = pStyle.GetRightBorderLine();
							pLine.GetColor().SetFromRgba(GetPaletteColorByIndex(nPaletteIndex));
						}
						if (nPaletteIndex == PALETTE_INDEX_DEFAULT_FOREGROUND)
						{
							Line pLine = pStyle.GetRightBorderLine();
							pLine.GetColor().Set(0x00, 0x00, 0x00);
						}
						pStyle.GetBottomBorderLine().SetType(pXF.GetBottomBorderType());
						nPaletteIndex = pXF.GetBottomBorderPaletteIndex();
						if (nPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
						{
							Line pLine = pStyle.GetBottomBorderLine();
							pLine.GetColor().SetFromRgba(GetPaletteColorByIndex(nPaletteIndex));
						}
						if (nPaletteIndex == PALETTE_INDEX_DEFAULT_FOREGROUND)
						{
							Line pLine = pStyle.GetBottomBorderLine();
							pLine.GetColor().Set(0x00, 0x00, 0x00);
						}
						pStyle.GetLeftBorderLine().SetType(pXF.GetLeftBorderType());
						nPaletteIndex = pXF.GetLeftBorderPaletteIndex();
						if (nPaletteIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY)
						{
							Line pLine = pStyle.GetLeftBorderLine();
							pLine.GetColor().SetFromRgba(GetPaletteColorByIndex(nPaletteIndex));
						}
						if (nPaletteIndex == PALETTE_INDEX_DEFAULT_FOREGROUND)
						{
							Line pLine = pStyle.GetLeftBorderLine();
							pLine.GetColor().Set(0x00, 0x00, 0x00);
						}
						if (bXFExtValid && pXF.GetHasXFExt())
						{
							for (int j = 0; j < pXFExtVector.GetSize(); j++)
							{
								XFExt pXFExt = pXFExtVector.Get(j);
								if (pXFExt.GetXFIndex() == i)
								{
									{
										Color pColor = pXFExt.GetTextColor(pTheme);
										if (pColor != null)
											pStyle.GetFont().GetColor(true).SetFromColor(pColor);
									}
									{
										Color pColor = pXFExt.GetForegroundColor(pTheme);
										if (pColor != null)
											pStyle.GetBackgroundColor(true).SetFromColor(pColor);
									}
									break;
								}
							}
						}
					}
				}
				{
					FontRecord pFontRecord = pFontRecordVector.Get(0);
					m_pHeaderFont.SetName(pFontRecord.GetName());
					m_pHeaderFont.m_pImpl.m_nSizeTwips = pFontRecord.GetSizeTwips();
					m_pHeaderFont.SetBold(pFontRecord.GetBold());
					m_pHeaderFont.SetItalic(pFontRecord.GetItalic());
					m_pHeaderFont.SetUnderline(pFontRecord.GetUnderline());
				}
				{
					pFontRecordVector = null;
				}
				{
					pFormatRecordVector = null;
				}
				{
					pXFVector = null;
				}
				{
					pXFExtVector = null;
				}
			}

			public static void Write(WorkbookGlobals pWorkbookGlobals, OwnedVector<Worksheet> pWorksheetVector, Stream pStream)
			{
				Vector<BoundSheet8Record> pBoundSheet8RecordVector = new Vector<BoundSheet8Record>();
				BiffRecordContainer pBiffRecordContainer = new BiffRecordContainer();
				OwnedVector<XF> pXFVector = new OwnedVector<XF>();
				pBiffRecordContainer.AddBiffRecord(new BofRecord(BofRecord.BofType.BOF_TYPE_WORKBOOK_GLOBALS));
				pBiffRecordContainer.AddBiffRecord(new InterfaceHdr());
				pBiffRecordContainer.AddBiffRecord(new Mms());
				pBiffRecordContainer.AddBiffRecord(new InterfaceEnd());
				pBiffRecordContainer.AddBiffRecord(new WriteAccess());
				pBiffRecordContainer.AddBiffRecord(new CodePage());
				pBiffRecordContainer.AddBiffRecord(new DSF());
				pBiffRecordContainer.AddBiffRecord(new Excel9File());
				pBiffRecordContainer.AddBiffRecord(new RRTabId((ushort)(pWorksheetVector.GetSize())));
				pBiffRecordContainer.AddBiffRecord(new BuiltInFnGroupCount());
				pBiffRecordContainer.AddBiffRecord(new WinProtect());
				pBiffRecordContainer.AddBiffRecord(new ProtectRecord());
				pBiffRecordContainer.AddBiffRecord(new PasswordRecord());
				pBiffRecordContainer.AddBiffRecord(new Prot4RevRecord());
				pBiffRecordContainer.AddBiffRecord(new Prot4RevPassRecord());
				pBiffRecordContainer.AddBiffRecord(new Window1(pWorkbookGlobals.m_Window1_nTabRatio));
				pBiffRecordContainer.AddBiffRecord(new Backup());
				pBiffRecordContainer.AddBiffRecord(new HideObj());
				pBiffRecordContainer.AddBiffRecord(new Date1904());
				pBiffRecordContainer.AddBiffRecord(new CalcPrecision());
				pBiffRecordContainer.AddBiffRecord(new RefreshAllRecord());
				pBiffRecordContainer.AddBiffRecord(new BookBool());
				{
					Font pHeaderFont = pWorkbookGlobals.m_pHeaderFont;
					pBiffRecordContainer.AddBiffRecord(new FontRecord(pHeaderFont.GetName(), (ushort)(pHeaderFont.m_pImpl.m_nSizeTwips), PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC, pHeaderFont.GetBold(), pHeaderFont.GetItalic(), pHeaderFont.GetUnderline()));
				}
				for (int i = 0; i < pWorkbookGlobals.m_pStyleVector.GetSize(); i++)
				{
					Style pStyle = pWorkbookGlobals.m_pStyleVector.Get(i);
					ushort nPaletteIndex = SnapToPalette(pStyle.GetFont().GetColor(false));
					if (nPaletteIndex == PALETTE_INDEX_DEFAULT)
						nPaletteIndex = PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC;
					pBiffRecordContainer.AddBiffRecord(new FontRecord(pStyle.GetFont().GetName(), (ushort)(pStyle.GetFont().m_pImpl.m_nSizeTwips), nPaletteIndex, pStyle.GetFont().GetBold(), pStyle.GetFont().GetItalic(), pStyle.GetFont().GetUnderline()));
				}
				ushort nFormatIndex = 164;
				for (int i = 0; i < pWorkbookGlobals.m_pStyleVector.GetSize(); i++)
				{
					Style pStyle = pWorkbookGlobals.m_pStyleVector.Get(i);
					pStyle.m_pImplementation.m_nFormatIndex = 0;
					for (int j = 0; j < i; j++)
					{
						Style pTestStyle = pWorkbookGlobals.m_pStyleVector.Get(j);
						if (ExternalString.Equal(pTestStyle.GetFormat(), pStyle.GetFormat()))
						{
							pStyle.m_pImplementation.m_nFormatIndex = pTestStyle.m_pImplementation.m_nFormatIndex;
							break;
						}
					}
					if (pStyle.m_pImplementation.m_nFormatIndex == 0 && nFormatIndex < 382)
					{
						pStyle.m_pImplementation.m_nFormatIndex = nFormatIndex;
						pBiffRecordContainer.AddBiffRecord(new Format(nFormatIndex, pStyle.GetFormat()));
						nFormatIndex++;
					}
				}
				for (int i = 0; i < 15; i++)
				{
					pXFVector.PushBack(new XF(0, 43, 1, 0xFFF, 1, 0, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, PALETTE_INDEX_DEFAULT_FOREGROUND, PALETTE_INDEX_DEFAULT_BACKGROUND));
				}
				for (ushort i = 0; i < (ushort)(pWorkbookGlobals.m_pStyleVector.GetSize()); i++)
				{
					Style pStyle = pWorkbookGlobals.m_pStyleVector.Get(i);
					pXFVector.PushBack(new XF((ushort)(i + 1), (ushort)(pStyle.m_pImplementation.m_nFormatIndex), pStyle, true));
				}
				XFCRC pXFCRC = new XFCRC(pXFVector);
				while (pXFVector.GetSize() > 0)
				{
					pBiffRecordContainer.AddBiffRecord(pXFVector.PopFront());
				}
				{
					NumberDuck.Secret.XFCRC __241720877 = pXFCRC;
					pXFCRC = null;
					pBiffRecordContainer.AddBiffRecord(__241720877);
				}
				for (ushort i = 0; i < (ushort)(pWorkbookGlobals.m_pStyleVector.GetSize()); i++)
				{
					Style pStyle = pWorkbookGlobals.m_pStyleVector.Get(i);
					pBiffRecordContainer.AddBiffRecord(new XFExt((ushort)(i + 15), pStyle.GetBackgroundColor(false), pStyle.GetFillPatternColor(false), pStyle.GetFont().GetColor(false)));
				}
				pBiffRecordContainer.AddBiffRecord(new PaletteRecord());
				for (int i = 0; i < pWorksheetVector.GetSize(); i++)
				{
					Worksheet pWorksheet = pWorksheetVector.Get(i);
					BoundSheet8Record pBoundSheet8Record = new BoundSheet8Record(pWorksheet.GetName());
					pBoundSheet8RecordVector.PushBack(pBoundSheet8Record);
					{
						NumberDuck.Secret.BoundSheet8Record __3176075103 = pBoundSheet8Record;
						pBoundSheet8Record = null;
						pBiffRecordContainer.AddBiffRecord(__3176075103);
					}
				}
				if (pWorkbookGlobals.m_pWorksheetRangeVector.GetSize() > 0)
				{
					pBiffRecordContainer.AddBiffRecord(new SupBookRecord((ushort)(pWorksheetVector.GetSize())));
					pBiffRecordContainer.AddBiffRecord(new ExternSheetRecord(pWorkbookGlobals.m_pWorksheetRangeVector));
				}
				{
					MsoDrawingGroupRecord pMsoDrawingGroupRecord = new MsoDrawingGroupRecord(pWorkbookGlobals.m_pSharedPictureVector);
					{
						NumberDuck.Secret.MsoDrawingGroupRecord __355026241 = pMsoDrawingGroupRecord;
						pMsoDrawingGroupRecord = null;
						pBiffRecordContainer.AddBiffRecord(__355026241);
					}
				}
				if (pWorkbookGlobals.m_pSharedStringContainer.GetSize() > 0)
				{
					SstRecord pSstRecord = new SstRecord(pWorkbookGlobals.m_pSharedStringContainer);
					{
						NumberDuck.Secret.SstRecord __1557921119 = pSstRecord;
						pSstRecord = null;
						pBiffRecordContainer.AddBiffRecord(__1557921119);
					}
				}
				pBiffRecordContainer.AddBiffRecord(new BookExtRecord());
				pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_EOF, 0));
				uint nStreamOffset = pBiffRecordContainer.GetSize();
				for (int i = 0; i < pBoundSheet8RecordVector.GetSize(); i++)
				{
					pBoundSheet8RecordVector.Get(i).SetStreamOffset(nStreamOffset);
					nStreamOffset += pWorkbookGlobals.m_pBiffWorksheetStreamSizeVector.Get(i).m_nSize;
				}
				pBiffRecordContainer.Write(pStream);
				{
					pBiffRecordContainer = null;
				}
				{
					pBoundSheet8RecordVector = null;
				}
			}

			public string GetNextWorksheetName()
			{
				if (m_sTempWorksheetName != null)
					{
						m_sTempWorksheetName = null;
					}
				m_sTempWorksheetName = m_sWorkheetNameVector.PopFront();
				return m_sTempWorksheetName.GetExternalString();
			}

			public Style GetStyleByXfIndex(ushort nXfIndex)
			{
				nbAssert.Assert(nXfIndex >= 15);
				Style pStyle = GetStyleByIndex((ushort)(nXfIndex - 15));
				return pStyle;
			}

			public MsoDrawingGroupRecord GetMsoDrawingGroupRecord()
			{
				return m_MsoDrawingGroupRecord;
			}

			public static byte SnapToPalette(Color pColor)
			{
				byte nIndex = 0;
				int nDiff = 2147483647;
				if (pColor != null)
				{
					for (byte i = 0; i < NUM_CUSTOM_PALETTE_ENTRY; i++)
					{
						uint nTestColor = DEFAULT_CUSTOM_COLOR[i];
						int nRedDiff = ((int)(nTestColor & 0xFF)) - ((int)(pColor.GetRed()));
						int nGreenDiff = ((int)((nTestColor >> 8) & 0xFF)) - ((int)(pColor.GetGreen()));
						int nBlueDiff = ((int)((nTestColor >> 16) & 0xFF)) - ((int)(pColor.GetBlue()));
						int nTestDiff = nRedDiff * nRedDiff + nGreenDiff * nGreenDiff + nBlueDiff * nBlueDiff;
						if (nTestDiff < nDiff)
						{
							nIndex = i;
							nDiff = nTestDiff;
						}
					}
					return (byte)(nIndex + NUM_DEFAULT_PALETTE_ENTRY);
				}
				return (byte)(PALETTE_INDEX_DEFAULT);
			}

			public static uint GetDefaultPaletteColorByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY);
				uint nColor = 0;
				if (nIndex < NUM_DEFAULT_PALETTE_ENTRY)
					nColor = DEFAULT_COLOR[nIndex];
				else
					nColor = DEFAULT_CUSTOM_COLOR[nIndex - NUM_DEFAULT_PALETTE_ENTRY];
				return nColor;
			}

			public uint GetPaletteColorByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex < NUM_DEFAULT_PALETTE_ENTRY + NUM_CUSTOM_PALETTE_ENTRY);
				if (m_pPaletteRecord != null && nIndex >= NUM_DEFAULT_PALETTE_ENTRY)
					return m_pPaletteRecord.GetColorByIndex((ushort)(nIndex - NUM_DEFAULT_PALETTE_ENTRY));
				return GetDefaultPaletteColorByIndex(nIndex);
			}

			~BiffWorkbookGlobals()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffStruct
		{
			public enum FontScheme
			{
				XFSNONE = 0x00,
				XFSMAJOR = 0x01,
				XFSMINOR = 0x02,
				XFSNIL = 0xFF,
			}

			public enum XColorType
			{
				XCLRAUTO = 0x00000000,
				XCLRINDEXED = 0x00000001,
				XCLRRGB = 0x00000002,
				XCLRTHEMED = 0x00000003,
				XCLRNINCHED = 0x00000004,
			}

			public virtual void BlobRead(BlobView pBlobView)
			{
			}

			public virtual void BlobWrite(BlobView pBlobView)
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtCmoStruct : BiffStruct
		{
			public FtCmoStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ft = pBlobView.UnpackUint16();
				m_cb = pBlobView.UnpackUint16();
				m_ot = pBlobView.UnpackUint16();
				m_id = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fLocked = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fDefaultSize = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fPublished = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fPrint = (ushort)((nBitmask0 >> 4) & 0x1);
				m_unused1 = (ushort)((nBitmask0 >> 5) & 0x1);
				m_unused2 = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fDisabled = (ushort)((nBitmask0 >> 7) & 0x1);
				m_fUIObj = (ushort)((nBitmask0 >> 8) & 0x1);
				m_fRecalcObj = (ushort)((nBitmask0 >> 9) & 0x1);
				m_unused3 = (ushort)((nBitmask0 >> 10) & 0x1);
				m_unused4 = (ushort)((nBitmask0 >> 11) & 0x1);
				m_fRecalcObjAlways = (ushort)((nBitmask0 >> 12) & 0x1);
				m_unused5 = (ushort)((nBitmask0 >> 13) & 0x1);
				m_unused6 = (ushort)((nBitmask0 >> 14) & 0x1);
				m_unused7 = (ushort)((nBitmask0 >> 15) & 0x1);
				m_unused8 = pBlobView.UnpackUint32();
				m_unused9 = pBlobView.UnpackUint32();
				m_unused10 = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ft);
				pBlobView.PackUint16(m_cb);
				pBlobView.PackUint16(m_ot);
				pBlobView.PackUint16(m_id);
				int nBitmask0 = 0;
				nBitmask0 += m_fLocked << 0;
				nBitmask0 += m_reserved << 1;
				nBitmask0 += m_fDefaultSize << 2;
				nBitmask0 += m_fPublished << 3;
				nBitmask0 += m_fPrint << 4;
				nBitmask0 += m_unused1 << 5;
				nBitmask0 += m_unused2 << 6;
				nBitmask0 += m_fDisabled << 7;
				nBitmask0 += m_fUIObj << 8;
				nBitmask0 += m_fRecalcObj << 9;
				nBitmask0 += m_unused3 << 10;
				nBitmask0 += m_unused4 << 11;
				nBitmask0 += m_fRecalcObjAlways << 12;
				nBitmask0 += m_unused5 << 13;
				nBitmask0 += m_unused6 << 14;
				nBitmask0 += m_unused7 << 15;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint32(m_unused8);
				pBlobView.PackUint32(m_unused9);
				pBlobView.PackUint32(m_unused10);
			}

			public const ushort SIZE = 22;
			public void SetDefaults()
			{
				m_ft = 0;
				m_cb = 0;
				m_ot = 0;
				m_id = 0;
				m_fLocked = 0;
				m_reserved = 0;
				m_fDefaultSize = 0;
				m_fPublished = 0;
				m_fPrint = 0;
				m_unused1 = 0;
				m_unused2 = 0;
				m_fDisabled = 0;
				m_fUIObj = 0;
				m_fRecalcObj = 0;
				m_unused3 = 0;
				m_unused4 = 0;
				m_fRecalcObjAlways = 0;
				m_unused5 = 0;
				m_unused6 = 0;
				m_unused7 = 0;
				m_unused8 = 0;
				m_unused9 = 0;
				m_unused10 = 0;
			}

			public ushort m_ft;
			public ushort m_cb;
			public ushort m_ot;
			public ushort m_id;
			public ushort m_fLocked;
			public ushort m_reserved;
			public ushort m_fDefaultSize;
			public ushort m_fPublished;
			public ushort m_fPrint;
			public ushort m_unused1;
			public ushort m_unused2;
			public ushort m_fDisabled;
			public ushort m_fUIObj;
			public ushort m_fRecalcObj;
			public ushort m_unused3;
			public ushort m_unused4;
			public ushort m_fRecalcObjAlways;
			public ushort m_unused5;
			public ushort m_unused6;
			public ushort m_unused7;
			public uint m_unused8;
			public uint m_unused9;
			public uint m_unused10;
			public enum ObjType
			{
				OBJ_TYPE_GROUP = 0x0000,
				OBJ_TYPE_LINE = 0x0001,
				OBJ_TYPE_RECTANGLE = 0x0002,
				OBJ_TYPE_OVAL = 0x0003,
				OBJ_TYPE_ARC = 0x0004,
				OBJ_TYPE_CHART = 0x0005,
				OBJ_TYPE_TEXT = 0x0006,
				OBJ_TYPE_BUTTON = 0x0007,
				OBJ_TYPE_PICTURE = 0x0008,
				OBJ_TYPE_POLYGON = 0x0009,
				OBJ_TYPE_CHECKBOX = 0x000B,
				OBJ_TYPE_RADIO_BUTTON = 0x000C,
				OBJ_TYPE_EDIT_BOX = 0x000D,
				OBJ_TYPE_LABEL = 0x000E,
				OBJ_TYPE_DIALOG_BOX = 0x000F,
				OBJ_TYPE_SPIN_CONTROL = 0x0010,
				OBJ_TYPE_SCROLLBAR = 0x0011,
				OBJ_TYPE_LIST = 0x0012,
				OBJ_TYPE_GROUP_BOX = 0x0013,
				OBJ_TYPE_DROPDOWN_LIST = 0x0014,
				OBJ_TYPE_NOTE = 0x0019,
				OBJ_TYPE_OFFICE_ART_OBJECT = 0x001E,
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorChain
		{
			protected int m_nSectorSize;
			protected Vector<Sector> m_pSectorVector;
			protected Blob m_pBlob;
			public SectorChain(int nSectorSize)
			{
				nbAssert.Assert((nSectorSize & (nSectorSize - 1)) == 0);
				m_nSectorSize = nSectorSize;
				m_pBlob = new Blob(false);
				m_pSectorVector = new Vector<Sector>();
			}

			public int GetNumSector()
			{
				return m_pSectorVector.GetSize();
			}

			public Sector GetSectorByIndex(int nIndex)
			{
				nbAssert.Assert(nIndex < m_pSectorVector.GetSize());
				return m_pSectorVector.Get(nIndex);
			}

			public int GetSectorSize()
			{
				return m_nSectorSize;
			}

			public int GetDataSize()
			{
				return m_pSectorVector.GetSize() * m_nSectorSize;
			}

			public void Read(int nSize, BlobView pBlobView)
			{
				m_pBlob.GetBlobView().Unpack(pBlobView, nSize);
			}

			public void ReadAt(int nOffset, int nSize, BlobView pBlobView)
			{
				m_pBlob.GetBlobView().UnpackAt(nOffset, pBlobView, nSize);
			}

			public void Write(int nSize, BlobView pBlobView)
			{
				m_pBlob.GetBlobView().Pack(pBlobView, nSize);
			}

			public void WriteAt(int nOffset, int nSize, BlobView pBlobView)
			{
				m_pBlob.GetBlobView().PackAt(nOffset, pBlobView, nSize);
			}

			public int GetOffset()
			{
				return m_pBlob.GetBlobView().GetOffset();
			}

			public void SetOffset(int nOffset)
			{
				m_pBlob.GetBlobView().SetOffset(nOffset);
			}

			public virtual void AppendSector(Sector pSector)
			{
				BlobView pBlobView = pSector.GetBlobView();
				m_pSectorVector.PushBack(pSector);
				int nSize = m_pSectorVector.GetSize() * m_nSectorSize;
				m_pBlob.Resize(nSize, false);
				pBlobView.SetOffset(0);
				m_pBlob.GetBlobView().PackAt(nSize - m_nSectorSize, pBlobView, m_nSectorSize);
			}

			public virtual void Extend(Sector pSector)
			{
				AppendSector(pSector);
			}

			public BlobView GetBlobView()
			{
				return m_pBlob.GetBlobView();
			}

			public void WriteToSectors()
			{
				BlobView pBlobView = m_pBlob.GetBlobView();
				pBlobView.SetOffset(0);
				for (int i = 0; i < m_pSectorVector.GetSize(); i++)
					m_pSectorVector.Get(i).GetBlobView().PackAt(0, pBlobView, m_nSectorSize);
			}

			~SectorChain()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDataStruct
		{
			public const int MAX_NAME_LENGTH = 32;
			public ushort[] m_pName = new ushort[MAX_NAME_LENGTH];
			public ushort m_nNameDataSize;
			public byte m_nType;
			public byte m_nNodeColour;
			public int m_nLeftChildNodeStreamId;
			public int m_nRightChildNodeStreamId;
			public int m_nRootNodeStreamId;
			public byte[] m_pUniqueIdentifier = new byte[16];
			public byte[] m_pUserFlags = new byte[4];
			public byte[] m_pCreationDate = new byte[8];
			public byte[] m_pModificationDate = new byte[8];
			public int m_nSectorId;
			public uint m_nStreamSize;
			public byte[] m_pUnused = new byte[4];
		}
		class Stream
		{
			public enum Type
			{
				TYPE_EMPTY = 0,
				TYPE_USER_STORAGE,
				TYPE_USER_STREAM,
				TYPE_LOCK_BYTES,
				TYPE_PROPERTY,
				TYPE_ROOT_STORAGE,
			}

			public const int DATA_SIZE = 128;
			protected int m_nStreamId;
			protected int m_nMinimumStandardStreamSize;
			protected StreamDataStruct m_pDataStruct;
			protected BlobView m_pBlobView;
			protected SectorChain m_pSectorChain;
			protected InternalString m_sNameTemp;
			protected CompoundFile m_pCompoundFile;
			public Stream(int nStreamId, int nMinimumStandardStreamSize, Blob pBlob, int nOffset, CompoundFile pCompoundFile)
			{
				int i;
				nbAssert.Assert(pCompoundFile != null);
				m_pDataStruct = new StreamDataStruct();
				m_pCompoundFile = pCompoundFile;
				m_nStreamId = nStreamId;
				m_nMinimumStandardStreamSize = nMinimumStandardStreamSize;
				m_sNameTemp = new InternalString("");
				m_pBlobView = new BlobView(pBlob, nOffset, nOffset + DATA_SIZE);
				for (i = 0; i < StreamDataStruct.MAX_NAME_LENGTH; i++)
					m_pDataStruct.m_pName[i] = m_pBlobView.UnpackUint16();
				m_pDataStruct.m_nNameDataSize = m_pBlobView.UnpackUint16();
				m_pDataStruct.m_nType = m_pBlobView.UnpackUint8();
				m_pDataStruct.m_nNodeColour = m_pBlobView.UnpackUint8();
				m_pDataStruct.m_nLeftChildNodeStreamId = m_pBlobView.UnpackInt32();
				m_pDataStruct.m_nRightChildNodeStreamId = m_pBlobView.UnpackInt32();
				m_pDataStruct.m_nRootNodeStreamId = m_pBlobView.UnpackInt32();
				for (i = 0; i < 16; i++)
					m_pDataStruct.m_pUniqueIdentifier[i] = m_pBlobView.UnpackUint8();
				for (i = 0; i < 4; i++)
					m_pDataStruct.m_pUserFlags[i] = m_pBlobView.UnpackUint8();
				for (i = 0; i < 8; i++)
					m_pDataStruct.m_pCreationDate[i] = m_pBlobView.UnpackUint8();
				for (i = 0; i < 8; i++)
					m_pDataStruct.m_pModificationDate[i] = m_pBlobView.UnpackUint8();
				m_pDataStruct.m_nSectorId = m_pBlobView.UnpackInt32();
				m_pDataStruct.m_nStreamSize = m_pBlobView.UnpackUint32();
				for (i = 0; i < 4; i++)
					m_pDataStruct.m_pUnused[i] = m_pBlobView.UnpackUint8();
				nbAssert.Assert(m_pBlobView.GetOffset() == m_pBlobView.GetSize());
				nbAssert.Assert(m_pBlobView.GetOffset() == DATA_SIZE);
				if (m_nStreamId == 0)
				{
					m_pDataStruct.m_nType = (byte)(Type.TYPE_ROOT_STORAGE);
				}
				m_pSectorChain = null;
				if (m_pDataStruct.m_nType != (byte)(Type.TYPE_EMPTY))
					m_pSectorChain = new SectorChain(m_pCompoundFile.GetSectorSize(GetShortSector()));
			}

			public void Allocate(Type eType, int nStreamSize)
			{
				nbAssert.Assert(m_pSectorChain == null);
				nbAssert.Assert(m_pDataStruct.m_nType == (byte)(Type.TYPE_EMPTY));
				nbAssert.Assert(eType != Type.TYPE_ROOT_STORAGE);
				m_pDataStruct.m_nType = (byte)(eType);
				m_pSectorChain = new SectorChain(m_pCompoundFile.GetSectorSize(GetShortSector()));
				Resize(nStreamSize);
			}

			public void FillSectorChain()
			{
				m_pCompoundFile.FillSectorChain(m_pSectorChain, GetSectorId(), GetShortSector());
			}

			public void WriteToSectors()
			{
				int i;
				if (m_pSectorChain != null)
					m_pSectorChain.WriteToSectors();
				m_pBlobView.SetOffset(0);
				for (i = 0; i < StreamDataStruct.MAX_NAME_LENGTH; i++)
					m_pBlobView.PackUint16(m_pDataStruct.m_pName[i]);
				m_pBlobView.PackUint16(m_pDataStruct.m_nNameDataSize);
				m_pBlobView.PackUint8(m_pDataStruct.m_nType);
				m_pBlobView.PackUint8(m_pDataStruct.m_nNodeColour);
				m_pBlobView.PackInt32(m_pDataStruct.m_nLeftChildNodeStreamId);
				m_pBlobView.PackInt32(m_pDataStruct.m_nRightChildNodeStreamId);
				m_pBlobView.PackInt32(m_pDataStruct.m_nRootNodeStreamId);
				for (i = 0; i < 16; i++)
					m_pBlobView.PackUint8(m_pDataStruct.m_pUniqueIdentifier[i]);
				for (i = 0; i < 4; i++)
					m_pBlobView.PackUint8(m_pDataStruct.m_pUserFlags[i]);
				for (i = 0; i < 8; i++)
					m_pBlobView.PackUint8(m_pDataStruct.m_pCreationDate[i]);
				for (i = 0; i < 8; i++)
					m_pBlobView.PackUint8(m_pDataStruct.m_pModificationDate[i]);
				m_pBlobView.PackInt32(m_pDataStruct.m_nSectorId);
				m_pBlobView.PackUint32(m_pDataStruct.m_nStreamSize);
				for (i = 0; i < 4; i++)
					m_pBlobView.PackUint8(m_pDataStruct.m_pUnused[i]);
				nbAssert.Assert(m_pBlobView.GetOffset() == m_pBlobView.GetSize());
				nbAssert.Assert(m_pBlobView.GetOffset() == DATA_SIZE);
			}

			public int GetStreamId()
			{
				return m_nStreamId;
			}

			public string GetName()
			{
				ushort i = 0;
				m_sNameTemp.Set("");
				while (i << 1 < m_pDataStruct.m_nNameDataSize)
				{
					ushort nTemp = m_pDataStruct.m_pName[i];
					if (nTemp == 0)
						break;
					m_sNameTemp.AppendChar((char)(nTemp));
					i++;
				}
				return m_sNameTemp.GetExternalString();
			}

			public void SetName(string sxName)
			{
				int i;
				m_sNameTemp.Set(sxName);
				int nLength = m_sNameTemp.GetLength();
				nbAssert.Assert(nLength < StreamDataStruct.MAX_NAME_LENGTH);
				m_pDataStruct.m_nNameDataSize = (ushort)((nLength + 1) << 1);
				for (i = 0; i < nLength; i++)
					m_pDataStruct.m_pName[i] = m_sNameTemp.GetChar(i);
				m_pDataStruct.m_pName[nLength] = 0;
			}

			public new Stream.Type GetType()
			{
				return (Type)(m_pDataStruct.m_nType);
			}

			public int GetSectorId()
			{
				return m_pDataStruct.m_nSectorId;
			}

			public bool GetShortSector()
			{
				if (GetType() == Type.TYPE_ROOT_STORAGE || (int)(m_pDataStruct.m_nStreamSize) >= m_nMinimumStandardStreamSize)
					return false;
				return true;
			}

			public int GetStreamSize()
			{
				return (int)(m_pDataStruct.m_nStreamSize);
			}

			public int GetSize()
			{
				return m_pSectorChain.GetDataSize();
			}

			public SectorChain GetSectorChain()
			{
				return m_pSectorChain;
			}

			public int GetOffset()
			{
				return m_pSectorChain.GetOffset();
			}

			public void SetOffset(int nOffset)
			{
				m_pSectorChain.SetOffset(nOffset);
			}

			public void SizeToFit(int nSize)
			{
				if (GetStreamSize() < m_pSectorChain.GetOffset() + nSize)
					Resize(m_pSectorChain.GetOffset() + nSize);
			}

			public bool Resize(int nSize)
			{
				if (GetType() == Type.TYPE_EMPTY)
					return false;
				int nOldSize = GetStreamSize();
				m_pDataStruct.m_nStreamSize = (uint)(nSize);
				bool bShortSector = GetShortSector();
				int nSectorSize = m_pCompoundFile.GetSectorSize(bShortSector);
				int nNumSector = nSize / nSectorSize + 1;
				if (m_pSectorChain == null)
					m_pSectorChain = new SectorChain(nSectorSize);
				if (nSectorSize != m_pSectorChain.GetSectorSize())
				{
					SectorChain pSectorChain = new SectorChain(nSectorSize);
					while (pSectorChain.GetNumSector() < nNumSector)
					{
						m_pCompoundFile.SectorChainExtend(pSectorChain);
					}
					BlobView pBlobView = m_pSectorChain.GetBlobView();
					pBlobView.SetOffset(0);
					int nDataSize = nOldSize;
					if (nSize < nOldSize)
						nDataSize = nSize;
					pSectorChain.Write(nDataSize, pBlobView);
					SectorAllocationTable pSectorAllocationTable = m_pCompoundFile.GetSectorAllocationTable(!bShortSector);
					for (int i = 0; i < m_pSectorChain.GetNumSector(); i++)
						pSectorAllocationTable.SetSectorId(m_pSectorChain.GetSectorByIndex(i).GetSectorId(), (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID));
					pSectorAllocationTable = m_pCompoundFile.GetSectorAllocationTable(bShortSector);
					for (int i = 0; i < pSectorChain.GetNumSector() - 1; i++)
						pSectorAllocationTable.SetSectorId(pSectorChain.GetSectorByIndex(i).GetSectorId(), pSectorChain.GetSectorByIndex(i + 1).GetSectorId());
					pSectorAllocationTable.SetSectorId(pSectorChain.GetSectorByIndex(pSectorChain.GetNumSector() - 1).GetSectorId(), (int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID));
					{
						m_pSectorChain = null;
					}
					{
						NumberDuck.Secret.SectorChain __175211134 = pSectorChain;
						pSectorChain = null;
						m_pSectorChain = __175211134;
					}
				}
				else
				{
					while (m_pSectorChain.GetNumSector() < nNumSector)
					{
						m_pCompoundFile.SectorChainExtend(m_pSectorChain);
					}
				}
				m_pDataStruct.m_nSectorId = m_pSectorChain.GetSectorByIndex(0).GetSectorId();
				return true;
			}

			public void SetNodeColour(byte nColour)
			{
				m_pDataStruct.m_nNodeColour = nColour;
			}

			public void SetLeftChildNodeStreamId(int nStreamId)
			{
				m_pDataStruct.m_nLeftChildNodeStreamId = nStreamId;
			}

			public void SetRightChildNodeStreamId(int nStreamId)
			{
				m_pDataStruct.m_nRightChildNodeStreamId = nStreamId;
			}

			public void SetRootNodeStreamId(int nStreamId)
			{
				m_pDataStruct.m_nRootNodeStreamId = nStreamId;
			}

			public ushort GetNameLengthUtf16()
			{
				if (m_pDataStruct.m_nNameDataSize == 0)
					return 0;
				return (ushort)((m_pDataStruct.m_nNameDataSize >> 1) - 1);
			}

			public ushort GetNameUtf16(ushort nIndex)
			{
				nbAssert.Assert(nIndex < StreamDataStruct.MAX_NAME_LENGTH);
				return m_pDataStruct.m_pName[nIndex];
			}

			~Stream()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtTertiaryFOPTRecord : OfficeArtRecord
		{
			public OfficeArtTertiaryFOPTRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_TERTIARY_FOPT);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_TERTIARY_FOPT;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OwnedVector<OfficeArtFOPTEStruct> m_pFoptVector;
			public OfficeArtTertiaryFOPTRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x3;
			}

			protected void PostSetDefaults()
			{
				m_pFoptVector = new OwnedVector<OfficeArtFOPTEStruct>();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
					m_pFoptVector.Get(i).BlobWrite(pBlobView);
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if (pFOPTE.m_pComplexData != null)
					{
						BlobView pComplexDataBlobView = pFOPTE.m_pComplexData.GetBlobView();
						pComplexDataBlobView.SetOffset(0);
						pBlobView.Pack(pComplexDataBlobView, pComplexDataBlobView.GetSize());
					}
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				ushort i;
				for (i = 0; i < m_pHeader.m_recInstance; i++)
				{
					OfficeArtFOPTEStruct pFOPTE = new OfficeArtFOPTEStruct();
					pFOPTE.BlobRead(pBlobView);
					{
						NumberDuck.Secret.OfficeArtFOPTEStruct __3616310584 = pFOPTE;
						pFOPTE = null;
						m_pFoptVector.PushBack(__3616310584);
					}
				}
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if (pFOPTE.m_opid.m_fComplex == 0x1)
					{
						pFOPTE.m_pComplexData = new Blob(false);
						pFOPTE.m_pComplexData.Resize(pFOPTE.m_op, false);
						pFOPTE.m_pComplexData.GetBlobView().Pack(pBlobView, pFOPTE.m_op);
					}
				}
			}

			public void AddProperty(ushort opid, byte fBid, int op)
			{
				nbAssert.Assert(opid <= 0x3FFF);
				nbAssert.Assert(fBid <= 0x1);
				OfficeArtFOPTEStruct pFOPTE = new OfficeArtFOPTEStruct();
				pFOPTE.m_opid.m_opid = opid;
				pFOPTE.m_opid.m_fBid = fBid;
				pFOPTE.m_opid.m_fComplex = 0x0;
				pFOPTE.m_op = op;
				{
					NumberDuck.Secret.OfficeArtFOPTEStruct __3616310584 = pFOPTE;
					pFOPTE = null;
					m_pFoptVector.PushBack(__3616310584);
				}
				m_pHeader.m_recInstance++;
				m_pHeader.m_recLen += OfficeArtFOPTEStruct.SIZE;
			}

			public OfficeArtFOPTEStruct GetProperty(OfficeArtRecord.OPIDType eType)
			{
				ushort i;
				for (i = 0; i < m_pHeader.m_recInstance; i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if ((OfficeArtRecord.OPIDType)(pFOPTE.m_opid.m_opid) == eType)
						return pFOPTE;
				}
				return null;
			}

			~OfficeArtTertiaryFOPTRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSplitMenuColorContainerRecord : OfficeArtRecord
		{
			public OfficeArtSplitMenuColorContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_smca0.BlobRead(pBlobView);
				m_smca1.BlobRead(pBlobView);
				m_smca2.BlobRead(pBlobView);
				m_smca3.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_smca0.BlobWrite(pBlobView);
				m_smca1.BlobWrite(pBlobView);
				m_smca2.BlobWrite(pBlobView);
				m_smca3.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 16;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_SPLIT_MENU_COLOR_CONTAINER;
			protected void SetDefaults()
			{
				m_smca0 = new MSOCRStruct();
				m_smca1 = new MSOCRStruct();
				m_smca2 = new MSOCRStruct();
				m_smca3 = new MSOCRStruct();
			}

			protected MSOCRStruct m_smca0;
			protected MSOCRStruct m_smca1;
			protected MSOCRStruct m_smca2;
			protected MSOCRStruct m_smca3;
			public OfficeArtSplitMenuColorContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x0;
				m_pHeader.m_recInstance = 4;
				m_smca0.m_red = 13;
				m_smca0.m_green = 0;
				m_smca0.m_blue = 0;
				m_smca0.m_fSchemeIndex = 1;
				m_smca1.m_red = 12;
				m_smca1.m_green = 0;
				m_smca1.m_blue = 0;
				m_smca1.m_fSchemeIndex = 1;
				m_smca2.m_red = 23;
				m_smca2.m_green = 0;
				m_smca2.m_blue = 0;
				m_smca2.m_fSchemeIndex = 1;
				m_smca3.m_red = 247;
				m_smca3.m_green = 0;
				m_smca3.m_blue = 0;
				m_smca3.m_unused2 = 1;
			}

			~OfficeArtSplitMenuColorContainerRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSpgrContainerRecord : OfficeArtRecord
		{
			public OfficeArtSpgrContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_SPGR_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = true;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_SPGR_CONTAINER;
			protected void SetDefaults()
			{
			}

			public OfficeArtSpgrContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0xF;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtSpContainerRecord : OfficeArtRecord
		{
			public OfficeArtSpContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_SP_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = true;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_SP_CONTAINER;
			protected void SetDefaults()
			{
			}

			public OfficeArtSpContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0xF;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFSPRecord : OfficeArtRecord
		{
			public OfficeArtFSPRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FSP);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_spid = pBlobView.UnpackUint32();
				uint nBitmask0 = pBlobView.UnpackUint32();
				m_fGroup = (uint)((nBitmask0 >> 0) & 0x1);
				m_fChild = (uint)((nBitmask0 >> 1) & 0x1);
				m_fPatriarch = (uint)((nBitmask0 >> 2) & 0x1);
				m_fDeleted = (uint)((nBitmask0 >> 3) & 0x1);
				m_fOleShape = (uint)((nBitmask0 >> 4) & 0x1);
				m_fHaveMaster = (uint)((nBitmask0 >> 5) & 0x1);
				m_fFlipH = (uint)((nBitmask0 >> 6) & 0x1);
				m_fFlipV = (uint)((nBitmask0 >> 7) & 0x1);
				m_fConnector = (uint)((nBitmask0 >> 8) & 0x1);
				m_fHaveAnchor = (uint)((nBitmask0 >> 9) & 0x1);
				m_fBackground = (uint)((nBitmask0 >> 10) & 0x1);
				m_fHaveSpt = (uint)((nBitmask0 >> 11) & 0x1);
				m_unused1 = (uint)((nBitmask0 >> 12) & 0xfffff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_spid);
				uint nBitmask0 = 0;
				nBitmask0 += m_fGroup << 0;
				nBitmask0 += m_fChild << 1;
				nBitmask0 += m_fPatriarch << 2;
				nBitmask0 += m_fDeleted << 3;
				nBitmask0 += m_fOleShape << 4;
				nBitmask0 += m_fHaveMaster << 5;
				nBitmask0 += m_fFlipH << 6;
				nBitmask0 += m_fFlipV << 7;
				nBitmask0 += m_fConnector << 8;
				nBitmask0 += m_fHaveAnchor << 9;
				nBitmask0 += m_fBackground << 10;
				nBitmask0 += m_fHaveSpt << 11;
				nBitmask0 += m_unused1 << 12;
				pBlobView.PackUint32((uint)(nBitmask0));
			}

			protected const ushort SIZE = 8;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FSP;
			protected void SetDefaults()
			{
				m_spid = 0;
				m_fGroup = 0;
				m_fChild = 0;
				m_fPatriarch = 0;
				m_fDeleted = 0;
				m_fOleShape = 0;
				m_fHaveMaster = 0;
				m_fFlipH = 0;
				m_fFlipV = 0;
				m_fConnector = 0;
				m_fHaveAnchor = 0;
				m_fBackground = 0;
				m_fHaveSpt = 0;
				m_unused1 = 0;
			}

			protected uint m_spid;
			protected uint m_fGroup;
			protected uint m_fChild;
			protected uint m_fPatriarch;
			protected uint m_fDeleted;
			protected uint m_fOleShape;
			protected uint m_fHaveMaster;
			protected uint m_fFlipH;
			protected uint m_fFlipV;
			protected uint m_fConnector;
			protected uint m_fHaveAnchor;
			protected uint m_fBackground;
			protected uint m_fHaveSpt;
			protected uint m_unused1;
			public OfficeArtFSPRecord(ushort recInstance, uint spid, byte fGroup, byte fPatriarch, byte fHaveAnchor, byte fHaveSpt) : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x2;
				m_pHeader.m_recInstance = recInstance;
				m_fGroup = fGroup;
				m_fPatriarch = fPatriarch;
				m_fHaveAnchor = fHaveAnchor;
				m_fHaveSpt = fHaveSpt;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFSPGRRecord : OfficeArtRecord
		{
			public OfficeArtFSPGRRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FSPGR);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_xLeft = pBlobView.UnpackUint32();
				m_yTop = pBlobView.UnpackUint32();
				m_xRight = pBlobView.UnpackUint32();
				m_yBottom = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_xLeft);
				pBlobView.PackUint32(m_yTop);
				pBlobView.PackUint32(m_xRight);
				pBlobView.PackUint32(m_yBottom);
			}

			protected const ushort SIZE = 16;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FSPGR;
			protected void SetDefaults()
			{
				m_xLeft = 0;
				m_yTop = 0;
				m_xRight = 0;
				m_yBottom = 0;
			}

			protected uint m_xLeft;
			protected uint m_yTop;
			protected uint m_xRight;
			protected uint m_yBottom;
			public OfficeArtFSPGRRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x1;
				m_pHeader.m_recInstance = 0x0000;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFRITContainerRecord : OfficeArtRecord
		{
			public OfficeArtFRITContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FRIT_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FRIT_CONTAINER;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OwnedVector<OfficeArtFRITStruct> m_pRgfritVector;
			protected void PostSetDefaults()
			{
				m_pRgfritVector = new OwnedVector<OfficeArtFRITStruct>();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < m_pRgfritVector.GetSize(); i++)
				{
					m_pRgfritVector.Get(i).BlobWrite(pBlobView);
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				ushort i;
				nbAssert.Assert(GetNumFRIT() == GetSize() / OfficeArtFRITStruct.SIZE);
				for (i = 0; i < m_pHeader.m_recInstance; i++)
				{
					OfficeArtFRITStruct pFRIT = new OfficeArtFRITStruct();
					pFRIT.BlobRead(pBlobView);
					{
						NumberDuck.Secret.OfficeArtFRITStruct __2212815663 = pFRIT;
						pFRIT = null;
						m_pRgfritVector.PushBack(__2212815663);
					}
				}
			}

			public ushort GetNumFRIT()
			{
				return GetInstance();
			}

			public OfficeArtFRITStruct GetFRITByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex >= GetNumFRIT());
				return m_pRgfritVector.Get(nIndex);
			}

			~OfficeArtFRITContainerRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTRecord : OfficeArtRecord
		{
			public OfficeArtFOPTRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FOPT);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FOPT;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OwnedVector<OfficeArtFOPTEStruct> m_pFoptVector;
			public OfficeArtFOPTRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x3;
			}

			protected void PostSetDefaults()
			{
				m_pFoptVector = new OwnedVector<OfficeArtFOPTEStruct>();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
					m_pFoptVector.Get(i).BlobWrite(pBlobView);
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if (pFOPTE.m_pComplexData != null)
					{
						BlobView pComplexDataBlobView = pFOPTE.m_pComplexData.GetBlobView();
						pComplexDataBlobView.SetOffset(0);
						pBlobView.Pack(pComplexDataBlobView, pComplexDataBlobView.GetSize());
					}
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				ushort i;
				for (i = 0; i < m_pHeader.m_recInstance; i++)
				{
					OfficeArtFOPTEStruct pFOPTE = new OfficeArtFOPTEStruct();
					pFOPTE.BlobRead(pBlobView);
					{
						NumberDuck.Secret.OfficeArtFOPTEStruct __3616310584 = pFOPTE;
						pFOPTE = null;
						m_pFoptVector.PushBack(__3616310584);
					}
				}
				for (i = 0; i < m_pFoptVector.GetSize(); i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if (pFOPTE.m_opid.m_fComplex == 0x1)
					{
						pFOPTE.m_pComplexData = new Blob(false);
						pFOPTE.m_pComplexData.Resize(pFOPTE.m_op, false);
						pFOPTE.m_pComplexData.GetBlobView().Pack(pBlobView, pFOPTE.m_op);
					}
				}
			}

			public void AddProperty(ushort opid, byte fBid, int op)
			{
				nbAssert.Assert(opid <= 0x3FFF);
				nbAssert.Assert(fBid <= 0x1);
				OfficeArtFOPTEStruct pFOPTE = new OfficeArtFOPTEStruct();
				pFOPTE.m_opid.m_opid = opid;
				pFOPTE.m_opid.m_fBid = fBid;
				pFOPTE.m_opid.m_fComplex = 0x0;
				pFOPTE.m_op = op;
				{
					NumberDuck.Secret.OfficeArtFOPTEStruct __3616310584 = pFOPTE;
					pFOPTE = null;
					m_pFoptVector.PushBack(__3616310584);
				}
				m_pHeader.m_recInstance++;
				m_pHeader.m_recLen += OfficeArtFOPTEStruct.SIZE;
			}

			public void AddStringProperty(ushort opid, string szString)
			{
				nbAssert.Assert((OfficeArtRecord.OPIDType)(opid) == OfficeArtRecord.OPIDType.OPID_WZ_NAME);
				InternalString sTemp = new InternalString(szString);
				Blob pBlob = new Blob(true);
				BlobView pBlobView = pBlob.GetBlobView();
				sTemp.BlobWrite16Bit(pBlobView, true);
				AddBlobProperty(opid, 1, pBlob);
			}

			public void AddBlobProperty(ushort opid, byte fBid, Blob pBlob)
			{
				nbAssert.Assert(opid <= 0x3FFF);
				nbAssert.Assert(fBid <= 0x1);
				BlobView pBlobView = pBlob.GetBlobView();
				pBlobView.SetOffset(0);
				OfficeArtFOPTEStruct pFOPTE = new OfficeArtFOPTEStruct();
				pFOPTE.m_opid.m_opid = opid;
				pFOPTE.m_opid.m_fBid = fBid;
				pFOPTE.m_opid.m_fComplex = 0x1;
				pFOPTE.m_op = (int)(pBlob.GetSize());
				pFOPTE.m_pComplexData = new Blob(false);
				pFOPTE.m_pComplexData.Resize(pFOPTE.m_op, false);
				pFOPTE.m_pComplexData.GetBlobView().Pack(pBlobView, pBlobView.GetSize());
				m_pHeader.m_recInstance++;
				m_pHeader.m_recLen += (uint)(OfficeArtFOPTEStruct.SIZE + pFOPTE.m_op);
				{
					NumberDuck.Secret.OfficeArtFOPTEStruct __3616310584 = pFOPTE;
					pFOPTE = null;
					m_pFoptVector.PushBack(__3616310584);
				}
			}

			public OfficeArtFOPTEStruct GetProperty(OfficeArtRecord.OPIDType eType)
			{
				ushort i;
				for (i = 0; i < m_pHeader.m_recInstance; i++)
				{
					OfficeArtFOPTEStruct pFOPTE = m_pFoptVector.Get(i);
					if ((OfficeArtRecord.OPIDType)(pFOPTE.m_opid.m_opid) == eType)
						return pFOPTE;
				}
				return null;
			}

			~OfficeArtFOPTRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFDGRecord : OfficeArtRecord
		{
			public OfficeArtFDGRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FDG);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_csp = pBlobView.UnpackUint32();
				m_spidCur = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_csp);
				pBlobView.PackUint32(m_spidCur);
			}

			protected const ushort SIZE = 8;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FDG;
			protected void SetDefaults()
			{
				m_csp = 0;
				m_spidCur = 0;
			}

			protected uint m_csp;
			protected uint m_spidCur;
			public OfficeArtFDGRecord(ushort recInstance, uint csp, uint spidCur) : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x0;
				m_pHeader.m_recInstance = recInstance;
				m_csp = csp;
				m_spidCur = spidCur;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFDGGBlockRecord : OfficeArtRecord
		{
			public OfficeArtFDGGBlockRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FDGG_BLOCK);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_spidMax = pBlobView.UnpackUint32();
				m_cidcl = pBlobView.UnpackUint32();
				m_cspSaved = pBlobView.UnpackUint32();
				m_cdgSaved = pBlobView.UnpackUint32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_spidMax);
				pBlobView.PackUint32(m_cidcl);
				pBlobView.PackUint32(m_cspSaved);
				pBlobView.PackUint32(m_cdgSaved);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 16;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FDGG_BLOCK;
			protected void SetDefaults()
			{
				m_spidMax = 0;
				m_cidcl = 0;
				m_cspSaved = 0;
				m_cdgSaved = 0;
				PostSetDefaults();
			}

			protected uint m_spidMax;
			protected uint m_cidcl;
			protected uint m_cspSaved;
			protected uint m_cdgSaved;
			protected OwnedVector<OfficeArtIDCLStruct> m_pRgidclVector;
			public OfficeArtFDGGBlockRecord(uint nNumPicture) : base(TYPE, IS_CONTAINER, SIZE + OfficeArtIDCLStruct.SIZE, true)
			{
				SetDefaults();
				m_spidMax = 2049;
				m_cidcl = 2;
				m_cspSaved = nNumPicture + 1;
				m_cdgSaved = 1;
				OfficeArtIDCLStruct pIdcl = new OfficeArtIDCLStruct();
				pIdcl.m_dgid = 1;
				pIdcl.m_cspidCur = nNumPicture * 2;
				{
					NumberDuck.Secret.OfficeArtIDCLStruct __1286637680 = pIdcl;
					pIdcl = null;
					m_pRgidclVector.PushBack(__1286637680);
				}
			}

			protected void PostSetDefaults()
			{
				m_pRgidclVector = new OwnedVector<OfficeArtIDCLStruct>();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < m_pRgidclVector.GetSize(); i++)
				{
					m_pRgidclVector.Get(i).BlobWrite(pBlobView);
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < (int)(m_cidcl) - 1; i++)
				{
					OfficeArtIDCLStruct pIdcl = new OfficeArtIDCLStruct();
					pIdcl.BlobRead(pBlobView);
					{
						NumberDuck.Secret.OfficeArtIDCLStruct __1286637680 = pIdcl;
						pIdcl = null;
						m_pRgidclVector.PushBack(__1286637680);
					}
				}
			}

			~OfficeArtFDGGBlockRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFBSERecord : OfficeArtRecord
		{
			public OfficeArtFBSERecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_FBSE);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_btWin32 = pBlobView.UnpackUint8();
				m_btMacOS = pBlobView.UnpackUint8();
				m_rgbUid.BlobRead(pBlobView);
				m_tag = pBlobView.UnpackUint16();
				m_size = pBlobView.UnpackUint32();
				m_cRef = pBlobView.UnpackUint32();
				m_foDelay = pBlobView.UnpackUint32();
				m_unused1 = pBlobView.UnpackUint8();
				m_cbName = pBlobView.UnpackUint8();
				m_unused2 = pBlobView.UnpackUint8();
				m_unused3 = pBlobView.UnpackUint8();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_btWin32);
				pBlobView.PackUint8(m_btMacOS);
				m_rgbUid.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_tag);
				pBlobView.PackUint32(m_size);
				pBlobView.PackUint32(m_cRef);
				pBlobView.PackUint32(m_foDelay);
				pBlobView.PackUint8(m_unused1);
				pBlobView.PackUint8(m_cbName);
				pBlobView.PackUint8(m_unused2);
				pBlobView.PackUint8(m_unused3);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 36;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_FBSE;
			protected void SetDefaults()
			{
				m_btWin32 = 0;
				m_btMacOS = 0;
				m_rgbUid = new MD4DigestStruct();
				m_tag = 0;
				m_size = 0;
				m_cRef = 0;
				m_foDelay = 0;
				m_unused1 = 0;
				m_cbName = 0;
				m_unused2 = 0;
				m_unused3 = 0;
				PostSetDefaults();
			}

			protected byte m_btWin32;
			protected byte m_btMacOS;
			protected MD4DigestStruct m_rgbUid;
			protected ushort m_tag;
			protected uint m_size;
			protected uint m_cRef;
			protected uint m_foDelay;
			protected byte m_unused1;
			protected byte m_cbName;
			protected byte m_unused2;
			protected byte m_unused3;
			protected OfficeArtBlipRecord m_pEmbeddedBlip;
			public OfficeArtFBSERecord(Picture pPicture) : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pEmbeddedBlip = new OfficeArtBlipRecord(pPicture);
				uint nBlipSize = m_pEmbeddedBlip.GetRecursiveSize();
				m_pHeader.m_recLen = SIZE + nBlipSize;
				m_pHeader.m_recVer = 0x2;
				{
					BlobView pPictureBlobView = pPicture.GetBlob().GetBlobView();
					Blob pBlob = new Blob(true);
					BlobView pBlobView = pBlob.GetBlobView();
					MD4 pMd4 = new MD4();
					pPictureBlobView.SetOffset(0);
					pMd4.Process(pPictureBlobView);
					pMd4.BlobWrite(pBlobView);
					pBlobView.SetOffset(0);
					m_rgbUid.BlobRead(pBlobView);
				}
				m_tag = 0xFF;
				m_size = nBlipSize;
				m_cRef = 1;
				if (pPicture.GetFormat() == Picture.Format.FORMAT_PNG)
				{
					m_pHeader.m_recInstance = 0x006;
					m_btWin32 = 6;
					m_btMacOS = 6;
					m_unused2 = 6;
				}
				else if (pPicture.GetFormat() == Picture.Format.FORMAT_JPEG)
				{
					m_pHeader.m_recInstance = 0x005;
					m_btWin32 = 5;
					m_btMacOS = 5;
					m_unused2 = 2;
				}
				else if (pPicture.GetFormat() == Picture.Format.FORMAT_WMF)
				{
					m_pHeader.m_recInstance = 0x003;
					m_btWin32 = 3;
					m_btMacOS = 4;
					m_unused2 = 0;
				}
				else
				{
					nbAssert.Assert(false);
				}
			}

			protected void PostSetDefaults()
			{
				m_pEmbeddedBlip = null;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_pEmbeddedBlip.RecursiveWrite(pBlobView);
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				pBlobView.SetOffset(pBlobView.GetOffset() + m_cbName);
				OfficeArtRecord pOfficeArtRecord = OfficeArtRecord.CreateOfficeArtRecord(pBlobView);
				{
					NumberDuck.Secret.OfficeArtRecord __3533451309 = pOfficeArtRecord;
					pOfficeArtRecord = null;
					m_pEmbeddedBlip = (OfficeArtBlipRecord)(__3533451309);
				}
			}

			public OfficeArtBlipRecord GetEmbeddedBlip()
			{
				return m_pEmbeddedBlip;
			}

			~OfficeArtFBSERecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDggContainerRecord : OfficeArtRecord
		{
			public OfficeArtDggContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_DGG_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = true;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_DGG_CONTAINER;
			protected void SetDefaults()
			{
			}

			public OfficeArtDggContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0xF;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDgContainerRecord : OfficeArtRecord
		{
			public OfficeArtDgContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_DG_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = true;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_DG_CONTAINER;
			protected void SetDefaults()
			{
			}

			public OfficeArtDgContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0xF;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtClientDataRecord : OfficeArtRecord
		{
			public OfficeArtClientDataRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_CLIENT_DATA);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_CLIENT_DATA;
			protected void SetDefaults()
			{
			}

			public OfficeArtClientDataRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x0;
				m_pHeader.m_recInstance = 0x0000;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtClientAnchorSheetRecord : OfficeArtRecord
		{
			public OfficeArtClientAnchorSheetRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fMove = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fSize = (ushort)((nBitmask0 >> 1) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 2) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved3 = (ushort)((nBitmask0 >> 4) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 5) & 0x7ff);
				m_colL = pBlobView.UnpackUint16();
				m_dxL = pBlobView.UnpackUint16();
				m_rwT = pBlobView.UnpackUint16();
				m_dyT = pBlobView.UnpackUint16();
				m_colR = pBlobView.UnpackUint16();
				m_dxR = pBlobView.UnpackUint16();
				m_rwB = pBlobView.UnpackUint16();
				m_dyB = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fMove << 0;
				nBitmask0 += m_fSize << 1;
				nBitmask0 += m_reserved1 << 2;
				nBitmask0 += m_reserved2 << 3;
				nBitmask0 += m_reserved3 << 4;
				nBitmask0 += m_unused << 5;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_colL);
				pBlobView.PackUint16(m_dxL);
				pBlobView.PackUint16(m_rwT);
				pBlobView.PackUint16(m_dyT);
				pBlobView.PackUint16(m_colR);
				pBlobView.PackUint16(m_dxR);
				pBlobView.PackUint16(m_rwB);
				pBlobView.PackUint16(m_dyB);
			}

			protected const ushort SIZE = 18;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET;
			protected void SetDefaults()
			{
				m_fMove = 0;
				m_fSize = 0;
				m_reserved1 = 0;
				m_reserved2 = 0;
				m_reserved3 = 0;
				m_unused = 0;
				m_colL = 0;
				m_dxL = 0;
				m_rwT = 0;
				m_dyT = 0;
				m_colR = 0;
				m_dxR = 0;
				m_rwB = 0;
				m_dyB = 0;
			}

			protected ushort m_fMove;
			protected ushort m_fSize;
			protected ushort m_reserved1;
			protected ushort m_reserved2;
			protected ushort m_reserved3;
			protected ushort m_unused;
			protected ushort m_colL;
			protected ushort m_dxL;
			protected ushort m_rwT;
			protected ushort m_dyT;
			protected ushort m_colR;
			protected ushort m_dxR;
			protected ushort m_rwB;
			protected ushort m_dyB;
			public OfficeArtClientAnchorSheetRecord(ushort colL, ushort dxL, ushort rwT, ushort dyT, ushort colR, ushort dxR, ushort rwB, ushort dyB) : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0x0;
				m_pHeader.m_recInstance = 0x0000;
				m_colL = colL;
				m_dxL = dxL;
				m_rwT = rwT;
				m_dyT = dyT;
				m_colR = colR;
				m_dxR = dxR;
				m_rwB = rwB;
				m_dyB = dyB;
			}

			public ushort GetCellX1()
			{
				return m_colL;
			}

			public ushort GetSubCellX1()
			{
				return m_dxL;
			}

			public ushort GetCellY1()
			{
				return m_rwT;
			}

			public ushort GetSubCellY1()
			{
				return m_dyT;
			}

			public ushort GetCellX2()
			{
				return m_colR;
			}

			public ushort GetSubCellX2()
			{
				return m_dxR;
			}

			public ushort GetCellY2()
			{
				return m_rwB;
			}

			public ushort GetSubCellY2()
			{
				return m_dyB;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtBlipRecord : OfficeArtRecord
		{
			public OfficeArtBlipRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rgbUid1.BlobRead(pBlobView);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rgbUid1.BlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 16;
			protected const bool IS_CONTAINER = false;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_EMF;
			protected void SetDefaults()
			{
				m_rgbUid1 = new MD4DigestStruct();
				PostSetDefaults();
			}

			protected MD4DigestStruct m_rgbUid1;
			protected MD4DigestStruct m_rgbUid2;
			protected byte m_tag;
			protected Blob m_BLIPFileData_;
			protected Blob m_OwnedBLIPFileData;
			public OfficeArtBlipRecord(Picture pPicture) : base(TYPE, IS_CONTAINER, (uint)(SIZE + 1 + pPicture.GetBlob().GetSize()), true)
			{
				SetDefaults();
				if (pPicture.GetFormat() == Picture.Format.FORMAT_PNG)
				{
					m_pHeader.m_recType = (ushort)(OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_PNG);
					m_pHeader.m_recVer = 0x0;
					m_pHeader.m_recInstance = 0x6E0;
				}
				else if (pPicture.GetFormat() == Picture.Format.FORMAT_JPEG)
				{
					m_pHeader.m_recType = (ushort)(OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_JPEG);
					m_pHeader.m_recVer = 0x0;
					m_pHeader.m_recInstance = 0x46A;
				}
				else if (pPicture.GetFormat() == Picture.Format.FORMAT_WMF)
				{
					m_pHeader.m_recType = (ushort)(OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_WMF);
					m_pHeader.m_recVer = 0x0;
					m_pHeader.m_recInstance = 0x216;
				}
				else
				{
					nbAssert.Assert(false);
				}
				{
					BlobView pPictureBlobView = pPicture.GetBlob().GetBlobView();
					Blob pBlob = new Blob(true);
					BlobView pBlobView = pBlob.GetBlobView();
					MD4 pMd4 = new MD4();
					pPictureBlobView.SetOffset(0);
					pMd4.Process(pPictureBlobView);
					pMd4.BlobWrite(pBlobView);
					pBlobView.SetOffset(0);
					m_rgbUid1.BlobRead(pBlobView);
				}
				m_tag = 0xFF;
				m_BLIPFileData_ = pPicture.GetBlob();
			}

			protected void PostSetDefaults()
			{
				m_tag = 0;
				m_rgbUid2 = null;
				m_BLIPFileData_ = null;
				m_OwnedBLIPFileData = null;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_rgbUid2 != null)
					m_rgbUid2.BlobWrite(pBlobView);
				pBlobView.PackUint8(m_tag);
				if (m_BLIPFileData_ != null)
				{
					BlobView pBLIPBlobView = m_BLIPFileData_.GetBlobView();
					pBLIPBlobView.SetOffset(0);
					pBlobView.Pack(pBLIPBlobView, pBLIPBlobView.GetSize());
				}
				if (m_OwnedBLIPFileData != null)
				{
					BlobView pBLIPBlobView = m_OwnedBLIPFileData.GetBlobView();
					pBLIPBlobView.SetOffset(0);
					pBlobView.Pack(pBLIPBlobView, pBLIPBlobView.GetSize());
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert((Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_EMF || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_WMF || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_PICT || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_JPEG || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_PNG || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_DIB || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_TIFF || (Type)(m_pHeader.m_recType) == Type.TYPE_OFFICE_ART_BLIP_JPEG_CMYK);
				uint nSize = m_pHeader.m_recLen - SIZE;
				if (m_pHeader.m_recInstance == 0x46B || m_pHeader.m_recInstance == 0x6E3)
				{
					m_rgbUid2 = new MD4DigestStruct();
					m_rgbUid2.BlobRead(pBlobView);
					nSize -= 16;
				}
				m_tag = pBlobView.UnpackUint8();
				nSize--;
				m_OwnedBLIPFileData = new Blob(true);
				m_OwnedBLIPFileData.GetBlobView().Pack(pBlobView, (int)(nSize));
			}

			public Blob GetBlob()
			{
				if (m_BLIPFileData_ != null)
					return m_BLIPFileData_;
				else
					return m_OwnedBLIPFileData;
			}

			~OfficeArtBlipRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtBStoreContainerRecord : OfficeArtRecord
		{
			public OfficeArtBStoreContainerRecord(OfficeArtRecordHeaderStruct pHeader, BlobView pBlobView) : base(pHeader, IS_CONTAINER, pBlobView)
			{
				nbAssert.Assert((OfficeArtRecord.Type)(pHeader.m_recType) == OfficeArtRecord.Type.TYPE_OFFICE_ART_B_STORE_CONTAINER);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const bool IS_CONTAINER = true;
			protected const OfficeArtRecord.Type TYPE = OfficeArtRecord.Type.TYPE_OFFICE_ART_B_STORE_CONTAINER;
			protected void SetDefaults()
			{
			}

			public OfficeArtBStoreContainerRecord() : base(TYPE, IS_CONTAINER, SIZE, true)
			{
				SetDefaults();
				m_pHeader.m_recVer = 0xF;
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				m_pHeader.m_recInstance = (ushort)(m_pOfficeArtRecordVector.GetSize());
			}

			public override void AddOfficeArtRecord(OfficeArtRecord pOfficeArtRecord)
			{
				nbAssert.Assert(pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_FBSE);
				base.AddOfficeArtRecord(pOfficeArtRecord);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgStrRecord : ParsedExpressionRecord
		{
			public PtgStrRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgStr);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 1;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgStr;
			protected void SetDefaults()
			{
				m_ptg = 0x17;
				m_reserved0 = 0;
				PostSetDefaults();
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected ShortXLUnicodeStringStruct m_string;
			public PtgStrRecord(string szString) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_string.m_rgb.Set(szString);
				m_nSize = (ushort)(m_nSize + ShortXLUnicodeStringStruct.SIZE + m_string.GetDynamicSize());
			}

			protected void PostSetDefaults()
			{
				m_string = new ShortXLUnicodeStringStruct();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				m_string.BlobRead(pBlobView);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_string.BlobWrite(pBlobView);
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new StringToken(m_string.m_rgb.GetExternalString());
			}

			~PtgStrRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgRefRecord : ParsedExpressionRecord
		{
			public PtgRefRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgRef);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved = (byte)((nBitmask0 >> 7) & 0x1);
				m_loc.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				m_loc.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 5;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgRef;
			protected void SetDefaults()
			{
				m_ptg = 0x04;
				m_type = 0x2;
				m_reserved = 0;
				m_loc = new RgceLocStruct();
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved;
			protected RgceLocStruct m_loc;
			public PtgRefRecord(Coordinate pCoordinate) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_loc.m_column.m_col = pCoordinate.m_nX;
				if (pCoordinate.m_bXRelative)
					m_loc.m_column.m_colRelative = 0x1;
				m_loc.m_row.m_rw = pCoordinate.m_nY;
				if (pCoordinate.m_bYRelative)
					m_loc.m_column.m_rowRelative = 0x1;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				Coordinate pCoordinate = new Coordinate(m_loc.m_column.m_col, m_loc.m_row.m_rw, m_loc.m_column.m_colRelative == 0x1, m_loc.m_column.m_rowRelative == 0x1);
				{
					NumberDuck.Secret.Coordinate __3642692973 = pCoordinate;
					pCoordinate = null;
					{
						return new CoordinateToken(__3642692973);
					}
				}
			}

			~PtgRefRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgRef3dRecord : ParsedExpressionRecord
		{
			public PtgRef3dRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgRef3d);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved = (byte)((nBitmask0 >> 7) & 0x1);
				m_ixti = pBlobView.UnpackUint16();
				m_loc.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint16(m_ixti);
				m_loc.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 7;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgRef3d;
			protected void SetDefaults()
			{
				m_ptg = 0x1A;
				m_type = 0x2;
				m_reserved = 0;
				m_ixti = 0;
				m_loc = new RgceLocStruct();
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved;
			protected ushort m_ixti;
			protected RgceLocStruct m_loc;
			public PtgRef3dRecord(Coordinate3d pCoordinate3d, WorkbookGlobals pWorkbookGlobals) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_type = 0x2;
				if (pCoordinate3d.m_nWorksheetFirst != pCoordinate3d.m_nWorksheetLast)
					m_type = 0x1;
				m_loc.m_column.m_col = pCoordinate3d.m_pCoordinate.m_nX;
				if (pCoordinate3d.m_pCoordinate.m_bXRelative)
					m_loc.m_column.m_colRelative = 0x1;
				m_loc.m_row.m_rw = pCoordinate3d.m_pCoordinate.m_nY;
				if (pCoordinate3d.m_pCoordinate.m_bYRelative)
					m_loc.m_column.m_rowRelative = 0x1;
				m_ixti = pWorkbookGlobals.GetWorksheetRangeIndex(pCoordinate3d.m_nWorksheetFirst, pCoordinate3d.m_nWorksheetLast);
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				WorksheetRange pWorksheetRange = pWorkbookGlobals.GetWorksheetRangeByIndex(m_ixti);
				Coordinate pCoordinate = new Coordinate(m_loc.m_column.m_col, m_loc.m_row.m_rw, m_loc.m_column.m_colRelative == 0x1, m_loc.m_column.m_rowRelative == 0x1);
				Coordinate3d pCoordinate3d;
				{
					NumberDuck.Secret.Coordinate __3642692973 = pCoordinate;
					pCoordinate = null;
					pCoordinate3d = new Coordinate3d(pWorksheetRange.m_nFirst, pWorksheetRange.m_nLast, __3642692973);
				}
				Coordinate3dToken pCoordinate3dToken;
				{
					NumberDuck.Secret.Coordinate3d __1094936853 = pCoordinate3d;
					pCoordinate3d = null;
					pCoordinate3dToken = new Coordinate3dToken(__1094936853);
				}
				{
					NumberDuck.Secret.Coordinate3dToken __3867610451 = pCoordinate3dToken;
					pCoordinate3dToken = null;
					{
						return __3867610451;
					}
				}
			}

			~PtgRef3dRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgParenRecord : ParsedExpressionRecord
		{
			public PtgParenRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgParenRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgParen);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
			}

			protected const ushort SIZE = 1;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgParen;
			protected void SetDefaults()
			{
				m_ptg = 0x15;
				m_reserved0 = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			public PtgParenRecord(byte ptg) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				switch (ptg)
				{
					case 0x03:
					{
						m_eType = Type.TYPE_PtgAdd;
						break;
					}

					case 0x04:
					{
						m_eType = Type.TYPE_PtgSub;
						break;
					}

					case 0x05:
					{
						m_eType = Type.TYPE_PtgMul;
						break;
					}

					case 0x06:
					{
						m_eType = Type.TYPE_PtgDiv;
						break;
					}

					case 0x07:
					{
						m_eType = Type.TYPE_PtgPower;
						break;
					}

					case 0x08:
					{
						m_eType = Type.TYPE_PtgConcat;
						break;
					}

					case 0x09:
					{
						m_eType = Type.TYPE_PtgLt;
						break;
					}

					case 0x0A:
					{
						m_eType = Type.TYPE_PtgLe;
						break;
					}

					case 0x0B:
					{
						m_eType = Type.TYPE_PtgEq;
						break;
					}

					case 0x0C:
					{
						m_eType = Type.TYPE_PtgGe;
						break;
					}

					case 0x0D:
					{
						m_eType = Type.TYPE_PtgGt;
						break;
					}

					case 0x0E:
					{
						m_eType = Type.TYPE_PtgNe;
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
				m_ptg = ptg;
			}

			public void PostBlobRead(BlobView pBlobView)
			{
				switch (m_ptg)
				{
					case 0x03:
					{
						m_eType = Type.TYPE_PtgAdd;
						break;
					}

					case 0x04:
					{
						m_eType = Type.TYPE_PtgSub;
						break;
					}

					case 0x05:
					{
						m_eType = Type.TYPE_PtgMul;
						break;
					}

					case 0x06:
					{
						m_eType = Type.TYPE_PtgDiv;
						break;
					}

					case 0x07:
					{
						m_eType = Type.TYPE_PtgPower;
						break;
					}

					case 0x08:
					{
						m_eType = Type.TYPE_PtgConcat;
						break;
					}

					case 0x09:
					{
						m_eType = Type.TYPE_PtgLt;
						break;
					}

					case 0x0A:
					{
						m_eType = Type.TYPE_PtgLe;
						break;
					}

					case 0x0B:
					{
						m_eType = Type.TYPE_PtgEq;
						break;
					}

					case 0x0C:
					{
						m_eType = Type.TYPE_PtgGe;
						break;
					}

					case 0x0D:
					{
						m_eType = Type.TYPE_PtgGt;
						break;
					}

					case 0x0E:
					{
						m_eType = Type.TYPE_PtgNe;
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new Token(Token.Type.TYPE_PAREN, Token.SubType.SUB_TYPE_FUNCTION, 1);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgOperatorRecord : ParsedExpressionRecord
		{
			public PtgOperatorRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
			}

			protected const ushort SIZE = 1;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_UNKNOWN;
			protected void SetDefaults()
			{
				m_ptg = 0;
				m_reserved0 = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			public PtgOperatorRecord(byte ptg) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				switch (ptg)
				{
					case 0x03:
					{
						m_eType = Type.TYPE_PtgAdd;
						break;
					}

					case 0x04:
					{
						m_eType = Type.TYPE_PtgSub;
						break;
					}

					case 0x05:
					{
						m_eType = Type.TYPE_PtgMul;
						break;
					}

					case 0x06:
					{
						m_eType = Type.TYPE_PtgDiv;
						break;
					}

					case 0x07:
					{
						m_eType = Type.TYPE_PtgPower;
						break;
					}

					case 0x08:
					{
						m_eType = Type.TYPE_PtgConcat;
						break;
					}

					case 0x09:
					{
						m_eType = Type.TYPE_PtgLt;
						break;
					}

					case 0x0A:
					{
						m_eType = Type.TYPE_PtgLe;
						break;
					}

					case 0x0B:
					{
						m_eType = Type.TYPE_PtgEq;
						break;
					}

					case 0x0C:
					{
						m_eType = Type.TYPE_PtgGe;
						break;
					}

					case 0x0D:
					{
						m_eType = Type.TYPE_PtgGt;
						break;
					}

					case 0x0E:
					{
						m_eType = Type.TYPE_PtgNe;
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
				m_ptg = ptg;
			}

			public void PostBlobRead(BlobView pBlobView)
			{
				switch (m_ptg)
				{
					case 0x03:
					{
						m_eType = Type.TYPE_PtgAdd;
						break;
					}

					case 0x04:
					{
						m_eType = Type.TYPE_PtgSub;
						break;
					}

					case 0x05:
					{
						m_eType = Type.TYPE_PtgMul;
						break;
					}

					case 0x06:
					{
						m_eType = Type.TYPE_PtgDiv;
						break;
					}

					case 0x07:
					{
						m_eType = Type.TYPE_PtgPower;
						break;
					}

					case 0x08:
					{
						m_eType = Type.TYPE_PtgConcat;
						break;
					}

					case 0x09:
					{
						m_eType = Type.TYPE_PtgLt;
						break;
					}

					case 0x0A:
					{
						m_eType = Type.TYPE_PtgLe;
						break;
					}

					case 0x0B:
					{
						m_eType = Type.TYPE_PtgEq;
						break;
					}

					case 0x0C:
					{
						m_eType = Type.TYPE_PtgGe;
						break;
					}

					case 0x0D:
					{
						m_eType = Type.TYPE_PtgGt;
						break;
					}

					case 0x0E:
					{
						m_eType = Type.TYPE_PtgNe;
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				switch ((Type)(m_eType))
				{
					case Type.TYPE_PtgAdd:
					{
						return new OperatorToken("+");
					}

					case Type.TYPE_PtgSub:
					{
						return new OperatorToken("-");
					}

					case Type.TYPE_PtgMul:
					{
						return new OperatorToken("*");
					}

					case Type.TYPE_PtgDiv:
					{
						return new OperatorToken("/");
					}

					case Type.TYPE_PtgPower:
					{
						return new OperatorToken("^");
					}

					case Type.TYPE_PtgConcat:
					{
						return new OperatorToken("&");
					}

					case Type.TYPE_PtgLt:
					{
						return new OperatorToken("<");
					}

					case Type.TYPE_PtgLe:
					{
						return new OperatorToken("<=");
					}

					case Type.TYPE_PtgEq:
					{
						return new OperatorToken("=");
					}

					case Type.TYPE_PtgGe:
					{
						return new OperatorToken(">=");
					}

					case Type.TYPE_PtgGt:
					{
						return new OperatorToken(">");
					}

					case Type.TYPE_PtgNe:
					{
						return new OperatorToken("<>");
					}

					default:
					{
						break;
					}

				}
				nbAssert.Assert(false);
				return null;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgNumRecord : ParsedExpressionRecord
		{
			public PtgNumRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgNum);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				m_value = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackDouble(m_value);
			}

			protected const ushort SIZE = 9;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgNum;
			protected void SetDefaults()
			{
				m_ptg = 0x1F;
				m_reserved0 = 0;
				m_value = 0.0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected double m_value;
			public PtgNumRecord(double value) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_value = value;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new NumToken(m_value);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgMissArgRecord : ParsedExpressionRecord
		{
			public PtgMissArgRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgMissArgRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgMissArg);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
			}

			protected const ushort SIZE = 1;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgMissArg;
			protected void SetDefaults()
			{
				m_ptg = 0x16;
				m_reserved0 = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new MissArgToken();
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgIntRecord : ParsedExpressionRecord
		{
			public PtgIntRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgInt);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				m_integer = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint16(m_integer);
			}

			protected const ushort SIZE = 3;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgInt;
			protected void SetDefaults()
			{
				m_ptg = 0x1E;
				m_reserved0 = 0;
				m_integer = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected ushort m_integer;
			public PtgIntRecord(ushort nInt) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_integer = nInt;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new IntToken(m_integer);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgFuncVarRecord : ParsedExpressionRecord
		{
			public PtgFuncVarRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgFuncVar);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				m_cparams = pBlobView.UnpackUint8();
				ushort nBitmask1 = pBlobView.UnpackUint16();
				m_tab = (ushort)((nBitmask1 >> 0) & 0x7fff);
				m_fCeFunc = (ushort)((nBitmask1 >> 15) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint8(m_cparams);
				int nBitmask1 = 0;
				nBitmask1 += m_tab << 0;
				nBitmask1 += m_fCeFunc << 15;
				pBlobView.PackUint16((ushort)(nBitmask1));
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgFuncVar;
			protected void SetDefaults()
			{
				m_ptg = 0x02;
				m_type = 1;
				m_reserved0 = 0;
				m_cparams = 0;
				m_tab = 0;
				m_fCeFunc = 0;
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved0;
			protected byte m_cparams;
			protected ushort m_tab;
			protected ushort m_fCeFunc;
			public PtgFuncVarRecord(ushort iftab, byte cparams) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_tab = iftab;
				m_cparams = cparams;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				switch ((Token.Type)(m_tab))
				{
					case Token.Type.TYPE_FUNC_SUM:
					{
						return new SumToken(m_cparams);
					}

					case Token.Type.TYPE_FUNC_TRUE:
					{
						return new BoolToken(true, true);
					}

					case Token.Type.TYPE_FUNC_FALSE:
					{
						return new BoolToken(false, true);
					}

				}
				return new Token((Token.Type)(m_tab), Token.SubType.SUB_TYPE_FUNCTION, m_cparams);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgFuncRecord : ParsedExpressionRecord
		{
			public PtgFuncRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgFunc);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				m_iftab = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint16(m_iftab);
			}

			protected const ushort SIZE = 3;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgFunc;
			protected void SetDefaults()
			{
				m_ptg = 0x01;
				m_type = 1;
				m_reserved0 = 0;
				m_iftab = 0;
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved0;
			protected ushort m_iftab;
			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				switch ((Token.Type)(m_iftab))
				{
					case Token.Type.TYPE_FUNC_SUM:
					{
						return new SumToken(1);
					}

					case Token.Type.TYPE_FUNC_TRUE:
					{
						return new BoolToken(true, true);
					}

					case Token.Type.TYPE_FUNC_FALSE:
					{
						return new BoolToken(false, true);
					}

					case Token.Type.TYPE_FUNC_PI:
					case Token.Type.TYPE_FUNC_NOW:
					case Token.Type.TYPE_FUNC_TODAY:
					{
						return new Token((Token.Type)(m_iftab), Token.SubType.SUB_TYPE_FUNCTION, 0);
					}

					case Token.Type.TYPE_FUNC_INT:
					case Token.Type.TYPE_FUNC_DAY:
					case Token.Type.TYPE_FUNC_MONTH:
					case Token.Type.TYPE_FUNC_YEAR:
					case Token.Type.TYPE_FUNC_WEEKDAY:
					case Token.Type.TYPE_FUNC_HOUR:
					case Token.Type.TYPE_FUNC_MINUTE:
					case Token.Type.TYPE_FUNC_SECOND:
					case Token.Type.TYPE_FUNC_DATEVALUE:
					case Token.Type.TYPE_FUNC_TIMEVALUE:
					{
						return new Token((Token.Type)(m_iftab), Token.SubType.SUB_TYPE_FUNCTION, 1);
					}

					case Token.Type.TYPE_FUNC_ROUND:
					{
						return new Token((Token.Type)(m_iftab), Token.SubType.SUB_TYPE_FUNCTION, 2);
					}

				}
				return new Token((Token.Type)(m_iftab), Token.SubType.SUB_TYPE_FUNCTION, 3);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgBoolRecord : ParsedExpressionRecord
		{
			public PtgBoolRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgBool);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				m_boolean = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint8(m_boolean);
			}

			protected const ushort SIZE = 2;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgBool;
			protected void SetDefaults()
			{
				m_ptg = 0x1D;
				m_reserved0 = 0;
				m_boolean = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_boolean;
			public PtgBoolRecord(bool bBool) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_boolean = bBool ? (byte)(0x1) : (byte)(0x0);
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new BoolToken(m_boolean == 0x1, false);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSumRecord : ParsedExpressionRecord
		{
			public PtgAttrSumRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgAttrSumRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgAttrSum);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_reserved2 = (byte)((nBitmask1 >> 0) & 0xf);
				m_bitIf = (byte)((nBitmask1 >> 4) & 0x1);
				m_reserved3 = (byte)((nBitmask1 >> 5) & 0x7);
				m_unused = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_reserved2 << 0;
				nBitmask1 += m_bitIf << 4;
				nBitmask1 += m_reserved3 << 5;
				pBlobView.PackUint8((byte)(nBitmask1));
				pBlobView.PackUint16(m_unused);
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgAttrSum;
			protected void SetDefaults()
			{
				m_ptg = 0x19;
				m_reserved0 = 0;
				m_reserved2 = 0;
				m_bitIf = 1;
				m_reserved3 = 0;
				m_unused = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_reserved2;
			protected byte m_bitIf;
			protected byte m_reserved3;
			protected ushort m_unused;
			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new SumToken(1);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSpaceRecord : ParsedExpressionRecord
		{
			public PtgAttrSpaceRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgAttrSpace);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_reserved2 = (byte)((nBitmask1 >> 0) & 0x3f);
				m_bitSpace = (byte)((nBitmask1 >> 6) & 0x1);
				m_reserved3 = (byte)((nBitmask1 >> 7) & 0x1);
				m_type.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_reserved2 << 0;
				nBitmask1 += m_bitSpace << 6;
				nBitmask1 += m_reserved3 << 7;
				pBlobView.PackUint8((byte)(nBitmask1));
				m_type.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgAttrSpace;
			protected void SetDefaults()
			{
				m_ptg = 0x19;
				m_reserved0 = 0;
				m_reserved2 = 0;
				m_bitSpace = 1;
				m_reserved3 = 0;
				m_type = new PtgAttrSpaceTypeStruct();
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_reserved2;
			protected byte m_bitSpace;
			protected byte m_reserved3;
			protected PtgAttrSpaceTypeStruct m_type;
			public PtgAttrSpaceRecord(byte nType, byte nCount) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_type.m_type = nType;
				m_type.m_cch = nCount;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return new SpaceToken((SpaceToken.SpaceType)(m_type.m_type), m_type.m_cch);
			}

			~PtgAttrSpaceRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSemiRecord : ParsedExpressionRecord
		{
			public PtgAttrSemiRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgAttrSemiRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgAttrSemi);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_bitSemi = (byte)((nBitmask1 >> 0) & 0x1);
				m_reserved2 = (byte)((nBitmask1 >> 1) & 0x7f);
				m_unused = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_bitSemi << 0;
				nBitmask1 += m_reserved2 << 1;
				pBlobView.PackUint8((byte)(nBitmask1));
				pBlobView.PackUint16(m_unused);
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgAttrSemi;
			protected void SetDefaults()
			{
				m_ptg = 0x19;
				m_reserved0 = 0;
				m_bitSemi = 1;
				m_reserved2 = 0;
				m_unused = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_bitSemi;
			protected byte m_reserved2;
			protected ushort m_unused;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrIfRecord : ParsedExpressionRecord
		{
			public PtgAttrIfRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgAttrIfRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgAttrIf);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_reserved2 = (byte)((nBitmask1 >> 0) & 0x1);
				m_bitIf = (byte)((nBitmask1 >> 1) & 0x1);
				m_reserved3 = (byte)((nBitmask1 >> 2) & 0x3f);
				m_offset = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_reserved2 << 0;
				nBitmask1 += m_bitIf << 1;
				nBitmask1 += m_reserved3 << 2;
				pBlobView.PackUint8((byte)(nBitmask1));
				pBlobView.PackUint16(m_offset);
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgAttrIf;
			protected void SetDefaults()
			{
				m_ptg = 0x19;
				m_reserved0 = 0;
				m_reserved2 = 0;
				m_bitIf = 1;
				m_reserved3 = 0;
				m_offset = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_reserved2;
			protected byte m_bitIf;
			protected byte m_reserved3;
			protected ushort m_offset;
			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				return null;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrGotoRecord : ParsedExpressionRecord
		{
			public PtgAttrGotoRecord() : base(TYPE, SIZE, true)
			{
				SetDefaults();
			}

			public PtgAttrGotoRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgAttrGoto);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x7f);
				m_reserved0 = (byte)((nBitmask0 >> 7) & 0x1);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_reserved2 = (byte)((nBitmask1 >> 0) & 0x7);
				m_bitGoto = (byte)((nBitmask1 >> 3) & 0x1);
				m_reserved3 = (byte)((nBitmask1 >> 4) & 0xf);
				m_unused = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_reserved0 << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_reserved2 << 0;
				nBitmask1 += m_bitGoto << 3;
				nBitmask1 += m_reserved3 << 4;
				pBlobView.PackUint8((byte)(nBitmask1));
				pBlobView.PackUint16(m_unused);
			}

			protected const ushort SIZE = 4;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgAttrGoto;
			protected void SetDefaults()
			{
				m_ptg = 0x19;
				m_reserved0 = 0;
				m_reserved2 = 0;
				m_bitGoto = 1;
				m_reserved3 = 0;
				m_unused = 0;
			}

			protected byte m_ptg;
			protected byte m_reserved0;
			protected byte m_reserved2;
			protected byte m_bitGoto;
			protected byte m_reserved3;
			protected ushort m_unused;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAreaRecord : ParsedExpressionRecord
		{
			public PtgAreaRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgArea);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved = (byte)((nBitmask0 >> 7) & 0x1);
				m_area.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				m_area.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 9;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgArea;
			protected void SetDefaults()
			{
				m_ptg = 0x05;
				m_type = 0x1;
				m_reserved = 0;
				m_area = new RgceAreaStruct();
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved;
			protected RgceAreaStruct m_area;
			public PtgAreaRecord(Area pArea) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				m_area.m_columnFirst.m_col = pArea.m_pTopLeft.m_nX;
				if (pArea.m_pTopLeft.m_bXRelative)
					m_area.m_columnFirst.m_colRelative = 0x1;
				m_area.m_rowFirst.m_rw = pArea.m_pTopLeft.m_nY;
				if (pArea.m_pTopLeft.m_bYRelative)
					m_area.m_columnFirst.m_rowRelative = 0x1;
				m_area.m_columnLast.m_col = pArea.m_pBottomRight.m_nX;
				if (pArea.m_pTopLeft.m_bXRelative)
					m_area.m_columnLast.m_colRelative = 0x1;
				m_area.m_rowLast.m_rw = pArea.m_pBottomRight.m_nY;
				if (pArea.m_pTopLeft.m_bYRelative)
					m_area.m_columnLast.m_rowRelative = 0x1;
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				Coordinate pTopLeft = new Coordinate(m_area.m_columnFirst.m_col, m_area.m_rowFirst.m_rw, m_area.m_columnFirst.m_colRelative == 0x1, m_area.m_columnFirst.m_rowRelative == 0x1);
				Coordinate pBottomRight = new Coordinate(m_area.m_columnLast.m_col, m_area.m_rowLast.m_rw, m_area.m_columnLast.m_colRelative == 0x1, m_area.m_columnLast.m_rowRelative == 0x1);
				Area pArea;
				{
					NumberDuck.Secret.Coordinate __2706830545 = pTopLeft;
					pTopLeft = null;
					{
						NumberDuck.Secret.Coordinate __2652773404 = pBottomRight;
						pBottomRight = null;
						pArea = new Area(__2706830545, __2652773404);
					}
				}
				{
					NumberDuck.Secret.Area __4245081970 = pArea;
					pArea = null;
					{
						return new AreaToken(__4245081970);
					}
				}
			}

			~PtgAreaRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgArea3dRecord : ParsedExpressionRecord
		{
			public PtgArea3dRecord(BlobView pBlobView) : base(TYPE, SIZE, true)
			{
				nbAssert.Assert((ParsedExpressionRecord.Type)(m_eType) == ParsedExpressionRecord.Type.TYPE_PtgArea3d);
				SetDefaults();
				BlobRead(pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_ptg = (byte)((nBitmask0 >> 0) & 0x1f);
				m_type = (byte)((nBitmask0 >> 5) & 0x3);
				m_reserved = (byte)((nBitmask0 >> 7) & 0x1);
				m_ixti = pBlobView.UnpackUint16();
				m_area.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ptg << 0;
				nBitmask0 += m_type << 5;
				nBitmask0 += m_reserved << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint16(m_ixti);
				m_area.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 11;
			protected const ParsedExpressionRecord.Type TYPE = ParsedExpressionRecord.Type.TYPE_PtgArea3d;
			protected void SetDefaults()
			{
				m_ptg = 0x1B;
				m_type = 0x1;
				m_reserved = 0;
				m_ixti = 0;
				m_area = new RgceAreaStruct();
			}

			protected byte m_ptg;
			protected byte m_type;
			protected byte m_reserved;
			protected ushort m_ixti;
			protected RgceAreaStruct m_area;
			public PtgArea3dRecord(Area3d pArea3d, WorkbookGlobals pWorkbookGlobals) : base(TYPE, SIZE, true)
			{
				SetDefaults();
				Area pArea = pArea3d.m_pArea;
				m_area.m_columnFirst.m_col = pArea.m_pTopLeft.m_nX;
				if (pArea.m_pTopLeft.m_bXRelative)
					m_area.m_columnFirst.m_colRelative = 0x1;
				m_area.m_rowFirst.m_rw = pArea.m_pTopLeft.m_nY;
				if (pArea.m_pTopLeft.m_bYRelative)
					m_area.m_columnFirst.m_rowRelative = 0x1;
				m_area.m_columnLast.m_col = pArea.m_pBottomRight.m_nX;
				if (pArea.m_pTopLeft.m_bXRelative)
					m_area.m_columnLast.m_colRelative = 0x1;
				m_area.m_rowLast.m_rw = pArea.m_pBottomRight.m_nY;
				if (pArea.m_pTopLeft.m_bYRelative)
					m_area.m_columnLast.m_rowRelative = 0x1;
				m_ixti = pWorkbookGlobals.GetWorksheetRangeIndex(pArea3d.m_pWorksheetRange.m_nFirst, pArea3d.m_pWorksheetRange.m_nLast);
			}

			public override Token GetToken(WorkbookGlobals pWorkbookGlobals)
			{
				WorksheetRange pWorksheetRange = pWorkbookGlobals.GetWorksheetRangeByIndex(m_ixti);
				Coordinate pTopLeft = new Coordinate(m_area.m_columnFirst.m_col, m_area.m_rowFirst.m_rw, m_area.m_columnFirst.m_colRelative == 0x1, m_area.m_columnFirst.m_rowRelative == 0x1);
				Coordinate pBottomRight = new Coordinate(m_area.m_columnLast.m_col, m_area.m_rowLast.m_rw, m_area.m_columnLast.m_colRelative == 0x1, m_area.m_columnLast.m_rowRelative == 0x1);
				Area pArea;
				{
					NumberDuck.Secret.Coordinate __2706830545 = pTopLeft;
					pTopLeft = null;
					{
						NumberDuck.Secret.Coordinate __2652773404 = pBottomRight;
						pBottomRight = null;
						pArea = new Area(__2706830545, __2652773404);
					}
				}
				Area3d pArea3d;
				{
					NumberDuck.Secret.Area __4245081970 = pArea;
					pArea = null;
					pArea3d = new Area3d(pWorksheetRange.m_nFirst, pWorksheetRange.m_nLast, __4245081970);
				}
				{
					NumberDuck.Secret.Area3d __2738670685 = pArea3d;
					pArea3d = null;
					{
						return new Area3dToken(__2738670685);
					}
				}
			}

			~PtgArea3dRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Hsl
		{
			public double m_fH;
			public double m_fS;
			public double m_fL;
			public void ToColor(Color pColor)
			{
				double fR = 0.0;
				double fG = 0.0;
				double fB = 0.0;
				if (m_fS == 0.0)
				{
					fR = m_fL;
					fG = m_fL;
					fB = m_fL;
				}
				else
				{
					double q = m_fL < 0.5 ? m_fL * (1.0 + m_fS) : m_fL + m_fS - m_fL * m_fS;
					double p = 2 * m_fL - q;
					fR = hue2rgb(p, q, m_fH + 1 / 3.0);
					fG = hue2rgb(p, q, m_fH);
					fB = hue2rgb(p, q, m_fH - 1 / 3.0);
				}
				pColor.Set((byte)(fR * 255.0 + 0.5), (byte)(fG * 255.0 + 0.5), (byte)(fB * 255.0 + 0.5));
			}

			public void Set(Color pColor)
			{
				double fR = pColor.GetRed() / 255.0;
				double fG = pColor.GetGreen() / 255.0;
				double fB = pColor.GetBlue() / 255.0;
				double fMin = min(fR, min(fG, fB));
				double fMax = max(fR, max(fG, fB));
				m_fH = (fMin + fMax) / 2.0;
				m_fS = m_fH;
				m_fL = m_fH;
				if (fMin == fMax)
				{
					m_fH = 0.0;
					m_fS = 0.0;
				}
				else
				{
					double fDiff = fMax - fMin;
					m_fS = m_fL > 0.5 ? fDiff / (2.0 - fMax - fMin) : fDiff / (fMax + fMin);
					if (fMax == fR)
						m_fH = (fG - fB) / fDiff + (fG < fB ? 6 : 0);
					else if (fMax == fG)
						m_fH = (fB - fR) / fDiff + 2;
					else
						m_fH = (fR - fG) / fDiff + 4;
					m_fH = m_fH / 6.0;
				}
			}

			protected double max(double a, double b)
			{
				return a > b ? a : b;
			}

			protected double min(double a, double b)
			{
				return a < b ? a : b;
			}

			protected double hue2rgb(double p, double q, double t)
			{
				if (t < 0.0)
					t += 1.0;
				if (t > 1.0)
					t -= 1.0;
				if (t < 1.0 / 6.0)
					return p + (q - p) * 6.0 * t;
				if (t < 1.0 / 2.0)
					return q;
				if (t < 2.0 / 3.0)
					return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
				return p;
			}

		}
		class XFExt : BiffRecord
		{
			public XFExt(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_XF_EXT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_reserved1 = pBlobView.UnpackUint16();
				m_ixfe = pBlobView.UnpackUint16();
				m_reserved2 = pBlobView.UnpackUint16();
				m_cexts = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_reserved1);
				pBlobView.PackUint16(m_ixfe);
				pBlobView.PackUint16(m_reserved2);
				pBlobView.PackUint16(m_cexts);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_XF_EXT;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_reserved1 = 0;
				m_ixfe = 0;
				m_reserved2 = 0;
				m_cexts = 0;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected ushort m_reserved1;
			protected ushort m_ixfe;
			protected ushort m_reserved2;
			protected ushort m_cexts;
			protected Color m_pTempColor;
			protected OwnedVector<ExtPropStruct> m_pExtPropVector;
			public XFExt(ushort nXFIndex, Color pForegroundColor, Color pBackgroundColor, Color pTextColor) : base(TYPE, SIZE + 3 * (ExtPropStruct.SIZE + FullColorExtStruct.SIZE))
			{
				SetDefaults();
				m_frtHeader.m_rt = (ushort)(BiffRecord.Type.TYPE_XF_EXT);
				m_ixfe = nXFIndex;
				m_cexts = 3;
				m_pExtPropVector = new OwnedVector<ExtPropStruct>();
				{
					ExtPropStruct pExtProp = new ExtPropStruct();
					pExtProp.m_extType = 0x000D;
					pExtProp.m_pFullColorExt = new FullColorExtStruct();
					pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRAUTO);
					if (pTextColor != null)
					{
						pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRRGB);
						pExtProp.m_pFullColorExt.m_xclrValue = pTextColor.GetRgba();
					}
					pExtProp.m_pFullColorExt.m_unusedA = 95;
					pExtProp.m_pFullColorExt.m_unusedB = 45;
					pExtProp.m_pFullColorExt.m_unusedC = 59;
					pExtProp.m_pFullColorExt.m_unusedD = 95;
					pExtProp.m_pFullColorExt.m_unusedE = 45;
					pExtProp.m_pFullColorExt.m_unusedF = 42;
					pExtProp.m_pFullColorExt.m_unusedG = 32;
					pExtProp.m_pFullColorExt.m_unusedH = 34;
					{
						NumberDuck.Secret.ExtPropStruct __1075823487 = pExtProp;
						pExtProp = null;
						m_pExtPropVector.PushBack(__1075823487);
					}
				}
				{
					ExtPropStruct pExtProp = new ExtPropStruct();
					pExtProp.m_extType = 0x0004;
					pExtProp.m_pFullColorExt = new FullColorExtStruct();
					pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRAUTO);
					if (pForegroundColor != null)
					{
						pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRRGB);
						pExtProp.m_pFullColorExt.m_xclrValue = pForegroundColor.GetRgba();
					}
					pExtProp.m_pFullColorExt.m_unusedA = 45;
					pExtProp.m_pFullColorExt.m_unusedB = 64;
					pExtProp.m_pFullColorExt.m_unusedC = 95;
					pExtProp.m_pFullColorExt.m_unusedD = 8;
					pExtProp.m_pFullColorExt.m_unusedE = 0;
					pExtProp.m_pFullColorExt.m_unusedF = 20;
					pExtProp.m_pFullColorExt.m_unusedG = 0;
					pExtProp.m_pFullColorExt.m_unusedH = 2;
					{
						NumberDuck.Secret.ExtPropStruct __1075823487 = pExtProp;
						pExtProp = null;
						m_pExtPropVector.PushBack(__1075823487);
					}
				}
				{
					ExtPropStruct pExtProp = new ExtPropStruct();
					pExtProp.m_extType = 0x0005;
					pExtProp.m_pFullColorExt = new FullColorExtStruct();
					pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRAUTO);
					if (pBackgroundColor != null)
					{
						pExtProp.m_pFullColorExt.m_xclrType = (ushort)(BiffStruct.XColorType.XCLRRGB);
						pExtProp.m_pFullColorExt.m_xclrValue = pBackgroundColor.GetRgba();
					}
					pExtProp.m_pFullColorExt.m_unusedA = 32;
					pExtProp.m_pFullColorExt.m_unusedB = 32;
					pExtProp.m_pFullColorExt.m_unusedC = 32;
					pExtProp.m_pFullColorExt.m_unusedD = 32;
					pExtProp.m_pFullColorExt.m_unusedE = 32;
					pExtProp.m_pFullColorExt.m_unusedF = 32;
					pExtProp.m_pFullColorExt.m_unusedG = 32;
					pExtProp.m_pFullColorExt.m_unusedH = 32;
					{
						NumberDuck.Secret.ExtPropStruct __1075823487 = pExtProp;
						pExtProp = null;
						m_pExtPropVector.PushBack(__1075823487);
					}
				}
			}

			protected void PostSetDefaults()
			{
				m_pTempColor = new Color(0, 0, 0);
				m_pExtPropVector = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				ushort i;
				nbAssert.Assert(m_pExtPropVector == null);
				m_pExtPropVector = new OwnedVector<ExtPropStruct>();
				for (i = 0; i < m_cexts; i++)
				{
					ExtPropStruct pExtProp = new ExtPropStruct();
					pExtProp.BlobRead(pBlobView);
					{
						NumberDuck.Secret.ExtPropStruct __1075823487 = pExtProp;
						pExtProp = null;
						m_pExtPropVector.PushBack(__1075823487);
					}
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(m_cexts == m_pExtPropVector.GetSize());
				int i;
				for (i = 0; i < m_pExtPropVector.GetSize(); i++)
				{
					ExtPropStruct pExtProp = m_pExtPropVector.Get(i);
					pExtProp.BlobWrite(pBlobView);
				}
			}

			public ushort GetXFIndex()
			{
				return m_ixfe;
			}

			public Color GetForegroundColor(Theme pTheme)
			{
				return GetColorByType(0x0004, pTheme);
			}

			public Color GetTextColor(Theme pTheme)
			{
				return GetColorByType(0x000D, pTheme);
			}

			protected ExtPropStruct GetExtPropByType(ushort nType)
			{
				int i;
				for (i = 0; i < m_pExtPropVector.GetSize(); i++)
				{
					ExtPropStruct pExtProp = m_pExtPropVector.Get(i);
					if (pExtProp.m_extType == nType)
						return pExtProp;
				}
				return null;
			}

			protected Color GetColorByType(ushort nType, Theme pTheme)
			{
				ExtPropStruct pExtProp = GetExtPropByType(nType);
				if (pExtProp != null)
				{
					switch ((BiffStruct.XColorType)(pExtProp.m_pFullColorExt.m_xclrType))
					{
						case BiffStruct.XColorType.XCLRINDEXED:
						{
							break;
						}

						case BiffStruct.XColorType.XCLRRGB:
						{
							uint nColor = pExtProp.m_pFullColorExt.m_xclrValue;
							m_pTempColor.Set((byte)(nColor & 0xFF), (byte)((nColor >> 8) & 0xFF), (byte)((nColor >> 16) & 0xFF));
							return m_pTempColor;
						}

						case BiffStruct.XColorType.XCLRTHEMED:
						{
							nbAssert.Assert(pTheme != null);
							uint nColor = pTheme.GetColorByIndex((int)(pExtProp.m_pFullColorExt.m_xclrValue));
							uint nR = (nColor >> 16) & 0xFF;
							uint nG = (nColor >> 8) & 0xFF;
							uint nB = nColor & 0xFF;
							{
								m_pTempColor.Set((byte)(nR), (byte)(nG), (byte)(nB));
								Hsl pHsl = new Hsl();
								pHsl.Set(m_pTempColor);
								double fTint = pExtProp.m_pFullColorExt.m_nTintShade / 32767.0;
								if (fTint < 0)
									pHsl.m_fL = pHsl.m_fL * (1.0 + fTint);
								if (fTint > 0)
									pHsl.m_fL = pHsl.m_fL * (1.0 - fTint) + (1.0 - 1.0 * (1.0 - fTint));
								pHsl.ToColor(m_pTempColor);
							}
							return m_pTempColor;
						}

					}
				}
				return null;
			}

			~XFExt()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XFCRC : BiffRecord
		{
			public XFCRC(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_XF_CRC);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_reserved = pBlobView.UnpackUint16();
				m_cxfs = pBlobView.UnpackUint16();
				m_crc = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_reserved);
				pBlobView.PackUint16(m_cxfs);
				pBlobView.PackUint32(m_crc);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_XF_CRC;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_reserved = 0;
				m_cxfs = 0;
				m_crc = 0;
			}

			protected FrtHeaderStruct m_frtHeader;
			protected ushort m_reserved;
			protected ushort m_cxfs;
			protected uint m_crc;
			public XFCRC(OwnedVector<XF> pXFVector) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_frtHeader.m_rt = (ushort)(BiffRecord.Type.TYPE_XF_CRC);
				m_reserved = 0;
				m_cxfs = (ushort)(pXFVector.GetSize());
				m_crc = ComputateCrc(pXFVector);
			}

			public bool ValidateCrc(Vector<XF> pXFVector)
			{
				if (m_cxfs == pXFVector.GetSize())
				{
					return ComputateCrc(pXFVector) == m_crc;
				}
				return false;
			}

			protected uint ComputateCrc(OwnedVector<XF> pXFVector)
			{
				Blob pBlob = new Blob(true);
				BlobView pBlobView = pBlob.GetBlobView();
				uint nCrcValue = 0;
				for (int i = 0; i < pXFVector.GetSize(); i++)
					pXFVector.Get(i).BlobWrite(pBlobView);
				pBlobView.SetOffset(0);
				while (pBlobView.GetOffset() < pBlobView.GetSize())
				{
					uint nTemp = pBlobView.UnpackUint8();
					nCrcValue = nCrcValue ^ (nTemp << 24);
					for (int k = 0; k < 8; k++)
					{
						if ((nCrcValue & 0x80000000) > 0)
							nCrcValue = (nCrcValue << 1) ^ 0xAF;
						else
							nCrcValue = nCrcValue << 1;
					}
				}
				{
					return nCrcValue;
				}
			}

			protected uint ComputateCrc(Vector<XF> pXFVector)
			{
				Blob pBlob = new Blob(true);
				BlobView pBlobView = pBlob.GetBlobView();
				uint nCrcValue = 0;
				for (int i = 0; i < pXFVector.GetSize(); i++)
					pXFVector.Get(i).BlobWrite(pBlobView);
				pBlobView.SetOffset(0);
				while (pBlobView.GetOffset() < pBlobView.GetSize())
				{
					uint nTemp = pBlobView.UnpackUint8();
					nCrcValue = nCrcValue ^ (nTemp << 24);
					for (int k = 0; k < 8; k++)
					{
						if ((nCrcValue & 0x80000000) > 0)
							nCrcValue = (nCrcValue << 1) ^ 0xAF;
						else
							nCrcValue = nCrcValue << 1;
					}
				}
				{
					return nCrcValue;
				}
			}

			~XFCRC()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XF : BiffRecord
		{
			public XF(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_XF);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ifnt = pBlobView.UnpackUint16();
				m_ifmt = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fLocked = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fHidden = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fStyle = (ushort)((nBitmask0 >> 2) & 0x1);
				m_f123Prefix = (ushort)((nBitmask0 >> 3) & 0x1);
				m_ixfParent = (ushort)((nBitmask0 >> 4) & 0xfff);
				byte nBitmask1 = pBlobView.UnpackUint8();
				m_alc = (byte)((nBitmask1 >> 0) & 0x7);
				m_fWrap = (byte)((nBitmask1 >> 3) & 0x1);
				m_alcV = (byte)((nBitmask1 >> 4) & 0x7);
				m_fJustLast = (byte)((nBitmask1 >> 7) & 0x1);
				m_trot = pBlobView.UnpackUint8();
				byte nBitmask2 = pBlobView.UnpackUint8();
				m_cIndent = (byte)((nBitmask2 >> 0) & 0xf);
				m_fShrinkToFit = (byte)((nBitmask2 >> 4) & 0x1);
				m_reserved1 = (byte)((nBitmask2 >> 5) & 0x1);
				m_iReadOrder = (byte)((nBitmask2 >> 6) & 0x3);
				byte nBitmask3 = pBlobView.UnpackUint8();
				m_reserved2 = (byte)((nBitmask3 >> 0) & 0x3);
				m_fAtrNum = (byte)((nBitmask3 >> 2) & 0x1);
				m_fAtrFnt = (byte)((nBitmask3 >> 3) & 0x1);
				m_fAtrAlc = (byte)((nBitmask3 >> 4) & 0x1);
				m_fAtrBdr = (byte)((nBitmask3 >> 5) & 0x1);
				m_fAtrPat = (byte)((nBitmask3 >> 6) & 0x1);
				m_fAtrProt = (byte)((nBitmask3 >> 7) & 0x1);
				uint nBitmask4 = pBlobView.UnpackUint32();
				m_dgLeft = (uint)((nBitmask4 >> 0) & 0xf);
				m_dgRight = (uint)((nBitmask4 >> 4) & 0xf);
				m_dgTop = (uint)((nBitmask4 >> 8) & 0xf);
				m_dgBottom = (uint)((nBitmask4 >> 12) & 0xf);
				m_icvLeft = (uint)((nBitmask4 >> 16) & 0x7f);
				m_icvRight = (uint)((nBitmask4 >> 23) & 0x7f);
				m_grbitDiag = (uint)((nBitmask4 >> 30) & 0x3);
				uint nBitmask5 = pBlobView.UnpackUint32();
				m_icvTop = (uint)((nBitmask5 >> 0) & 0x7f);
				m_icvBottom = (uint)((nBitmask5 >> 7) & 0x7f);
				m_icvDiag = (uint)((nBitmask5 >> 14) & 0x7f);
				m_dgDiag = (uint)((nBitmask5 >> 21) & 0xf);
				m_fHasXFExt = (uint)((nBitmask5 >> 25) & 0x1);
				m_fls = (uint)((nBitmask5 >> 26) & 0x3f);
				ushort nBitmask6 = pBlobView.UnpackUint16();
				m_icvFore = (ushort)((nBitmask6 >> 0) & 0x7f);
				m_icvBack = (ushort)((nBitmask6 >> 7) & 0x7f);
				m_fsxButton = (ushort)((nBitmask6 >> 14) & 0x1);
				m_reserved3 = (ushort)((nBitmask6 >> 15) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ifnt);
				pBlobView.PackUint16(m_ifmt);
				int nBitmask0 = 0;
				nBitmask0 += m_fLocked << 0;
				nBitmask0 += m_fHidden << 1;
				nBitmask0 += m_fStyle << 2;
				nBitmask0 += m_f123Prefix << 3;
				nBitmask0 += m_ixfParent << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
				int nBitmask1 = 0;
				nBitmask1 += m_alc << 0;
				nBitmask1 += m_fWrap << 3;
				nBitmask1 += m_alcV << 4;
				nBitmask1 += m_fJustLast << 7;
				pBlobView.PackUint8((byte)(nBitmask1));
				pBlobView.PackUint8(m_trot);
				int nBitmask2 = 0;
				nBitmask2 += m_cIndent << 0;
				nBitmask2 += m_fShrinkToFit << 4;
				nBitmask2 += m_reserved1 << 5;
				nBitmask2 += m_iReadOrder << 6;
				pBlobView.PackUint8((byte)(nBitmask2));
				int nBitmask3 = 0;
				nBitmask3 += m_reserved2 << 0;
				nBitmask3 += m_fAtrNum << 2;
				nBitmask3 += m_fAtrFnt << 3;
				nBitmask3 += m_fAtrAlc << 4;
				nBitmask3 += m_fAtrBdr << 5;
				nBitmask3 += m_fAtrPat << 6;
				nBitmask3 += m_fAtrProt << 7;
				pBlobView.PackUint8((byte)(nBitmask3));
				uint nBitmask4 = 0;
				nBitmask4 += m_dgLeft << 0;
				nBitmask4 += m_dgRight << 4;
				nBitmask4 += m_dgTop << 8;
				nBitmask4 += m_dgBottom << 12;
				nBitmask4 += m_icvLeft << 16;
				nBitmask4 += m_icvRight << 23;
				nBitmask4 += m_grbitDiag << 30;
				pBlobView.PackUint32((uint)(nBitmask4));
				uint nBitmask5 = 0;
				nBitmask5 += m_icvTop << 0;
				nBitmask5 += m_icvBottom << 7;
				nBitmask5 += m_icvDiag << 14;
				nBitmask5 += m_dgDiag << 21;
				nBitmask5 += m_fHasXFExt << 25;
				nBitmask5 += m_fls << 26;
				pBlobView.PackUint32((uint)(nBitmask5));
				int nBitmask6 = 0;
				nBitmask6 += m_icvFore << 0;
				nBitmask6 += m_icvBack << 7;
				nBitmask6 += m_fsxButton << 14;
				nBitmask6 += m_reserved3 << 15;
				pBlobView.PackUint16((ushort)(nBitmask6));
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_XF;
			protected void SetDefaults()
			{
				m_ifnt = 0;
				m_ifmt = 0;
				m_fLocked = 0;
				m_fHidden = 0;
				m_fStyle = 0;
				m_f123Prefix = 0;
				m_ixfParent = 0;
				m_alc = 0;
				m_fWrap = 0;
				m_alcV = 0;
				m_fJustLast = 0;
				m_trot = 0;
				m_cIndent = 0;
				m_fShrinkToFit = 0;
				m_reserved1 = 0;
				m_iReadOrder = 0;
				m_reserved2 = 0;
				m_fAtrNum = 0;
				m_fAtrFnt = 0;
				m_fAtrAlc = 0;
				m_fAtrBdr = 0;
				m_fAtrPat = 0;
				m_fAtrProt = 0;
				m_dgLeft = 0;
				m_dgRight = 0;
				m_dgTop = 0;
				m_dgBottom = 0;
				m_icvLeft = 0;
				m_icvRight = 0;
				m_grbitDiag = 0;
				m_icvTop = 0;
				m_icvBottom = 0;
				m_icvDiag = 0;
				m_dgDiag = 0;
				m_fHasXFExt = 0;
				m_fls = 0;
				m_icvFore = 0;
				m_icvBack = 0;
				m_fsxButton = 0;
				m_reserved3 = 0;
			}

			protected ushort m_ifnt;
			protected ushort m_ifmt;
			protected ushort m_fLocked;
			protected ushort m_fHidden;
			protected ushort m_fStyle;
			protected ushort m_f123Prefix;
			protected ushort m_ixfParent;
			protected byte m_alc;
			protected byte m_fWrap;
			protected byte m_alcV;
			protected byte m_fJustLast;
			protected byte m_trot;
			protected byte m_cIndent;
			protected byte m_fShrinkToFit;
			protected byte m_reserved1;
			protected byte m_iReadOrder;
			protected byte m_reserved2;
			protected byte m_fAtrNum;
			protected byte m_fAtrFnt;
			protected byte m_fAtrAlc;
			protected byte m_fAtrBdr;
			protected byte m_fAtrPat;
			protected byte m_fAtrProt;
			protected uint m_dgLeft;
			protected uint m_dgRight;
			protected uint m_dgTop;
			protected uint m_dgBottom;
			protected uint m_icvLeft;
			protected uint m_icvRight;
			protected uint m_grbitDiag;
			protected uint m_icvTop;
			protected uint m_icvBottom;
			protected uint m_icvDiag;
			protected uint m_dgDiag;
			protected uint m_fHasXFExt;
			protected uint m_fls;
			protected ushort m_icvFore;
			protected ushort m_icvBack;
			protected ushort m_fsxButton;
			protected ushort m_reserved3;
			public XF(ushort ifnt, ushort ifmt, byte fStyle, ushort ixfParent, byte fAtrNum, byte fAtrFnt, byte fAtrAlc, byte fAtrBdr, byte fAtrPat, byte fAtrProt, byte dgLeft, byte dgRight, byte dgTop, byte dgBottom, byte icvLeft, byte icvRight, byte icvTop, byte icvBottom, byte fHasXFExt, byte fls, byte icvFore, byte icvBack) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_ifnt = ifnt;
				m_ifmt = ifmt;
				m_fLocked = 0x1;
				m_fStyle = (ushort)(fStyle & 0x1);
				m_ixfParent = (ushort)(ixfParent & 0xFFF);
				m_alcV = 2;
				m_fJustLast = 0;
				m_fAtrNum = (byte)(fAtrNum & 0x1);
				m_fAtrFnt = (byte)(fAtrFnt & 0x1);
				m_fAtrAlc = (byte)(fAtrAlc & 0x1);
				m_fAtrBdr = (byte)(fAtrBdr & 0x1);
				m_fAtrPat = (byte)(fAtrPat & 0x1);
				m_fAtrProt = (byte)(fAtrProt & 0x1);
				m_dgLeft = (uint)(dgLeft & 0xF);
				m_dgRight = (uint)(dgRight & 0xF);
				m_dgTop = (uint)(dgTop & 0xF);
				m_dgBottom = (uint)(dgBottom & 0xF);
				m_icvLeft = (uint)(icvLeft & 0x7F);
				m_icvRight = (uint)(icvRight & 0x7F);
				m_icvTop = (uint)(icvTop & 0x7F);
				m_icvBottom = (uint)(icvBottom & 0x7F);
				m_fHasXFExt = (uint)(fHasXFExt & 0x1);
				m_fls = (uint)(fls & 0x3F);
				m_icvFore = (ushort)(icvFore & 0x7F);
				m_icvBack = (ushort)(icvBack & 0x7F);
			}

			public XF(ushort nFontIndex, ushort nFormatIndex, Style pStyle, bool bHasXFExt) : base(TYPE, SIZE)
			{
				SetDefaults();
				if (nFontIndex >= 4)
					nFontIndex++;
				m_ifnt = nFontIndex;
				m_ifmt = nFormatIndex;
				byte nBackgroundColourIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetBackgroundColor(false));
				if (nBackgroundColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT)
					nBackgroundColourIndex = BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_FOREGROUND;
				byte nFillPatternColourIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetFillPatternColor(false));
				if (nFillPatternColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT)
					nFillPatternColourIndex = BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_BACKGROUND;
				byte nTopBorderType = LineTypeToBorderStyle(pStyle.GetTopBorderLine().GetType());
				ushort nTopBorderPaletteIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetTopBorderLine().GetColor());
				byte nRightBorderType = LineTypeToBorderStyle(pStyle.GetRightBorderLine().GetType());
				ushort nRightBorderPaletteIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetRightBorderLine().GetColor());
				byte nBottomBorderType = LineTypeToBorderStyle(pStyle.GetBottomBorderLine().GetType());
				ushort nBottomBorderPaletteIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetBottomBorderLine().GetColor());
				byte nLeftBorderType = LineTypeToBorderStyle(pStyle.GetLeftBorderLine().GetType());
				ushort nLeftBorderPaletteIndex = BiffWorkbookGlobals.SnapToPalette(pStyle.GetLeftBorderLine().GetColor());
				byte nHorizontalAlign = (byte)(pStyle.GetHorizontalAlign());
				byte nVerticalAlign = (byte)(pStyle.GetVerticalAlign());
				m_fLocked = 0x1;
				ushort fls = 0x00;
				switch (pStyle.GetFillPattern())
				{
					case Style.FillPattern.FILL_PATTERN_NONE:
					{
						if (nBackgroundColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_FOREGROUND)
							fls = 0x00;
						else
							fls = 0x01;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_50:
					{
						fls = 0x02;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_75:
					{
						fls = 0x03;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_25:
					{
						fls = 0x04;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_HORIZONTAL_STRIPE:
					{
						fls = 0x05;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_VARTICAL_STRIPE:
					{
						fls = 0x06;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_REVERSE_DIAGONAL_STRIPE:
					{
						fls = 0x07;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_DIAGONAL_STRIPE:
					{
						fls = 0x08;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_DIAGONAL_CROSSHATCH:
					{
						fls = 0x09;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THICK_DIAGONAL_CROSSHATCH:
					{
						fls = 0x0A;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_HORIZONTAL_STRIPE:
					{
						fls = 0x0B;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_VERTICAL_STRIPE:
					{
						fls = 0x0C;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_REVERSE_VERTICAL_STRIPE:
					{
						fls = 0x0D;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_DIAGONAL_STRIPE:
					{
						fls = 0x0E;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_HORIZONTAL_CROSSHATCH:
					{
						fls = 0x0F;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_THIN_DIAGONAL_CROSSHATCH:
					{
						fls = 0x10;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_125:
					{
						fls = 0x11;
						break;
					}

					case Style.FillPattern.FILL_PATTERN_625:
					{
						fls = 0x12;
						break;
					}

				}
				nbAssert.Assert(nBackgroundColourIndex >= BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY && nBackgroundColourIndex < BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY || nBackgroundColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_FOREGROUND);
				nbAssert.Assert(nFillPatternColourIndex >= BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY && nFillPatternColourIndex < BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY || nFillPatternColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_BACKGROUND);
				{
					m_fls = fls;
					m_icvFore = nBackgroundColourIndex;
					m_icvBack = nFillPatternColourIndex;
				}
				m_alc = (byte)(nHorizontalAlign & 0x7);
				m_alcV = (byte)(nVerticalAlign & 0x7);
				m_dgLeft = (uint)(nLeftBorderType & 0xF);
				m_dgRight = (uint)(nRightBorderType & 0xF);
				m_dgTop = (uint)(nTopBorderType & 0xF);
				m_dgBottom = (uint)(nBottomBorderType & 0xF);
				m_icvLeft = (uint)(nLeftBorderPaletteIndex & 0x7F);
				m_icvRight = (uint)(nRightBorderPaletteIndex & 0x7F);
				m_icvTop = (uint)(nTopBorderPaletteIndex & 0x7F);
				m_icvBottom = (uint)(nBottomBorderPaletteIndex & 0x7F);
				if (bHasXFExt)
					m_fHasXFExt = 0x1;
			}

			public ushort GetFontIndex()
			{
				return m_ifnt;
			}

			public ushort GetFormatIndex()
			{
				return m_ifmt;
			}

			public ushort GetBackgroundPaletteIndex()
			{
				if (m_fls > 0)
					return m_icvFore;
				return BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT;
			}

			public Style.FillPattern GetFillPattern()
			{
				switch (m_fls)
				{
					case 0x02:
					{
						return Style.FillPattern.FILL_PATTERN_50;
					}

					case 0x03:
					{
						return Style.FillPattern.FILL_PATTERN_75;
					}

					case 0x04:
					{
						return Style.FillPattern.FILL_PATTERN_25;
					}

					case 0x05:
					{
						return Style.FillPattern.FILL_PATTERN_HORIZONTAL_STRIPE;
					}

					case 0x06:
					{
						return Style.FillPattern.FILL_PATTERN_VARTICAL_STRIPE;
					}

					case 0x07:
					{
						return Style.FillPattern.FILL_PATTERN_REVERSE_DIAGONAL_STRIPE;
					}

					case 0x08:
					{
						return Style.FillPattern.FILL_PATTERN_DIAGONAL_STRIPE;
					}

					case 0x09:
					{
						return Style.FillPattern.FILL_PATTERN_DIAGONAL_CROSSHATCH;
					}

					case 0x0A:
					{
						return Style.FillPattern.FILL_PATTERN_THICK_DIAGONAL_CROSSHATCH;
					}

					case 0x0B:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_HORIZONTAL_STRIPE;
					}

					case 0x0C:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_VERTICAL_STRIPE;
					}

					case 0x0D:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_REVERSE_VERTICAL_STRIPE;
					}

					case 0x0E:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_DIAGONAL_STRIPE;
					}

					case 0x0F:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_HORIZONTAL_CROSSHATCH;
					}

					case 0x10:
					{
						return Style.FillPattern.FILL_PATTERN_THIN_DIAGONAL_CROSSHATCH;
					}

					case 0x11:
					{
						return Style.FillPattern.FILL_PATTERN_125;
					}

					case 0x12:
					{
						return Style.FillPattern.FILL_PATTERN_625;
					}

				}
				return Style.FillPattern.FILL_PATTERN_NONE;
			}

			public ushort GetForegroundPaletteIndex()
			{
				if (m_fls > 0)
					return m_icvBack;
				return BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT;
			}

			public byte GetHorizontalAlign()
			{
				return m_alc;
			}

			public byte GetVerticalAlign()
			{
				return m_alcV;
			}

			protected Line.Type BorderStyleToLineType(byte nBorderStyle)
			{
				switch (nBorderStyle)
				{
					case 0x0000:
					{
						return Line.Type.TYPE_NONE;
					}

					case 0x0001:
					{
						return Line.Type.TYPE_THIN;
					}

					case 0x0002:
					{
						return Line.Type.TYPE_MEDIUM;
					}

					case 0x0003:
					{
						return Line.Type.TYPE_DASHED;
					}

					case 0x0004:
					{
						return Line.Type.TYPE_DOTTED;
					}

					case 0x0005:
					{
						return Line.Type.TYPE_THICK;
					}

					case 0x0006:
					{
						return Line.Type.TYPE_DOUBLE;
					}

					case 0x0007:
					{
						return Line.Type.TYPE_HAIR;
					}

					case 0x0008:
					{
						return Line.Type.TYPE_MEDIUM_DASHED;
					}

					case 0x0009:
					{
						return Line.Type.TYPE_DASH_DOT;
					}

					case 0x000A:
					{
						return Line.Type.TYPE_MEDIUM_DASH_DOT;
					}

					case 0x000B:
					{
						return Line.Type.TYPE_DASH_DOT_DOT;
					}

					case 0x000C:
					{
						return Line.Type.TYPE_MEDIUM_DASH_DOT_DOT;
					}

					case 0x000D:
					{
						return Line.Type.TYPE_SLANT_DASH_DOT_DOT;
					}

				}
				return Line.Type.TYPE_NONE;
			}

			protected byte LineTypeToBorderStyle(Line.Type eLineType)
			{
				switch (eLineType)
				{
					case Line.Type.TYPE_NONE:
					{
						return 0x0000;
					}

					case Line.Type.TYPE_THIN:
					{
						return 0x0001;
					}

					case Line.Type.TYPE_MEDIUM:
					{
						return 0x0002;
					}

					case Line.Type.TYPE_DASHED:
					{
						return 0x0003;
					}

					case Line.Type.TYPE_DOTTED:
					{
						return 0x0004;
					}

					case Line.Type.TYPE_THICK:
					{
						return 0x0005;
					}

					case Line.Type.TYPE_DOUBLE:
					{
						return 0x0006;
					}

					case Line.Type.TYPE_HAIR:
					{
						return 0x0007;
					}

					case Line.Type.TYPE_MEDIUM_DASHED:
					{
						return 0x0008;
					}

					case Line.Type.TYPE_DASH_DOT:
					{
						return 0x0009;
					}

					case Line.Type.TYPE_MEDIUM_DASH_DOT:
					{
						return 0x000A;
					}

					case Line.Type.TYPE_DASH_DOT_DOT:
					{
						return 0x000B;
					}

					case Line.Type.TYPE_MEDIUM_DASH_DOT_DOT:
					{
						return 0x000C;
					}

					case Line.Type.TYPE_SLANT_DASH_DOT_DOT:
					{
						return 0x000D;
					}

				}
				return 0x0000;
			}

			public Line.Type GetTopBorderType()
			{
				return BorderStyleToLineType((byte)(m_dgTop));
			}

			public Line.Type GetRightBorderType()
			{
				return BorderStyleToLineType((byte)(m_dgRight));
			}

			public Line.Type GetBottomBorderType()
			{
				return BorderStyleToLineType((byte)(m_dgBottom));
			}

			public Line.Type GetLeftBorderType()
			{
				return BorderStyleToLineType((byte)(m_dgLeft));
			}

			public ushort GetTopBorderPaletteIndex()
			{
				return (ushort)(m_icvTop);
			}

			public ushort GetRightBorderPaletteIndex()
			{
				return (ushort)(m_icvRight);
			}

			public ushort GetBottomBorderPaletteIndex()
			{
				return (ushort)(m_icvBottom);
			}

			public ushort GetLeftBorderPaletteIndex()
			{
				return (ushort)(m_icvLeft);
			}

			public bool GetHasXFExt()
			{
				return m_fHasXFExt == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WsBoolRecord : BiffRecord
		{
			public WsBoolRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public WsBoolRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_WSBOOL);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_haxFlags = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_haxFlags);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_WSBOOL;
			protected void SetDefaults()
			{
				m_haxFlags = 1217;
			}

			protected ushort m_haxFlags;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WriteAccess : BiffRecord
		{
			public WriteAccess(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_WRITE_ACCESS);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_userName.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_userName.BlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 3;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_WRITE_ACCESS;
			protected void SetDefaults()
			{
				m_userName = new XLUnicodeStringStruct();
			}

			protected XLUnicodeStringStruct m_userName;
			protected const uint FORCE_SIZE = 112;
			public WriteAccess() : base(TYPE, SIZE)
			{
				SetDefaults();
				m_userName.m_rgb.Set("Number Duck");
				m_pHeader.m_nSize = FORCE_SIZE;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				while (pBlobView.GetSize() < (int)(FORCE_SIZE))
				{
					pBlobView.PackUint8(0);
				}
			}

			~WriteAccess()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Window2Record : BiffRecord
		{
			public Window2Record(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_WINDOW2);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fDspFmlaRt = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fDspGridRt = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fDspRwColRt = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fFrozenRt = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fDspZerosRt = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fDefaultHdr = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fRightToLeft = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fDspGuts = (ushort)((nBitmask0 >> 7) & 0x1);
				m_fFrozenNoSplit = (ushort)((nBitmask0 >> 8) & 0x1);
				m_fSelected = (ushort)((nBitmask0 >> 9) & 0x1);
				m_fPaged = (ushort)((nBitmask0 >> 10) & 0x1);
				m_fSLV = (ushort)((nBitmask0 >> 11) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 12) & 0xf);
				m_rwTop = pBlobView.UnpackUint16();
				m_colLeft = pBlobView.UnpackUint16();
				m_icvHdr = pBlobView.UnpackUint16();
				m_reserved2 = pBlobView.UnpackUint16();
				m_wScaleSLV = pBlobView.UnpackUint16();
				m_wScaleNormal = pBlobView.UnpackUint16();
				m_unused = pBlobView.UnpackUint16();
				m_reserved3 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fDspFmlaRt << 0;
				nBitmask0 += m_fDspGridRt << 1;
				nBitmask0 += m_fDspRwColRt << 2;
				nBitmask0 += m_fFrozenRt << 3;
				nBitmask0 += m_fDspZerosRt << 4;
				nBitmask0 += m_fDefaultHdr << 5;
				nBitmask0 += m_fRightToLeft << 6;
				nBitmask0 += m_fDspGuts << 7;
				nBitmask0 += m_fFrozenNoSplit << 8;
				nBitmask0 += m_fSelected << 9;
				nBitmask0 += m_fPaged << 10;
				nBitmask0 += m_fSLV << 11;
				nBitmask0 += m_reserved1 << 12;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_rwTop);
				pBlobView.PackUint16(m_colLeft);
				pBlobView.PackUint16(m_icvHdr);
				pBlobView.PackUint16(m_reserved2);
				pBlobView.PackUint16(m_wScaleSLV);
				pBlobView.PackUint16(m_wScaleNormal);
				pBlobView.PackUint16(m_unused);
				pBlobView.PackUint16(m_reserved3);
			}

			protected const ushort SIZE = 18;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_WINDOW2;
			protected void SetDefaults()
			{
				m_fDspFmlaRt = 0;
				m_fDspGridRt = 0;
				m_fDspRwColRt = 1;
				m_fFrozenRt = 0;
				m_fDspZerosRt = 1;
				m_fDefaultHdr = 1;
				m_fRightToLeft = 0;
				m_fDspGuts = 1;
				m_fFrozenNoSplit = 0;
				m_fSelected = 0;
				m_fPaged = 0;
				m_fSLV = 0;
				m_reserved1 = 0;
				m_rwTop = 0;
				m_colLeft = 0;
				m_icvHdr = 64;
				m_reserved2 = 0;
				m_wScaleSLV = 0;
				m_wScaleNormal = 0;
				m_unused = 0;
				m_reserved3 = 0;
			}

			protected ushort m_fDspFmlaRt;
			protected ushort m_fDspGridRt;
			protected ushort m_fDspRwColRt;
			protected ushort m_fFrozenRt;
			protected ushort m_fDspZerosRt;
			protected ushort m_fDefaultHdr;
			protected ushort m_fRightToLeft;
			protected ushort m_fDspGuts;
			protected ushort m_fFrozenNoSplit;
			protected ushort m_fSelected;
			protected ushort m_fPaged;
			protected ushort m_fSLV;
			protected ushort m_reserved1;
			protected ushort m_rwTop;
			protected ushort m_colLeft;
			protected ushort m_icvHdr;
			protected ushort m_reserved2;
			protected ushort m_wScaleSLV;
			protected ushort m_wScaleNormal;
			protected ushort m_unused;
			protected ushort m_reserved3;
			public Window2Record(bool fSelected, bool fPaged, bool bShowGridlines) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_fSelected = fSelected ? (ushort)(0x1) : (ushort)(0x0);
				m_fPaged = fPaged ? (ushort)(0x1) : (ushort)(0x0);
				m_fDspGridRt = bShowGridlines ? (ushort)(0x1) : (ushort)(0x0);
			}

			public bool GetShowGridlines()
			{
				return m_fDspGridRt == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Window1 : BiffRecord
		{
			public Window1(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_WINDOW1);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_xWn = pBlobView.UnpackInt16();
				m_yWn = pBlobView.UnpackInt16();
				m_dxWn = pBlobView.UnpackInt16();
				m_dyWn = pBlobView.UnpackInt16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fHidden = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fIconic = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fVeryHidden = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fDspHScroll = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fDspVScroll = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fBotAdornment = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fNoAFDateGroup = (ushort)((nBitmask0 >> 6) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 7) & 0x1ff);
				m_itabCur = pBlobView.UnpackUint16();
				m_itabFirst = pBlobView.UnpackUint16();
				m_ctabSel = pBlobView.UnpackUint16();
				m_wTabRatio = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_xWn);
				pBlobView.PackInt16(m_yWn);
				pBlobView.PackInt16(m_dxWn);
				pBlobView.PackInt16(m_dyWn);
				int nBitmask0 = 0;
				nBitmask0 += m_fHidden << 0;
				nBitmask0 += m_fIconic << 1;
				nBitmask0 += m_fVeryHidden << 2;
				nBitmask0 += m_fDspHScroll << 3;
				nBitmask0 += m_fDspVScroll << 4;
				nBitmask0 += m_fBotAdornment << 5;
				nBitmask0 += m_fNoAFDateGroup << 6;
				nBitmask0 += m_reserved << 7;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_itabCur);
				pBlobView.PackUint16(m_itabFirst);
				pBlobView.PackUint16(m_ctabSel);
				pBlobView.PackUint16(m_wTabRatio);
			}

			protected const ushort SIZE = 18;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_WINDOW1;
			protected void SetDefaults()
			{
				m_xWn = 0;
				m_yWn = 0;
				m_dxWn = 0;
				m_dyWn = 0;
				m_fHidden = 0;
				m_fIconic = 0;
				m_fVeryHidden = 0;
				m_fDspHScroll = 0;
				m_fDspVScroll = 0;
				m_fBotAdornment = 0;
				m_fNoAFDateGroup = 0;
				m_reserved = 0;
				m_itabCur = 0;
				m_itabFirst = 0;
				m_ctabSel = 0;
				m_wTabRatio = 0;
			}

			protected short m_xWn;
			protected short m_yWn;
			protected short m_dxWn;
			protected short m_dyWn;
			protected ushort m_fHidden;
			protected ushort m_fIconic;
			protected ushort m_fVeryHidden;
			protected ushort m_fDspHScroll;
			protected ushort m_fDspVScroll;
			protected ushort m_fBotAdornment;
			protected ushort m_fNoAFDateGroup;
			protected ushort m_reserved;
			protected ushort m_itabCur;
			protected ushort m_itabFirst;
			protected ushort m_ctabSel;
			protected ushort m_wTabRatio;
			public Window1(ushort nTabRatio) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_xWn = 480;
				m_yWn = 105;
				m_dxWn = 11475;
				m_dyWn = 7740;
				m_fDspHScroll = 0x1;
				m_fDspVScroll = 0x1;
				m_fBotAdornment = 0x1;
				m_ctabSel = 1;
				m_wTabRatio = nTabRatio;
			}

			public ushort GetTabRatio()
			{
				return m_wTabRatio;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WinProtect : BiffRecord
		{
			public WinProtect() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public WinProtect(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_WIN_PROTECT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fLockWn = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fLockWn);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_WIN_PROTECT;
			protected void SetDefaults()
			{
				m_fLockWn = 0;
			}

			protected ushort m_fLockWn;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ValueRangeRecord : BiffRecord
		{
			public ValueRangeRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ValueRangeRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ValueRange);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_numMin = pBlobView.UnpackDouble();
				m_numMax = pBlobView.UnpackDouble();
				m_numMajor = pBlobView.UnpackDouble();
				m_numMinor = pBlobView.UnpackDouble();
				m_numCross = pBlobView.UnpackDouble();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoMin = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fAutoMax = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fAutoMajor = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fAutoMinor = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fAutoCross = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fLog = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fReversed = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fMaxCross = (ushort)((nBitmask0 >> 7) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 8) & 0xff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackDouble(m_numMin);
				pBlobView.PackDouble(m_numMax);
				pBlobView.PackDouble(m_numMajor);
				pBlobView.PackDouble(m_numMinor);
				pBlobView.PackDouble(m_numCross);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoMin << 0;
				nBitmask0 += m_fAutoMax << 1;
				nBitmask0 += m_fAutoMajor << 2;
				nBitmask0 += m_fAutoMinor << 3;
				nBitmask0 += m_fAutoCross << 4;
				nBitmask0 += m_fLog << 5;
				nBitmask0 += m_fReversed << 6;
				nBitmask0 += m_fMaxCross << 7;
				nBitmask0 += m_unused << 8;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 42;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ValueRange;
			protected void SetDefaults()
			{
				m_numMin = 0.0;
				m_numMax = 0.0;
				m_numMajor = 0.0;
				m_numMinor = 0.0;
				m_numCross = 0.0;
				m_fAutoMin = 1;
				m_fAutoMax = 1;
				m_fAutoMajor = 1;
				m_fAutoMinor = 1;
				m_fAutoCross = 1;
				m_fLog = 0;
				m_fReversed = 0;
				m_fMaxCross = 0;
				m_unused = 0;
			}

			protected double m_numMin;
			protected double m_numMax;
			protected double m_numMajor;
			protected double m_numMinor;
			protected double m_numCross;
			protected ushort m_fAutoMin;
			protected ushort m_fAutoMax;
			protected ushort m_fAutoMajor;
			protected ushort m_fAutoMinor;
			protected ushort m_fAutoCross;
			protected ushort m_fLog;
			protected ushort m_fReversed;
			protected ushort m_fMaxCross;
			protected ushort m_unused;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class VCenterRecord : BiffRecord
		{
			public VCenterRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_VCENTER);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_vcenter = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_vcenter);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_VCENTER;
			protected void SetDefaults()
			{
				m_vcenter = 0;
			}

			protected ushort m_vcenter;
			public VCenterRecord(bool bCenter) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_vcenter = bCenter ? (ushort)(0x1) : (ushort)(0x0);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class UnitsRecord : BiffRecord
		{
			public UnitsRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public UnitsRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Units);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_unused = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_unused);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Units;
			protected void SetDefaults()
			{
				m_unused = 0;
			}

			protected ushort m_unused;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TopMarginRecord : BiffRecord
		{
			public TopMarginRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_TOP_MARGIN);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_num = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackDouble(m_num);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_TOP_MARGIN;
			protected void SetDefaults()
			{
				m_num = 0.0;
			}

			protected double m_num;
			public TopMarginRecord(double num) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_num = num;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TickRecord : BiffRecord
		{
			public TickRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Tick);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_tktMajor = pBlobView.UnpackUint8();
				m_tktMinor = pBlobView.UnpackUint8();
				m_tlt = pBlobView.UnpackUint8();
				m_wBkgMode = pBlobView.UnpackUint8();
				m_rgb = pBlobView.UnpackUint32();
				m_reserved1 = pBlobView.UnpackUint32();
				m_reserved2 = pBlobView.UnpackUint32();
				m_reserved3 = pBlobView.UnpackUint32();
				m_reserved4 = pBlobView.UnpackUint32();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoCo = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fAutoMode = (ushort)((nBitmask0 >> 1) & 0x1);
				m_rot = (ushort)((nBitmask0 >> 2) & 0x7);
				m_fAutoRot = (ushort)((nBitmask0 >> 5) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 6) & 0xff);
				m_iReadingOrder = (ushort)((nBitmask0 >> 14) & 0x3);
				m_icv = pBlobView.UnpackUint16();
				m_trot = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_tktMajor);
				pBlobView.PackUint8(m_tktMinor);
				pBlobView.PackUint8(m_tlt);
				pBlobView.PackUint8(m_wBkgMode);
				pBlobView.PackUint32(m_rgb);
				pBlobView.PackUint32(m_reserved1);
				pBlobView.PackUint32(m_reserved2);
				pBlobView.PackUint32(m_reserved3);
				pBlobView.PackUint32(m_reserved4);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoCo << 0;
				nBitmask0 += m_fAutoMode << 1;
				nBitmask0 += m_rot << 2;
				nBitmask0 += m_fAutoRot << 5;
				nBitmask0 += m_unused << 6;
				nBitmask0 += m_iReadingOrder << 14;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icv);
				pBlobView.PackUint16(m_trot);
			}

			protected const ushort SIZE = 30;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Tick;
			protected void SetDefaults()
			{
				m_tktMajor = 0;
				m_tktMinor = 0;
				m_tlt = 0;
				m_wBkgMode = 0;
				m_rgb = 0;
				m_reserved1 = 0;
				m_reserved2 = 0;
				m_reserved3 = 0;
				m_reserved4 = 0;
				m_fAutoCo = 0;
				m_fAutoMode = 0;
				m_rot = 0;
				m_fAutoRot = 0;
				m_unused = 0;
				m_iReadingOrder = 0;
				m_icv = 0;
				m_trot = 0;
			}

			protected byte m_tktMajor;
			protected byte m_tktMinor;
			protected byte m_tlt;
			protected byte m_wBkgMode;
			protected uint m_rgb;
			protected uint m_reserved1;
			protected uint m_reserved2;
			protected uint m_reserved3;
			protected uint m_reserved4;
			protected ushort m_fAutoCo;
			protected ushort m_fAutoMode;
			protected ushort m_rot;
			protected ushort m_fAutoRot;
			protected ushort m_unused;
			protected ushort m_iReadingOrder;
			protected ushort m_icv;
			protected ushort m_trot;
			public TickRecord(bool fAutoRot) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_tktMajor = 0x0002;
				m_tlt = 0x0003;
				m_wBkgMode = 0x0001;
				if (fAutoRot)
					m_fAutoRot = 0x1;
				m_icv = 0x08;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Theme : BiffRecord
		{
			public Theme(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_THEME);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_dwThemeVersion = pBlobView.UnpackUint32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint32(m_dwThemeVersion);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_THEME;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_dwThemeVersion = 124226;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected uint m_dwThemeVersion;
			public Vector<uint> m_nColorVector;
			protected void PostSetDefaults()
			{
				m_nColorVector = new Vector<uint>();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_dwThemeVersion == 0)
				{
					XmlFile pXmlFile = new XmlFile();
					Blob pXmlBlob = new Blob(true);
					BlobView pXmlBlobView = pXmlBlob.GetBlobView();
					bool bContinue = true;
					Zip pZip = new Zip();
					if (!pZip.LoadBlobView(pBlobView))
						bContinue = false;
					if (bContinue)
					{
						if (!pZip.ExtractFileByName("theme/theme/theme1.xml", pXmlBlobView))
							bContinue = false;
					}
					if (bContinue)
					{
						pXmlBlobView.SetOffset(0);
						if (!pXmlFile.Load(pXmlBlobView))
							bContinue = false;
					}
					XmlNode pTheme = null;
					XmlNode pThemeElements = null;
					XmlNode pClrScheme = null;
					XmlNode pChild = null;
					XmlNode pClr = null;
					if (bContinue)
					{
						pTheme = pXmlFile.GetFirstChildElement("a:theme");
						if (pTheme == null)
							bContinue = false;
					}
					if (bContinue)
					{
						pThemeElements = pTheme.GetFirstChildElement("a:themeElements");
						if (pThemeElements == null)
							bContinue = false;
					}
					if (bContinue)
					{
						pClrScheme = pThemeElements.GetFirstChildElement("a:clrScheme");
						if (pThemeElements == null)
							bContinue = false;
					}
					if (bContinue)
					{
						pChild = pClrScheme.GetFirstChildElement(null);
						while (pChild != null)
						{
							string szName;
							string szValue;
							uint nColor;
							pClr = pChild.GetFirstChildElement(null);
							if (pClr == null)
							{
								bContinue = false;
								break;
							}
							szName = pClr.GetValue();
							if (ExternalString.Equal(szName, "a:sysClr"))
								szValue = pClr.GetAttribute("lastClr");
							else
								szValue = pClr.GetAttribute("val");
							nColor = (uint)(ExternalString.hextol(szValue)) + ((uint)(0xFF) << 24);
							m_nColorVector.PushBack(nColor);
							pChild = pChild.GetNextSiblingElement(null);
						}
					}
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(false);
			}

			public uint GetColorByIndex(int nIndex)
			{
				if (m_dwThemeVersion == 0)
				{
					nbAssert.Assert(nIndex < m_nColorVector.GetSize());
					return m_nColorVector.Get(nIndex);
				}
				else
				{
					nbAssert.Assert(nIndex < 12);
					uint[] nDefaultArray = {0xFFFFFF, 0x000000, 0xEEECE1, 0x1F497D, 0x4F81BD, 0xC0504D, 0x9BBB59, 0x8064A2, 0x4BACC6, 0xF79646, 0x0000FF, 0x800080};
					return nDefaultArray[nIndex];
				}
			}

			~Theme()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TextRecord : BiffRecord
		{
			public TextRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Text);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_at = pBlobView.UnpackUint8();
				m_vat = pBlobView.UnpackUint8();
				m_wBkgMode = pBlobView.UnpackUint16();
				m_rgbText = pBlobView.UnpackUint32();
				m_x = pBlobView.UnpackInt32();
				m_y = pBlobView.UnpackInt32();
				m_dx = pBlobView.UnpackInt32();
				m_dy = pBlobView.UnpackInt32();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoColor = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fShowKey = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fShowValue = (ushort)((nBitmask0 >> 2) & 0x1);
				m_unused1 = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fAutoText = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fGenerated = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fDeleted = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fAutoMode = (ushort)((nBitmask0 >> 7) & 0x1);
				m_unused2 = (ushort)((nBitmask0 >> 8) & 0x7);
				m_fShowLabelAndPerc = (ushort)((nBitmask0 >> 11) & 0x1);
				m_fShowPercent = (ushort)((nBitmask0 >> 12) & 0x1);
				m_fShowBubbleSizes = (ushort)((nBitmask0 >> 13) & 0x1);
				m_fShowLabel = (ushort)((nBitmask0 >> 14) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 15) & 0x1);
				m_icvText = pBlobView.UnpackUint16();
				ushort nBitmask1 = pBlobView.UnpackUint16();
				m_dlp = (ushort)((nBitmask1 >> 0) & 0xf);
				m_unused3 = (ushort)((nBitmask1 >> 4) & 0x3ff);
				m_iReadingOrder = (ushort)((nBitmask1 >> 14) & 0x3);
				m_trot = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_at);
				pBlobView.PackUint8(m_vat);
				pBlobView.PackUint16(m_wBkgMode);
				pBlobView.PackUint32(m_rgbText);
				pBlobView.PackInt32(m_x);
				pBlobView.PackInt32(m_y);
				pBlobView.PackInt32(m_dx);
				pBlobView.PackInt32(m_dy);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoColor << 0;
				nBitmask0 += m_fShowKey << 1;
				nBitmask0 += m_fShowValue << 2;
				nBitmask0 += m_unused1 << 3;
				nBitmask0 += m_fAutoText << 4;
				nBitmask0 += m_fGenerated << 5;
				nBitmask0 += m_fDeleted << 6;
				nBitmask0 += m_fAutoMode << 7;
				nBitmask0 += m_unused2 << 8;
				nBitmask0 += m_fShowLabelAndPerc << 11;
				nBitmask0 += m_fShowPercent << 12;
				nBitmask0 += m_fShowBubbleSizes << 13;
				nBitmask0 += m_fShowLabel << 14;
				nBitmask0 += m_reserved << 15;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icvText);
				int nBitmask1 = 0;
				nBitmask1 += m_dlp << 0;
				nBitmask1 += m_unused3 << 4;
				nBitmask1 += m_iReadingOrder << 14;
				pBlobView.PackUint16((ushort)(nBitmask1));
				pBlobView.PackUint16(m_trot);
			}

			protected const ushort SIZE = 32;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Text;
			protected void SetDefaults()
			{
				m_at = 0;
				m_vat = 0;
				m_wBkgMode = 0;
				m_rgbText = 0;
				m_x = 0;
				m_y = 0;
				m_dx = 0;
				m_dy = 0;
				m_fAutoColor = 0;
				m_fShowKey = 0;
				m_fShowValue = 0;
				m_unused1 = 0;
				m_fAutoText = 0;
				m_fGenerated = 0;
				m_fDeleted = 0;
				m_fAutoMode = 0;
				m_unused2 = 0;
				m_fShowLabelAndPerc = 0;
				m_fShowPercent = 0;
				m_fShowBubbleSizes = 0;
				m_fShowLabel = 0;
				m_reserved = 0;
				m_icvText = 0;
				m_dlp = 0;
				m_unused3 = 0;
				m_iReadingOrder = 0;
				m_trot = 0;
			}

			protected byte m_at;
			protected byte m_vat;
			protected ushort m_wBkgMode;
			protected uint m_rgbText;
			protected int m_x;
			protected int m_y;
			protected int m_dx;
			protected int m_dy;
			protected ushort m_fAutoColor;
			protected ushort m_fShowKey;
			protected ushort m_fShowValue;
			protected ushort m_unused1;
			protected ushort m_fAutoText;
			protected ushort m_fGenerated;
			protected ushort m_fDeleted;
			protected ushort m_fAutoMode;
			protected ushort m_unused2;
			protected ushort m_fShowLabelAndPerc;
			protected ushort m_fShowPercent;
			protected ushort m_fShowBubbleSizes;
			protected ushort m_fShowLabel;
			protected ushort m_reserved;
			protected ushort m_icvText;
			protected ushort m_dlp;
			protected ushort m_unused3;
			protected ushort m_iReadingOrder;
			protected ushort m_trot;
			public TextRecord(int x, int y, uint rgbText, bool fAutoColor, bool fAutoText, bool fGenerated, bool fAutoMode, ushort icvText, ushort trot) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_at = 0x02;
				m_vat = 0x02;
				m_wBkgMode = 0x0001;
				m_x = x;
				m_y = y;
				m_rgbText = rgbText;
				if (fAutoColor)
					m_fAutoColor = 0x1;
				if (fAutoText)
					m_fAutoText = 0x1;
				if (fGenerated)
					m_fGenerated = 0x1;
				if (fAutoMode)
					m_fAutoMode = 0x1;
				m_icvText = icvText;
				m_trot = trot;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SupBookRecord : BiffRecord
		{
			public SupBookRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SUP_BOOK);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ctab = pBlobView.UnpackUint16();
				m_cch = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ctab);
				pBlobView.PackUint16(m_cch);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SUP_BOOK;
			protected void SetDefaults()
			{
				m_ctab = 0;
				m_cch = 0;
			}

			protected ushort m_ctab;
			protected ushort m_cch;
			public SupBookRecord(ushort nNumWorksheet) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_ctab = nNumWorksheet;
				m_cch = 0x0401;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StyleRecord : BiffRecord
		{
			public StyleRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_STYLE);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_ixfe = (ushort)((nBitmask0 >> 0) & 0xfff);
				m_unused = (ushort)((nBitmask0 >> 12) & 0x7);
				m_fBuiltIn = (ushort)((nBitmask0 >> 15) & 0x1);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_ixfe << 0;
				nBitmask0 += m_unused << 12;
				nBitmask0 += m_fBuiltIn << 15;
				pBlobView.PackUint16((ushort)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_STYLE;
			protected void SetDefaults()
			{
				m_ixfe = 0;
				m_unused = 0;
				m_fBuiltIn = 0;
				PostSetDefaults();
			}

			protected ushort m_ixfe;
			protected ushort m_unused;
			protected ushort m_fBuiltIn;
			protected BuiltInStyleStruct m_builtInData;
			protected XLUnicodeStringStruct m_user;
			public StyleRecord(ushort nXFIndex, string szName) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_ixfe = nXFIndex;
				m_user.m_rgb.Set(szName);
				m_builtInData.m_istyBuiltIn = 0;
				m_builtInData.m_iLevel = 0xFF;
				if (m_user.m_rgb.GetLength() == 0)
				{
					m_fBuiltIn = 0x1;
					m_pHeader.m_nSize += BuiltInStyleStruct.SIZE;
				}
				else
				{
					m_fBuiltIn = 0x0;
					m_pHeader.m_nSize += (uint)(XLUnicodeStringStruct.SIZE + m_user.GetDynamicSize());
				}
			}

			protected void PostSetDefaults()
			{
				m_builtInData = new BuiltInStyleStruct();
				m_user = new XLUnicodeStringStruct();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_fBuiltIn != 0)
					m_builtInData.BlobRead(pBlobView);
				else
					m_user.BlobRead(pBlobView);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_fBuiltIn != 0)
					m_builtInData.BlobWrite(pBlobView);
				else
					m_user.BlobWrite(pBlobView);
			}

			~StyleRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StartBlockRecord : BiffRecord
		{
			public StartBlockRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_StartBlock);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobRead(pBlobView);
				m_iObjectKind = pBlobView.UnpackUint16();
				m_iObjectContext = pBlobView.UnpackUint16();
				m_iObjectInstance1 = pBlobView.UnpackUint16();
				m_iObjectInstance2 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_iObjectKind);
				pBlobView.PackUint16(m_iObjectContext);
				pBlobView.PackUint16(m_iObjectInstance1);
				pBlobView.PackUint16(m_iObjectInstance2);
			}

			protected const ushort SIZE = 12;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_StartBlock;
			protected void SetDefaults()
			{
				m_frtHeaderOld = new FrtHeaderOldStruct();
				m_iObjectKind = 0;
				m_iObjectContext = 0;
				m_iObjectInstance1 = 0;
				m_iObjectInstance2 = 0;
				PostSetDefaults();
			}

			protected FrtHeaderOldStruct m_frtHeaderOld;
			protected ushort m_iObjectKind;
			protected ushort m_iObjectContext;
			protected ushort m_iObjectInstance1;
			protected ushort m_iObjectInstance2;
			public StartBlockRecord(ushort iObjectKind, ushort iObjectContext, ushort iObjectInstance1, ushort iObjectInstance2) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_iObjectKind = iObjectKind;
				m_iObjectContext = iObjectContext;
				m_iObjectInstance1 = iObjectInstance1;
				m_iObjectInstance2 = iObjectInstance2;
			}

			protected void PostSetDefaults()
			{
				m_frtHeaderOld.m_rt = 2130;
			}

			~StartBlockRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SstRecord : BiffRecord
		{
			public SstRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SST);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cstTotal = pBlobView.UnpackInt32();
				m_cstUnique = pBlobView.UnpackInt32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt32(m_cstTotal);
				pBlobView.PackInt32(m_cstUnique);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SST;
			protected void SetDefaults()
			{
				m_cstTotal = 0;
				m_cstUnique = 0;
				PostSetDefaults();
			}

			protected int m_cstTotal;
			protected int m_cstUnique;
			protected Blob m_pHaxBlob;
			protected OwnedVector<XLUnicodeRichExtendedString> m_pRgb;
			public SstRecord(SharedStringContainer pSharedStringContainer) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cstTotal = pSharedStringContainer.GetSize();
				m_cstUnique = pSharedStringContainer.GetSize();
				m_pHaxBlob = new Blob(true);
				BlobView pHaxBlobView = m_pHaxBlob.GetBlobView();
				for (int i = 0; i < pSharedStringContainer.GetSize(); i++)
				{
					XLUnicodeRichExtendedString pXLUnicodeRichExtendedString = new XLUnicodeRichExtendedString(pSharedStringContainer.Get(i));
					pXLUnicodeRichExtendedString.ContinueAwareBlobWrite(pHaxBlobView, m_pContinueInfoVector);
					{
						NumberDuck.Secret.XLUnicodeRichExtendedString __555132520 = pXLUnicodeRichExtendedString;
						pXLUnicodeRichExtendedString = null;
						m_pRgb.PushBack(__555132520);
					}
				}
				m_pHeader.m_nSize += (uint)(pHaxBlobView.GetSize());
			}

			protected void PostSetDefaults()
			{
				m_pHaxBlob = null;
				m_pRgb = new OwnedVector<XLUnicodeRichExtendedString>();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_cstUnique; i++)
				{
					XLUnicodeRichExtendedString pXLUnicodeRichExtendedString = new XLUnicodeRichExtendedString();
					pXLUnicodeRichExtendedString.ContinueAwareBlobRead(pBlobView, m_pContinueInfoVector);
					{
						NumberDuck.Secret.XLUnicodeRichExtendedString __555132520 = pXLUnicodeRichExtendedString;
						pXLUnicodeRichExtendedString = null;
						m_pRgb.PushBack(__555132520);
					}
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_pHaxBlob.GetBlobView().SetOffset(0);
				pBlobView.Pack(m_pHaxBlob.GetBlobView(), m_pHaxBlob.GetSize());
			}

			public int GetNumString()
			{
				return m_cstUnique;
			}

			public string GetString(int nIndex)
			{
				return m_pRgb.Get(nIndex).m_rgb.GetExternalString();
			}

			~SstRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShtPropsRecord : BiffRecord
		{
			public ShtPropsRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ShtPropsRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ShtProps);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fManSerAlloc = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fPlotVisOnly = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fNotSizeWith = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fManPlotArea = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fAlwaysAutoPlotArea = (ushort)((nBitmask0 >> 4) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 5) & 0x7ff);
				m_mdBlank = pBlobView.UnpackUint8();
				m_reserved2 = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fManSerAlloc << 0;
				nBitmask0 += m_fPlotVisOnly << 1;
				nBitmask0 += m_fNotSizeWith << 2;
				nBitmask0 += m_fManPlotArea << 3;
				nBitmask0 += m_fAlwaysAutoPlotArea << 4;
				nBitmask0 += m_reserved1 << 5;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint8(m_mdBlank);
				pBlobView.PackUint8(m_reserved2);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ShtProps;
			protected void SetDefaults()
			{
				m_fManSerAlloc = 1;
				m_fPlotVisOnly = 1;
				m_fNotSizeWith = 0;
				m_fManPlotArea = 1;
				m_fAlwaysAutoPlotArea = 1;
				m_reserved1 = 0;
				m_mdBlank = 0;
				m_reserved2 = 0;
			}

			protected ushort m_fManSerAlloc;
			protected ushort m_fPlotVisOnly;
			protected ushort m_fNotSizeWith;
			protected ushort m_fManPlotArea;
			protected ushort m_fAlwaysAutoPlotArea;
			protected ushort m_reserved1;
			protected byte m_mdBlank;
			protected byte m_reserved2;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShapePropsStreamRecord : BiffRecord
		{
			public ShapePropsStreamRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ShapePropsStream);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_wObjContext = pBlobView.UnpackUint16();
				m_unused = pBlobView.UnpackUint16();
				m_dwChecksum = pBlobView.UnpackUint32();
				m_cb = pBlobView.UnpackUint32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_wObjContext);
				pBlobView.PackUint16(m_unused);
				pBlobView.PackUint32(m_dwChecksum);
				pBlobView.PackUint32(m_cb);
			}

			protected const ushort SIZE = 24;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ShapePropsStream;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_wObjContext = 0;
				m_unused = 0;
				m_dwChecksum = 0;
				m_cb = 0;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected ushort m_wObjContext;
			protected ushort m_unused;
			protected uint m_dwChecksum;
			protected uint m_cb;
			protected InternalString m_rgb;
			public ShapePropsStreamRecord(ushort wObjContext, ushort unused, uint dwChecksum) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_wObjContext = wObjContext;
				m_unused = unused;
				m_dwChecksum = dwChecksum;
			}

			protected void PostSetDefaults()
			{
				m_rgb = new InternalString("");
				m_frtHeader.m_rt = 2212;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				uint i;
				for (i = 0; i < m_cb; i++)
					m_rgb.AppendChar((char)(pBlobView.UnpackUint8()));
			}

			~ShapePropsStreamRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SetupRecord : BiffRecord
		{
			public SetupRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SETUP);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_iPaperSize = pBlobView.UnpackUint16();
				m_iScale = pBlobView.UnpackUint16();
				m_iPageStart = pBlobView.UnpackInt16();
				m_iFitWidth = pBlobView.UnpackUint16();
				m_iFitHeight = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fLeftToRight = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fPortrait = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fNoPls = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fNoColor = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fDraft = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fNotes = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fNoOrient = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fUsePage = (ushort)((nBitmask0 >> 7) & 0x1);
				m_unused1 = (ushort)((nBitmask0 >> 8) & 0x1);
				m_fEndNotes = (ushort)((nBitmask0 >> 9) & 0x1);
				m_iErrors = (ushort)((nBitmask0 >> 10) & 0x3);
				m_reserved = (ushort)((nBitmask0 >> 12) & 0xf);
				m_iRes = pBlobView.UnpackUint16();
				m_iVRes = pBlobView.UnpackUint16();
				m_numHdr = pBlobView.UnpackDouble();
				m_numFtr = pBlobView.UnpackDouble();
				m_iCopies = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_iPaperSize);
				pBlobView.PackUint16(m_iScale);
				pBlobView.PackInt16(m_iPageStart);
				pBlobView.PackUint16(m_iFitWidth);
				pBlobView.PackUint16(m_iFitHeight);
				int nBitmask0 = 0;
				nBitmask0 += m_fLeftToRight << 0;
				nBitmask0 += m_fPortrait << 1;
				nBitmask0 += m_fNoPls << 2;
				nBitmask0 += m_fNoColor << 3;
				nBitmask0 += m_fDraft << 4;
				nBitmask0 += m_fNotes << 5;
				nBitmask0 += m_fNoOrient << 6;
				nBitmask0 += m_fUsePage << 7;
				nBitmask0 += m_unused1 << 8;
				nBitmask0 += m_fEndNotes << 9;
				nBitmask0 += m_iErrors << 10;
				nBitmask0 += m_reserved << 12;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_iRes);
				pBlobView.PackUint16(m_iVRes);
				pBlobView.PackDouble(m_numHdr);
				pBlobView.PackDouble(m_numFtr);
				pBlobView.PackUint16(m_iCopies);
			}

			protected const ushort SIZE = 34;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SETUP;
			protected void SetDefaults()
			{
				m_iPaperSize = 1;
				m_iScale = 100;
				m_iPageStart = 1;
				m_iFitWidth = 1;
				m_iFitHeight = 1;
				m_fLeftToRight = 0;
				m_fPortrait = 0;
				m_fNoPls = 0;
				m_fNoColor = 0;
				m_fDraft = 0;
				m_fNotes = 0;
				m_fNoOrient = 0;
				m_fUsePage = 0;
				m_unused1 = 0;
				m_fEndNotes = 0;
				m_iErrors = 0;
				m_reserved = 0;
				m_iRes = 600;
				m_iVRes = 600;
				m_numHdr = 0.3;
				m_numFtr = 0.3;
				m_iCopies = 1;
			}

			protected ushort m_iPaperSize;
			protected ushort m_iScale;
			protected short m_iPageStart;
			protected ushort m_iFitWidth;
			protected ushort m_iFitHeight;
			protected ushort m_fLeftToRight;
			protected ushort m_fPortrait;
			protected ushort m_fNoPls;
			protected ushort m_fNoColor;
			protected ushort m_fDraft;
			protected ushort m_fNotes;
			protected ushort m_fNoOrient;
			protected ushort m_fUsePage;
			protected ushort m_unused1;
			protected ushort m_fEndNotes;
			protected ushort m_iErrors;
			protected ushort m_reserved;
			protected ushort m_iRes;
			protected ushort m_iVRes;
			protected double m_numHdr;
			protected double m_numFtr;
			protected ushort m_iCopies;
			public SetupRecord(bool bPortrait) : base(TYPE, SIZE)
			{
				SetDefaults();
				if (bPortrait)
					m_fPortrait = 0x1;
			}

			public bool GetPortrait()
			{
				return m_fPortrait == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesTextRecord : BiffRecord
		{
			public SeriesTextRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SeriesText);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_reserved = pBlobView.UnpackUint16();
				m_stText.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_reserved);
				m_stText.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SeriesText;
			protected void SetDefaults()
			{
				m_reserved = 0;
				m_stText = new ShortXLUnicodeStringStruct();
			}

			protected ushort m_reserved;
			protected ShortXLUnicodeStringStruct m_stText;
			public SeriesTextRecord(string szText) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_stText.m_rgb.Set(szText);
				m_pHeader.m_nSize += (uint)(m_stText.GetDynamicSize());
			}

			~SeriesTextRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesRecord : BiffRecord
		{
			public SeriesRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public SeriesRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Series);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_sdtX = pBlobView.UnpackUint16();
				m_sdtY = pBlobView.UnpackUint16();
				m_cValx = pBlobView.UnpackUint16();
				m_cValy = pBlobView.UnpackUint16();
				m_sdtBSize = pBlobView.UnpackUint16();
				m_cValBSize = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_sdtX);
				pBlobView.PackUint16(m_sdtY);
				pBlobView.PackUint16(m_cValx);
				pBlobView.PackUint16(m_cValy);
				pBlobView.PackUint16(m_sdtBSize);
				pBlobView.PackUint16(m_cValBSize);
			}

			protected const ushort SIZE = 12;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Series;
			protected void SetDefaults()
			{
				m_sdtX = 0x0001;
				m_sdtY = 0x0001;
				m_cValx = 10;
				m_cValy = 10;
				m_sdtBSize = 0x0001;
				m_cValBSize = 0;
			}

			protected ushort m_sdtX;
			protected ushort m_sdtY;
			protected ushort m_cValx;
			protected ushort m_cValy;
			protected ushort m_sdtBSize;
			protected ushort m_cValBSize;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SerToCrtRecord : BiffRecord
		{
			public SerToCrtRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SerToCrt);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_id = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_id);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SerToCrt;
			protected void SetDefaults()
			{
				m_id = 0;
			}

			protected ushort m_id;
			public SerToCrtRecord(ushort id) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_id = id;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SelectionRecord : BiffRecord
		{
			public SelectionRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public SelectionRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SELECTION);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_pnn = pBlobView.UnpackUint8();
				m_rwAct = pBlobView.UnpackUint16();
				m_colAct = pBlobView.UnpackUint16();
				m_irefAct = pBlobView.UnpackUint16();
				m_cref = pBlobView.UnpackUint16();
				m_rwFirst = pBlobView.UnpackUint16();
				m_rwLast = pBlobView.UnpackUint16();
				m_colFirst = pBlobView.UnpackUint8();
				m_colLast = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_pnn);
				pBlobView.PackUint16(m_rwAct);
				pBlobView.PackUint16(m_colAct);
				pBlobView.PackUint16(m_irefAct);
				pBlobView.PackUint16(m_cref);
				pBlobView.PackUint16(m_rwFirst);
				pBlobView.PackUint16(m_rwLast);
				pBlobView.PackUint8(m_colFirst);
				pBlobView.PackUint8(m_colLast);
			}

			protected const ushort SIZE = 15;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SELECTION;
			protected void SetDefaults()
			{
				m_pnn = 3;
				m_rwAct = 0;
				m_colAct = 0;
				m_irefAct = 0;
				m_cref = 1;
				m_rwFirst = 0;
				m_rwLast = 0;
				m_colFirst = 0;
				m_colLast = 0;
			}

			protected byte m_pnn;
			protected ushort m_rwAct;
			protected ushort m_colAct;
			protected ushort m_irefAct;
			protected ushort m_cref;
			protected ushort m_rwFirst;
			protected ushort m_rwLast;
			protected byte m_colFirst;
			protected byte m_colLast;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SclRecord : BiffRecord
		{
			public SclRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public SclRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SCL);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_nscl = pBlobView.UnpackInt16();
				m_dscl = pBlobView.UnpackInt16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_nscl);
				pBlobView.PackInt16(m_dscl);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SCL;
			protected void SetDefaults()
			{
				m_nscl = 1;
				m_dscl = 1;
			}

			protected short m_nscl;
			protected short m_dscl;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ScatterRecord : BiffRecord
		{
			public ScatterRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Scatter);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_pcBubbleSizeRatio = pBlobView.UnpackUint16();
				m_wBubbleSize = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fBubbles = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fShowNegBubbles = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fHasShadow = (ushort)((nBitmask0 >> 2) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 3) & 0x1fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_pcBubbleSizeRatio);
				pBlobView.PackUint16(m_wBubbleSize);
				int nBitmask0 = 0;
				nBitmask0 += m_fBubbles << 0;
				nBitmask0 += m_fShowNegBubbles << 1;
				nBitmask0 += m_fHasShadow << 2;
				nBitmask0 += m_reserved << 3;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 6;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Scatter;
			protected void SetDefaults()
			{
				m_pcBubbleSizeRatio = 100;
				m_wBubbleSize = 0x0001;
				m_fBubbles = 0;
				m_fShowNegBubbles = 0;
				m_fHasShadow = 0;
				m_reserved = 0;
			}

			protected ushort m_pcBubbleSizeRatio;
			protected ushort m_wBubbleSize;
			protected ushort m_fBubbles;
			protected ushort m_fShowNegBubbles;
			protected ushort m_fHasShadow;
			protected ushort m_reserved;
			public ScatterRecord(Chart.Type eType) : base(TYPE, SIZE)
			{
				SetDefaults();
				nbAssert.Assert(eType == Chart.Type.TYPE_SCATTER);
			}

			public Chart.Type GetChartType()
			{
				return Chart.Type.TYPE_SCATTER;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SIIndexRecord : BiffRecord
		{
			public SIIndexRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_SIIndex);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_numIndex = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_numIndex);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_SIIndex;
			protected void SetDefaults()
			{
				m_numIndex = 0;
			}

			protected ushort m_numIndex;
			public SIIndexRecord(ushort numIndex) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_numIndex = numIndex;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RowRecord : BiffRecord
		{
			public RowRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ROW);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw.BlobRead(pBlobView);
				m_colMic = pBlobView.UnpackUint16();
				m_colMac = pBlobView.UnpackUint16();
				m_miyRw = pBlobView.UnpackUint16();
				m_reserved1 = pBlobView.UnpackUint16();
				m_unused1 = pBlobView.UnpackUint16();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_iOutLevel = (byte)((nBitmask0 >> 0) & 0x7);
				m_reserved2 = (byte)((nBitmask0 >> 3) & 0x1);
				m_fCollapsed = (byte)((nBitmask0 >> 4) & 0x1);
				m_fDyZero = (byte)((nBitmask0 >> 5) & 0x1);
				m_fUnsynced = (byte)((nBitmask0 >> 6) & 0x1);
				m_fGhostDirty = (byte)((nBitmask0 >> 7) & 0x1);
				m_reserved3 = pBlobView.UnpackUint8();
				ushort nBitmask1 = pBlobView.UnpackUint16();
				m_ixfe_val = (ushort)((nBitmask1 >> 0) & 0xfff);
				m_fExAsc = (ushort)((nBitmask1 >> 12) & 0x1);
				m_fExDes = (ushort)((nBitmask1 >> 13) & 0x1);
				m_fPhonetic = (ushort)((nBitmask1 >> 14) & 0x1);
				m_unused2 = (ushort)((nBitmask1 >> 15) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rw.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_colMic);
				pBlobView.PackUint16(m_colMac);
				pBlobView.PackUint16(m_miyRw);
				pBlobView.PackUint16(m_reserved1);
				pBlobView.PackUint16(m_unused1);
				int nBitmask0 = 0;
				nBitmask0 += m_iOutLevel << 0;
				nBitmask0 += m_reserved2 << 3;
				nBitmask0 += m_fCollapsed << 4;
				nBitmask0 += m_fDyZero << 5;
				nBitmask0 += m_fUnsynced << 6;
				nBitmask0 += m_fGhostDirty << 7;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint8(m_reserved3);
				int nBitmask1 = 0;
				nBitmask1 += m_ixfe_val << 0;
				nBitmask1 += m_fExAsc << 12;
				nBitmask1 += m_fExDes << 13;
				nBitmask1 += m_fPhonetic << 14;
				nBitmask1 += m_unused2 << 15;
				pBlobView.PackUint16((ushort)(nBitmask1));
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ROW;
			protected void SetDefaults()
			{
				m_rw = new RwStruct();
				m_colMic = 0;
				m_colMac = 0;
				m_miyRw = 0;
				m_reserved1 = 0;
				m_unused1 = 0;
				m_iOutLevel = 0;
				m_reserved2 = 0;
				m_fCollapsed = 0;
				m_fDyZero = 0;
				m_fUnsynced = 1;
				m_fGhostDirty = 0;
				m_reserved3 = 1;
				m_ixfe_val = 15;
				m_fExAsc = 0;
				m_fExDes = 0;
				m_fPhonetic = 0;
				m_unused2 = 0;
				PostSetDefaults();
			}

			protected RwStruct m_rw;
			protected ushort m_colMic;
			protected ushort m_colMac;
			protected ushort m_miyRw;
			protected ushort m_reserved1;
			protected ushort m_unused1;
			protected byte m_iOutLevel;
			protected byte m_reserved2;
			protected byte m_fCollapsed;
			protected byte m_fDyZero;
			protected byte m_fUnsynced;
			protected byte m_fGhostDirty;
			protected byte m_reserved3;
			protected ushort m_ixfe_val;
			protected ushort m_fExAsc;
			protected ushort m_fExDes;
			protected ushort m_fPhonetic;
			protected ushort m_unused2;
			public bool m_bTopMedium;
			public bool m_bTopThick;
			public bool m_bBottomMedium;
			public bool m_bBottomThick;
			public RowRecord(ushort nRow, ushort nHeight) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_rw.m_rw = nRow;
				m_miyRw = nHeight;
			}

			protected void PostSetDefaults()
			{
				m_bTopMedium = false;
				m_bTopThick = false;
				m_bBottomMedium = false;
				m_bBottomThick = false;
			}

			public ushort GetRow()
			{
				return m_rw.m_rw;
			}

			public ushort GetHeight()
			{
				return m_miyRw;
			}

			public void SetHeight(ushort nHeight)
			{
				m_miyRw = nHeight;
			}

			public void SetTopMedium()
			{
				m_bTopMedium = true;
			}

			public void SetTopThick()
			{
				m_fExAsc = 0x1;
				m_bTopThick = true;
			}

			public void SetBottomMedium()
			{
				m_fExDes = 0x1;
				m_bBottomMedium = true;
			}

			public void SetBottomThick()
			{
				m_fExDes = 0x1;
				m_bBottomThick = true;
			}

			public bool GetBottomThick()
			{
				return m_bBottomThick;
			}

			~RowRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RkRecord : BiffRecord
		{
			public RkRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_RK);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_nRowIndex = pBlobView.UnpackUint16();
				m_nColumnIndex = pBlobView.UnpackUint16();
				m_nXfIndex = pBlobView.UnpackUint16();
				m_nRkValue = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_nRowIndex);
				pBlobView.PackUint16(m_nColumnIndex);
				pBlobView.PackUint16(m_nXfIndex);
				pBlobView.PackUint32(m_nRkValue);
			}

			protected const ushort SIZE = 10;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_RK;
			protected void SetDefaults()
			{
				m_nRowIndex = 0;
				m_nColumnIndex = 0;
				m_nXfIndex = 0;
				m_nRkValue = 0;
			}

			protected ushort m_nRowIndex;
			protected ushort m_nColumnIndex;
			protected ushort m_nXfIndex;
			protected uint m_nRkValue;
			public ushort GetX()
			{
				return m_nColumnIndex;
			}

			public ushort GetY()
			{
				return m_nRowIndex;
			}

			public ushort GetXfIndex()
			{
				return m_nXfIndex;
			}

			public uint GetRkValue()
			{
				return m_nRkValue;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RightMarginRecord : BiffRecord
		{
			public RightMarginRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_RIGHT_MARGIN);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_num = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackDouble(m_num);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_RIGHT_MARGIN;
			protected void SetDefaults()
			{
				m_num = 0.0;
			}

			protected double m_num;
			public RightMarginRecord(double num) : base(TYPE, SIZE)
			{
				m_num = num;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RefreshAllRecord : BiffRecord
		{
			public RefreshAllRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public RefreshAllRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_REFRESH_ALL);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_refreshAll = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_refreshAll);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_REFRESH_ALL;
			protected void SetDefaults()
			{
				m_refreshAll = 0;
			}

			protected ushort m_refreshAll;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RRTabId : BiffRecord
		{
			public RRTabId(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_RR_TAB_ID);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_RR_TAB_ID;
			protected void SetDefaults()
			{
			}

			public ushort m_nNumWorksheet;
			public RRTabId(ushort nNumWorksheet) : base(TYPE, (ushort)(SIZE + 2 * nNumWorksheet))
			{
				SetDefaults();
				m_nNumWorksheet = nNumWorksheet;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				m_nNumWorksheet = (ushort)((pBlobView.GetSize() - pBlobView.GetOffset()) >> 1);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				ushort i;
				for (i = 0; i < m_nNumWorksheet; i++)
					pBlobView.PackUint16(i);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ProtectRecord : BiffRecord
		{
			public ProtectRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ProtectRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PROTECT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fLock = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fLock);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PROTECT;
			protected void SetDefaults()
			{
				m_fLock = 0;
			}

			protected ushort m_fLock;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Prot4RevRecord : BiffRecord
		{
			public Prot4RevRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Prot4RevRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PROT_4_REV);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fRevLock = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fRevLock);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PROT_4_REV;
			protected void SetDefaults()
			{
				m_fRevLock = 0;
			}

			protected ushort m_fRevLock;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Prot4RevPassRecord : BiffRecord
		{
			public Prot4RevPassRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Prot4RevPassRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PROT_4_REV_PASS);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_protPwdRev = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_protPwdRev);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PROT_4_REV_PASS;
			protected void SetDefaults()
			{
				m_protPwdRev = 0;
			}

			protected ushort m_protPwdRev;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintSizeRecord : BiffRecord
		{
			public PrintSizeRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PrintSizeRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PRINT_SIZE);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_printSize = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_printSize);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PRINT_SIZE;
			protected void SetDefaults()
			{
				m_printSize = 0x0003;
			}

			protected ushort m_printSize;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintRowColRecord : BiffRecord
		{
			public PrintRowColRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PrintRowColRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PrintRowCol);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_printRwCol = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_printRwCol);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PrintRowCol;
			protected void SetDefaults()
			{
				m_printRwCol = 0;
			}

			protected ushort m_printRwCol;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PrintGridRecord : BiffRecord
		{
			public PrintGridRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PrintGrid);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fPrintGrid = (ushort)((nBitmask0 >> 0) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 1) & 0x7fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fPrintGrid << 0;
				nBitmask0 += m_unused << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PrintGrid;
			protected void SetDefaults()
			{
				m_fPrintGrid = 0;
				m_unused = 0;
			}

			protected ushort m_fPrintGrid;
			protected ushort m_unused;
			public PrintGridRecord(bool bPrintGridlines) : base(TYPE, SIZE)
			{
				SetDefaults();
				if (bPrintGridlines)
					m_fPrintGrid = 0x1;
			}

			public bool GetPrintGridlines()
			{
				return m_fPrintGrid == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PosRecord : BiffRecord
		{
			public PosRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Pos);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_mdTopLt = pBlobView.UnpackUint16();
				m_mdBotRt = pBlobView.UnpackUint16();
				m_x1 = pBlobView.UnpackInt16();
				m_unused1 = pBlobView.UnpackUint16();
				m_y1 = pBlobView.UnpackInt16();
				m_unused2 = pBlobView.UnpackUint16();
				m_x2 = pBlobView.UnpackInt16();
				m_unused3 = pBlobView.UnpackUint16();
				m_y2 = pBlobView.UnpackInt16();
				m_unused4 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_mdTopLt);
				pBlobView.PackUint16(m_mdBotRt);
				pBlobView.PackInt16(m_x1);
				pBlobView.PackUint16(m_unused1);
				pBlobView.PackInt16(m_y1);
				pBlobView.PackUint16(m_unused2);
				pBlobView.PackInt16(m_x2);
				pBlobView.PackUint16(m_unused3);
				pBlobView.PackInt16(m_y2);
				pBlobView.PackUint16(m_unused4);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Pos;
			protected void SetDefaults()
			{
				m_mdTopLt = 0x0002;
				m_mdBotRt = 0x0002;
				m_x1 = 0;
				m_unused1 = 0;
				m_y1 = 0;
				m_unused2 = 0;
				m_x2 = 0;
				m_unused3 = 0;
				m_y2 = 0;
				m_unused4 = 0;
			}

			protected ushort m_mdTopLt;
			protected ushort m_mdBotRt;
			protected short m_x1;
			protected ushort m_unused1;
			protected short m_y1;
			protected ushort m_unused2;
			protected short m_x2;
			protected ushort m_unused3;
			protected short m_y2;
			protected ushort m_unused4;
			public PosRecord(short x1, short y1, ushort unused2, short x2, short y2) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_x1 = x1;
				m_y1 = y1;
				m_unused2 = unused2;
				m_x2 = x2;
				m_y2 = y2;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PlotGrowthRecord : BiffRecord
		{
			public PlotGrowthRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PlotGrowthRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PlotGrowth);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_dxPlotGrowth = pBlobView.UnpackInt32();
				m_dyPlotGrowth = pBlobView.UnpackInt32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt32(m_dxPlotGrowth);
				pBlobView.PackInt32(m_dyPlotGrowth);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PlotGrowth;
			protected void SetDefaults()
			{
				m_dxPlotGrowth = 65536;
				m_dyPlotGrowth = 65536;
			}

			protected int m_dxPlotGrowth;
			protected int m_dyPlotGrowth;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PlotAreaRecord : BiffRecord
		{
			public PlotAreaRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PlotAreaRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PlotArea);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PlotArea;
			protected void SetDefaults()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PieFormatRecord : BiffRecord
		{
			public PieFormatRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PieFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PieFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_pcExplode = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_pcExplode);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PieFormat;
			protected void SetDefaults()
			{
				m_pcExplode = 0;
			}

			protected ushort m_pcExplode;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PasswordRecord : BiffRecord
		{
			public PasswordRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public PasswordRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PASSWORD);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_wPassword = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_wPassword);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PASSWORD;
			protected void SetDefaults()
			{
				m_wPassword = 0;
			}

			protected ushort m_wPassword;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PaletteRecord : BiffRecord
		{
			public PaletteRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_PALETTE);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ccv = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ccv);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_PALETTE;
			protected void SetDefaults()
			{
				m_ccv = 0;
				PostSetDefaults();
			}

			protected ushort m_ccv;
			protected uint[] m_rgColor = new uint[BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY];
			public PaletteRecord() : base(TYPE, SIZE + BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY * 4)
			{
				SetDefaults();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_ccv; i++)
					m_rgColor[i] = pBlobView.UnpackUint32();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_ccv; i++)
					pBlobView.PackUint32(m_rgColor[i]);
			}

			protected void PostSetDefaults()
			{
				m_ccv = BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY;
				for (ushort i = 0; i < m_ccv; i++)
					m_rgColor[i] = BiffWorkbookGlobals.DEFAULT_CUSTOM_COLOR[i] & 0xFFFFFF;
			}

			public uint GetColorByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex <= m_ccv);
				return m_rgColor[nIndex];
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ObjectLinkRecord : BiffRecord
		{
			public ObjectLinkRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ObjectLink);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_wLinkObj = pBlobView.UnpackUint16();
				m_wLinkVar1 = pBlobView.UnpackUint16();
				m_wLinkVar2 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_wLinkObj);
				pBlobView.PackUint16(m_wLinkVar1);
				pBlobView.PackUint16(m_wLinkVar2);
			}

			protected const ushort SIZE = 6;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ObjectLink;
			protected void SetDefaults()
			{
				m_wLinkObj = 0;
				m_wLinkVar1 = 0;
				m_wLinkVar2 = 0;
			}

			protected ushort m_wLinkObj;
			protected ushort m_wLinkVar1;
			protected ushort m_wLinkVar2;
			public ObjectLinkRecord(ushort nLinkObject, ushort nSeriesIndex, ushort nCategoryIndex) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_wLinkObj = nLinkObject;
				m_wLinkVar1 = nSeriesIndex;
				m_wLinkVar2 = nSeriesIndex;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ObjRecord : BiffRecord
		{
			public ObjRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_OBJ);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cmo.BlobRead(pBlobView);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_cmo.BlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 22;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_OBJ;
			protected void SetDefaults()
			{
				m_cmo = new FtCmoStruct();
				PostSetDefaults();
			}

			protected FtCmoStruct m_cmo;
			protected FtCfStruct m_pictFormat;
			protected FtPioGrbitStruct m_pictFlags;
			public ObjRecord(ushort nIndex, FtCmoStruct.ObjType eType) : base(TYPE, (ushort)(SIZE + (eType == (FtCmoStruct.ObjType.OBJ_TYPE_PICTURE) ? (FtCfStruct.SIZE + FtPioGrbitStruct.SIZE) : 0) + 4))
			{
				SetDefaults();
				nbAssert.Assert(eType == FtCmoStruct.ObjType.OBJ_TYPE_PICTURE || eType == FtCmoStruct.ObjType.OBJ_TYPE_CHART);
				m_cmo.m_ft = 0x15;
				m_cmo.m_cb = 0x12;
				m_cmo.m_ot = (ushort)(eType);
				m_cmo.m_id = nIndex;
				m_cmo.m_fLocked = 1;
				m_cmo.m_fPrint = 1;
				m_cmo.m_unused6 = 1;
				if (eType == FtCmoStruct.ObjType.OBJ_TYPE_PICTURE)
				{
					m_pictFormat = new FtCfStruct();
					m_pictFormat.m_cf = 0xffff;
					m_pictFlags = new FtPioGrbitStruct();
					m_pictFlags.m_fAutoPict = 0x1;
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if ((FtCmoStruct.ObjType)(m_cmo.m_ot) == FtCmoStruct.ObjType.OBJ_TYPE_PICTURE)
				{
					m_pictFormat = new FtCfStruct();
					m_pictFormat.BlobRead(pBlobView);
					m_pictFlags = new FtPioGrbitStruct();
					m_pictFlags.BlobRead(pBlobView);
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if ((FtCmoStruct.ObjType)(m_cmo.m_ot) == FtCmoStruct.ObjType.OBJ_TYPE_PICTURE)
				{
					m_pictFormat.BlobWrite(pBlobView);
					m_pictFlags.BlobWrite(pBlobView);
				}
				pBlobView.PackUint32(0);
			}

			protected void PostSetDefaults()
			{
				m_pictFormat = null;
				m_pictFlags = null;
			}

			public new FtCmoStruct.ObjType GetType()
			{
				return (FtCmoStruct.ObjType)(m_cmo.m_ot);
			}

			~ObjRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class NumberRecord : BiffRecord
		{
			public NumberRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_NUMBER);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cell.BlobRead(pBlobView);
				m_num = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_cell.BlobWrite(pBlobView);
				pBlobView.PackDouble(m_num);
			}

			protected const ushort SIZE = 14;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_NUMBER;
			protected void SetDefaults()
			{
				m_cell = new CellStruct();
				m_num = 0.0;
			}

			protected CellStruct m_cell;
			protected double m_num;
			public NumberRecord(ushort nX, ushort nY, ushort nXfIndex, double fNumber) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cell.m_rw.m_rw = nY;
				m_cell.m_col.m_col = nX;
				m_cell.m_ixfe.m_ixfe = nXfIndex;
				m_num = fNumber;
			}

			public ushort GetX()
			{
				return m_cell.m_col.m_col;
			}

			public ushort GetY()
			{
				return m_cell.m_rw.m_rw;
			}

			public ushort GetXfIndex()
			{
				return m_cell.m_ixfe.m_ixfe;
			}

			public double GetNumber()
			{
				return m_num;
			}

			~NumberRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MulRkRecord : BiffRecord
		{
			public MulRkRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MULRK);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw.BlobRead(pBlobView);
				m_col.BlobRead(pBlobView);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rw.BlobWrite(pBlobView);
				m_col.BlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MULRK;
			protected void SetDefaults()
			{
				m_rw = new RwStruct();
				m_col = new ColStruct();
				PostSetDefaults();
			}

			protected RwStruct m_rw;
			protected ColStruct m_col;
			protected OwnedVector<RkRecStruct> m_pRkRecVector;
			protected ushort m_colLast;
			protected void PostSetDefaults()
			{
				m_pRkRecVector = new OwnedVector<RkRecStruct>();
				m_colLast = 0;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				uint nNumColumn = (m_pHeader.m_nSize - SIZE - 2) / RkRecStruct.SIZE;
				for (uint i = 0; i < nNumColumn; i++)
				{
					RkRecStruct pRkRec = new RkRecStruct();
					pRkRec.BlobRead(pBlobView);
					{
						NumberDuck.Secret.RkRecStruct __3209674599 = pRkRec;
						pRkRec = null;
						m_pRkRecVector.PushBack(__3209674599);
					}
				}
				m_colLast = pBlobView.UnpackUint16();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				for (int i = 0; i < m_pRkRecVector.GetSize(); i++)
					m_pRkRecVector.Get(i).BlobWrite(pBlobView);
				pBlobView.PackUint16(m_colLast);
			}

			public ushort GetX()
			{
				return m_col.m_col;
			}

			public ushort GetY()
			{
				return m_rw.m_rw;
			}

			public ushort GetNumRk()
			{
				return (ushort)(m_pRkRecVector.GetSize());
			}

			public ushort GetXfIndexByIndex(ushort nIndex)
			{
				return m_pRkRecVector.Get(nIndex).m_ixfe;
			}

			public uint GetRkValueByIndex(ushort nIndex)
			{
				return m_pRkRecVector.Get(nIndex).m_RK;
			}

			~MulRkRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MulBlank : BiffRecord
		{
			public MulBlank(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MULBLANK);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw.BlobRead(pBlobView);
				m_col.BlobRead(pBlobView);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rw.BlobWrite(pBlobView);
				m_col.BlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MULBLANK;
			protected void SetDefaults()
			{
				m_rw = new RwStruct();
				m_col = new ColStruct();
				PostSetDefaults();
			}

			protected RwStruct m_rw;
			protected ColStruct m_col;
			protected OwnedVector<IXFCellStruct> m_pIXFCellVector;
			protected ushort m_colLast;
			public MulBlank(ushort nX, ushort nY, Vector<int> pXfIndexVector) : base(TYPE, SIZE + IXFCellStruct.SIZE * (uint)(pXfIndexVector.GetSize()) + 2)
			{
				SetDefaults();
				nbAssert.Assert(pXfIndexVector.GetSize() > 1);
				m_rw.m_rw = nY;
				m_col.m_col = nX;
				for (int i = 0; i < pXfIndexVector.GetSize(); i++)
				{
					IXFCellStruct pIXFCell = new IXFCellStruct();
					pIXFCell.m_ixfe = (ushort)(pXfIndexVector.Get(i));
					{
						NumberDuck.Secret.IXFCellStruct __62881043 = pIXFCell;
						pIXFCell = null;
						m_pIXFCellVector.PushBack(__62881043);
					}
				}
				m_colLast = (ushort)(m_col.m_col + pXfIndexVector.GetSize() - 1);
			}

			protected void PostSetDefaults()
			{
				m_pIXFCellVector = new OwnedVector<IXFCellStruct>();
				m_colLast = 0;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				uint nNumColumn = (m_pHeader.m_nSize - SIZE - 2) / IXFCellStruct.SIZE;
				for (uint i = 0; i < nNumColumn; i++)
				{
					IXFCellStruct pIXFCell = new IXFCellStruct();
					pIXFCell.BlobRead(pBlobView);
					{
						NumberDuck.Secret.IXFCellStruct __62881043 = pIXFCell;
						pIXFCell = null;
						m_pIXFCellVector.PushBack(__62881043);
					}
				}
				m_colLast = pBlobView.UnpackUint16();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				for (int i = 0; i < m_pIXFCellVector.GetSize(); i++)
					m_pIXFCellVector.Get(i).BlobWrite(pBlobView);
				pBlobView.PackUint16(m_colLast);
			}

			public ushort GetX()
			{
				return m_col.m_col;
			}

			public ushort GetY()
			{
				return m_rw.m_rw;
			}

			public ushort GetNumColumn()
			{
				return (ushort)(m_pIXFCellVector.GetSize());
			}

			public ushort GetXfIndexByIndex(ushort nIndex)
			{
				IXFCellStruct pIXFCell = m_pIXFCellVector.Get(nIndex);
				return pIXFCell.m_ixfe;
			}

			~MulBlank()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MsoDrawingRecord_Position
		{
			public ushort m_nCellX1;
			public ushort m_nSubCellX1;
			public ushort m_nCellY1;
			public ushort m_nSubCellY1;
			public ushort m_nCellX2;
			public ushort m_nSubCellX2;
			public ushort m_nCellY2;
			public ushort m_nSubCellY2;
		}
		class MsoDrawingRecord : BiffRecord
		{
			public MsoDrawingRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MSO_DRAWING);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MSO_DRAWING;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OfficeArtRecord m_pOfficeArtRecord;
			protected MsoDrawingRecord_Position m_pPosition;
			protected Blob m_pBlob;
			public MsoDrawingRecord(OfficeArtDgContainerRecord pOfficeArtDgContainerRecord, uint nWriteSize) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_pBlob = new Blob(true);
				pOfficeArtDgContainerRecord.RecursiveWrite(m_pBlob.GetBlobView());
				m_pBlob.GetBlobView().SetOffset(0);
				m_pHeader.m_nSize = nWriteSize;
				m_pBlob.Resize((int)(nWriteSize), true);
			}

			public MsoDrawingRecord(OfficeArtSpContainerRecord pOfficeArtSpContainerRecord) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_pBlob = new Blob(true);
				pOfficeArtSpContainerRecord.RecursiveWrite(m_pBlob.GetBlobView());
				m_pHeader.m_nSize = (uint)(m_pBlob.GetSize());
			}

			protected void PostSetDefaults()
			{
				m_pOfficeArtRecord = null;
				m_pBlob = null;
				m_pPosition = new MsoDrawingRecord_Position();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_pOfficeArtRecord == null);
				m_pOfficeArtRecord = OfficeArtRecord.CreateOfficeArtRecord(pBlobView);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_pBlob.GetBlobView().SetOffset(0);
				pBlobView.Pack(m_pBlob.GetBlobView(), m_pBlob.GetSize());
			}

			public MsoDrawingRecord_Position GetPosition()
			{
				if (m_pOfficeArtRecord != null)
				{
					if (m_pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_DG_CONTAINER || m_pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_SP_CONTAINER)
					{
						OfficeArtClientAnchorSheetRecord pOfficeArtClientAnchorSheetRecord = (OfficeArtClientAnchorSheetRecord)(m_pOfficeArtRecord.FindOfficeArtRecordByType(OfficeArtRecord.Type.TYPE_OFFICE_ART_CLIENT_ANCHOR_SHEET));
						if (pOfficeArtClientAnchorSheetRecord != null)
						{
							m_pPosition.m_nCellX1 = pOfficeArtClientAnchorSheetRecord.GetCellX1();
							m_pPosition.m_nSubCellX1 = pOfficeArtClientAnchorSheetRecord.GetSubCellX1();
							m_pPosition.m_nCellY1 = pOfficeArtClientAnchorSheetRecord.GetCellY1();
							m_pPosition.m_nSubCellY1 = pOfficeArtClientAnchorSheetRecord.GetSubCellY1();
							m_pPosition.m_nCellX2 = pOfficeArtClientAnchorSheetRecord.GetCellX2();
							m_pPosition.m_nSubCellX2 = pOfficeArtClientAnchorSheetRecord.GetSubCellX2();
							m_pPosition.m_nCellY2 = pOfficeArtClientAnchorSheetRecord.GetCellY2();
							m_pPosition.m_nSubCellY2 = pOfficeArtClientAnchorSheetRecord.GetSubCellY2();
							return m_pPosition;
						}
					}
				}
				return null;
			}

			public OfficeArtFOPTEStruct GetProperty(OfficeArtRecord.OPIDType eType)
			{
				if (m_pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_DG_CONTAINER || m_pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_SP_CONTAINER)
				{
					OfficeArtFOPTRecord pOfficeArtFOPTRecord = (OfficeArtFOPTRecord)(m_pOfficeArtRecord.FindOfficeArtRecordByType(OfficeArtRecord.Type.TYPE_OFFICE_ART_FOPT));
					if (pOfficeArtFOPTRecord != null)
						return pOfficeArtFOPTRecord.GetProperty(eType);
				}
				return null;
			}

			~MsoDrawingRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MsoDrawingGroupRecord : BiffRecord
		{
			public MsoDrawingGroupRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MSO_DRAWING_GROUP);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MSO_DRAWING_GROUP;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OfficeArtDggContainerRecord m_pOfficeArtDggContainerRecord;
			public MsoDrawingGroupRecord(Vector<Picture> pPictureVector) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_pOfficeArtDggContainerRecord = new OfficeArtDggContainerRecord();
				m_pOfficeArtDggContainerRecord.AddOfficeArtRecord(new OfficeArtFDGGBlockRecord((uint)(pPictureVector.GetSize())));
				OfficeArtBStoreContainerRecord pOfficeArtBStoreContainerRecord = new OfficeArtBStoreContainerRecord();
				for (int i = 0; i < pPictureVector.GetSize(); i++)
					pOfficeArtBStoreContainerRecord.AddOfficeArtRecord(new OfficeArtFBSERecord(pPictureVector.Get(i)));
				{
					NumberDuck.Secret.OfficeArtBStoreContainerRecord __3451512242 = pOfficeArtBStoreContainerRecord;
					pOfficeArtBStoreContainerRecord = null;
					m_pOfficeArtDggContainerRecord.AddOfficeArtRecord(__3451512242);
				}
				OfficeArtFOPTRecord pOfficeArtFOPTRecord = new OfficeArtFOPTRecord();
				pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_TEXT_BOOLEAN_PROPERTIES), 0, 524296);
				pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_COLOR), 0, 134217793);
				pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_LINE_COLOR), 0, 134217792);
				{
					NumberDuck.Secret.OfficeArtFOPTRecord __1214438724 = pOfficeArtFOPTRecord;
					pOfficeArtFOPTRecord = null;
					m_pOfficeArtDggContainerRecord.AddOfficeArtRecord(__1214438724);
				}
				m_pOfficeArtDggContainerRecord.AddOfficeArtRecord(new OfficeArtSplitMenuColorContainerRecord());
				m_pHeader.m_nSize = m_pOfficeArtDggContainerRecord.GetRecursiveSize();
				uint nOffset = MAX_DATA_SIZE;
				while (nOffset < m_pHeader.m_nSize)
				{
					m_pContinueInfoVector.PushBack(new BiffRecord_ContinueInfo((int)(nOffset), 0));
					nOffset += MAX_DATA_SIZE;
				}
			}

			protected void PostSetDefaults()
			{
				m_pOfficeArtDggContainerRecord = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_pOfficeArtDggContainerRecord == null);
				OfficeArtRecord pOfficeArtRecord = OfficeArtRecord.CreateOfficeArtRecord(pBlobView);
				nbAssert.Assert(pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_DGG_CONTAINER);
				{
					NumberDuck.Secret.OfficeArtRecord __3533451309 = pOfficeArtRecord;
					pOfficeArtRecord = null;
					m_pOfficeArtDggContainerRecord = (OfficeArtDggContainerRecord)(__3533451309);
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_pOfficeArtDggContainerRecord.RecursiveWrite(pBlobView);
			}

			public OfficeArtDggContainerRecord GetOfficeArtDggContainerRecord()
			{
				return m_pOfficeArtDggContainerRecord;
			}

			~MsoDrawingGroupRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Mms : BiffRecord
		{
			public Mms() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Mms(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MMS);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_reserved1 = pBlobView.UnpackUint8();
				m_reserved2 = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_reserved1);
				pBlobView.PackUint8(m_reserved2);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MMS;
			protected void SetDefaults()
			{
				m_reserved1 = 0;
				m_reserved2 = 0;
			}

			protected byte m_reserved1;
			protected byte m_reserved2;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MergeCellsRecord : BiffRecord
		{
			public MergeCellsRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MergeCells);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cmcs = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_cmcs);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MergeCells;
			protected void SetDefaults()
			{
				m_cmcs = 0;
				PostSetDefaults();
			}

			protected ushort m_cmcs;
			protected OwnedVector<Ref8Struct> m_pRef8Vector;
			public MergeCellsRecord(OwnedVector<MergedCell> pMergedCellVector, int nOffset, int nSize) : base(TYPE, SIZE + Ref8Struct.SIZE * (uint)(nSize))
			{
				SetDefaults();
				m_cmcs = (ushort)(nSize);
				for (ushort i = 0; i < m_cmcs; i++)
				{
					MergedCell pMergedCell = pMergedCellVector.Get(nOffset + i);
					Ref8Struct pRef8 = new Ref8Struct();
					pRef8.m_rwFirst = (ushort)(pMergedCell.GetY());
					uint nTemp = (uint)(pRef8.m_rwFirst) + (uint)(pMergedCell.GetHeight()) - 1;
					if (nTemp > 0xFFFF)
						nTemp = 0xFFFF;
					pRef8.m_rwLast = (ushort)(nTemp);
					pRef8.m_colFirst = (ushort)(pMergedCell.GetX());
					nTemp = (uint)(pRef8.m_colFirst) + (uint)(pMergedCell.GetWidth()) - 1;
					if (nTemp > 0xFF)
						nTemp = 0xFF;
					pRef8.m_colLast = (ushort)(nTemp);
					{
						NumberDuck.Secret.Ref8Struct __2933356801 = pRef8;
						pRef8 = null;
						m_pRef8Vector.PushBack(__2933356801);
					}
				}
			}

			protected void PostSetDefaults()
			{
				m_pRef8Vector = new OwnedVector<Ref8Struct>();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_cmcs; i++)
				{
					Ref8Struct pRef8 = new Ref8Struct();
					pRef8.BlobRead(pBlobView);
					{
						NumberDuck.Secret.Ref8Struct __2933356801 = pRef8;
						pRef8 = null;
						m_pRef8Vector.PushBack(__2933356801);
					}
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_cmcs; i++)
					m_pRef8Vector.Get(i).BlobWrite(pBlobView);
			}

			public ushort GetNumMergedCell()
			{
				return m_cmcs;
			}

			public Ref8Struct GetMergedCell(ushort nIndex)
			{
				nbAssert.Assert(nIndex < m_cmcs);
				return m_pRef8Vector.Get(nIndex);
			}

			~MergeCellsRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MarkerFormatRecord : BiffRecord
		{
			public MarkerFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_MarkerFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rgbFore = pBlobView.UnpackUint32();
				m_rgbBack = pBlobView.UnpackUint32();
				m_imk = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAuto = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x7);
				m_fNotShowInt = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fNotShowBrd = (ushort)((nBitmask0 >> 5) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 6) & 0x3ff);
				m_icvFore = pBlobView.UnpackUint16();
				m_icvBack = pBlobView.UnpackUint16();
				m_miSize = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_rgbFore);
				pBlobView.PackUint32(m_rgbBack);
				pBlobView.PackUint16(m_imk);
				int nBitmask0 = 0;
				nBitmask0 += m_fAuto << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fNotShowInt << 4;
				nBitmask0 += m_fNotShowBrd << 5;
				nBitmask0 += m_reserved2 << 6;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icvFore);
				pBlobView.PackUint16(m_icvBack);
				pBlobView.PackUint32(m_miSize);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_MarkerFormat;
			protected void SetDefaults()
			{
				m_rgbFore = 0;
				m_rgbBack = 0;
				m_imk = 0;
				m_fAuto = 0;
				m_reserved1 = 0;
				m_fNotShowInt = 0;
				m_fNotShowBrd = 0;
				m_reserved2 = 0;
				m_icvFore = 0;
				m_icvBack = 0;
				m_miSize = 0;
			}

			protected uint m_rgbFore;
			protected uint m_rgbBack;
			protected ushort m_imk;
			protected ushort m_fAuto;
			protected ushort m_reserved1;
			protected ushort m_fNotShowInt;
			protected ushort m_fNotShowBrd;
			protected ushort m_reserved2;
			protected ushort m_icvFore;
			protected ushort m_icvBack;
			protected uint m_miSize;
			public MarkerFormatRecord(Marker pMarker) : base(TYPE, SIZE)
			{
				SetDefaults();
				Color pColor;
				switch (pMarker.GetType())
				{
					case Marker.Type.TYPE_NONE:
					{
						m_imk = 0x0000;
						break;
					}

					case Marker.Type.TYPE_SQUARE:
					{
						m_imk = 0x0001;
						break;
					}

					case Marker.Type.TYPE_DIAMOND:
					{
						m_imk = 0x0002;
						break;
					}

					case Marker.Type.TYPE_TRIANGLE:
					{
						m_imk = 0x0003;
						break;
					}

					case Marker.Type.TYPE_X:
					{
						m_imk = 0x0004;
						break;
					}

					case Marker.Type.TYPE_ASTERISK:
					{
						m_imk = 0x0005;
						break;
					}

					case Marker.Type.TYPE_SHORT_BAR:
					{
						m_imk = 0x0006;
						break;
					}

					case Marker.Type.TYPE_LONG_BAR:
					{
						m_imk = 0x0007;
						break;
					}

					case Marker.Type.TYPE_CIRCULAR:
					{
						m_imk = 0x0008;
						break;
					}

					case Marker.Type.TYPE_PLUS:
					{
						m_imk = 0x0009;
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
				pColor = pMarker.GetBorderColor(false);
				if (pColor != null)
				{
					m_icvFore = BiffWorkbookGlobals.SnapToPalette(pColor);
					m_rgbFore = BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(m_icvFore) & 0xFFFFFF;
				}
				else
				{
					m_fNotShowBrd = 0x1;
					m_icvFore = 0x004D;
				}
				pColor = pMarker.GetFillColor(false);
				if (pColor != null)
				{
					m_icvBack = BiffWorkbookGlobals.SnapToPalette(pColor);
					m_rgbBack = BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(m_icvBack) & 0xFFFFFF;
				}
				else
				{
					m_fNotShowInt = 0x1;
					m_icvBack = 0x004D;
				}
				m_miSize = WorksheetImplementation.PixelsToTwips((ushort)(pMarker.GetSize()));
			}

			public void ModifyMarker(Marker pMarker, BiffWorkbookGlobals pBiffWorkbookGlobals)
			{
				if (m_fAuto == 0x1)
					return;
				switch (m_imk)
				{
					case 0x0001:
					{
						pMarker.SetType(Marker.Type.TYPE_SQUARE);
						break;
					}

					case 0x0002:
					{
						pMarker.SetType(Marker.Type.TYPE_DIAMOND);
						break;
					}

					case 0x0003:
					{
						pMarker.SetType(Marker.Type.TYPE_TRIANGLE);
						break;
					}

					case 0x0004:
					{
						pMarker.SetType(Marker.Type.TYPE_X);
						break;
					}

					case 0x0005:
					{
						pMarker.SetType(Marker.Type.TYPE_ASTERISK);
						break;
					}

					case 0x0006:
					{
						pMarker.SetType(Marker.Type.TYPE_SHORT_BAR);
						break;
					}

					case 0x0007:
					{
						pMarker.SetType(Marker.Type.TYPE_LONG_BAR);
						break;
					}

					case 0x0008:
					{
						pMarker.SetType(Marker.Type.TYPE_CIRCULAR);
						break;
					}

					case 0x0009:
					{
						pMarker.SetType(Marker.Type.TYPE_PLUS);
						break;
					}

					default:
					{
						pMarker.SetType(Marker.Type.TYPE_NONE);
						break;
					}

				}
				if (m_fNotShowBrd == 0x0)
					pMarker.GetBorderColor(true).Set((byte)(m_rgbFore & 0xFF), (byte)((m_rgbFore >> 8) & 0xFF), (byte)((m_rgbFore >> 16) & 0xFF));
				else
					pMarker.ClearBorderColor();
				if (m_fNotShowInt == 0x0)
					pMarker.GetFillColor(true).Set((byte)(m_rgbBack & 0xFF), (byte)((m_rgbBack >> 8) & 0xFF), (byte)((m_rgbBack >> 16) & 0xFF));
				else
					pMarker.ClearFillColor();
				pMarker.SetSize(WorksheetImplementation.TwipsToPixels((ushort)(m_miSize)));
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineRecord : BiffRecord
		{
			public LineRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Line);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fStacked = (ushort)((nBitmask0 >> 0) & 0x1);
				m_f100 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fHasShadow = (ushort)((nBitmask0 >> 2) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 3) & 0x1fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fStacked << 0;
				nBitmask0 += m_f100 << 1;
				nBitmask0 += m_fHasShadow << 2;
				nBitmask0 += m_reserved << 3;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Line;
			protected void SetDefaults()
			{
				m_fStacked = 0;
				m_f100 = 0;
				m_fHasShadow = 0;
				m_reserved = 0;
			}

			protected ushort m_fStacked;
			protected ushort m_f100;
			protected ushort m_fHasShadow;
			protected ushort m_reserved;
			public LineRecord(Chart.Type eType) : base(TYPE, SIZE)
			{
				SetDefaults();
				nbAssert.Assert(eType == Chart.Type.TYPE_LINE || eType == Chart.Type.TYPE_LINE_STACKED || eType == Chart.Type.TYPE_LINE_STACKED_100);
				if (eType == Chart.Type.TYPE_LINE_STACKED || eType == Chart.Type.TYPE_LINE_STACKED_100)
					m_fStacked = 0x1;
				if (eType == Chart.Type.TYPE_LINE_STACKED_100)
					m_f100 = 0x1;
			}

			public Chart.Type GetChartType()
			{
				if (m_fStacked == 0x1)
					if (m_f100 == 0x1)
						return Chart.Type.TYPE_LINE_STACKED_100;
					else
						return Chart.Type.TYPE_LINE_STACKED;
				return Chart.Type.TYPE_LINE;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineFormatRecord : BiffRecord
		{
			public LineFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_LineFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rgb = pBlobView.UnpackUint32();
				m_lns = pBlobView.UnpackUint16();
				m_we = pBlobView.UnpackInt16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAuto = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fAxisOn = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fAutoCo = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 4) & 0xfff);
				m_icv = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_rgb);
				pBlobView.PackUint16(m_lns);
				pBlobView.PackInt16(m_we);
				int nBitmask0 = 0;
				nBitmask0 += m_fAuto << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fAxisOn << 2;
				nBitmask0 += m_fAutoCo << 3;
				nBitmask0 += m_reserved2 << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icv);
			}

			protected const ushort SIZE = 12;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_LineFormat;
			protected void SetDefaults()
			{
				m_rgb = 0;
				m_lns = 0;
				m_we = 0;
				m_fAuto = 0;
				m_reserved1 = 0;
				m_fAxisOn = 0;
				m_fAutoCo = 0;
				m_reserved2 = 0;
				m_icv = 0;
			}

			protected uint m_rgb;
			protected ushort m_lns;
			protected short m_we;
			protected ushort m_fAuto;
			protected ushort m_reserved1;
			protected ushort m_fAxisOn;
			protected ushort m_fAutoCo;
			protected ushort m_reserved2;
			protected ushort m_icv;
			public LineFormatRecord(uint rgb, ushort lns, short we, bool fAuto, bool fAxisOn, bool fAutoCo, ushort icv) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_rgb = rgb;
				m_lns = lns;
				m_we = we;
				m_fAuto = fAuto ? (ushort)(0x1) : (ushort)(0x0);
				m_fAxisOn = fAxisOn ? (ushort)(0x1) : (ushort)(0x0);
				m_fAutoCo = fAutoCo ? (ushort)(0x1) : (ushort)(0x0);
				m_icv = icv;
			}

			public LineFormatRecord(Line pLine, bool bAxis) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_icv = BiffWorkbookGlobals.SnapToPalette(pLine.GetColor());
				m_rgb = BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(m_icv) & 0xFFFFFF;
				switch (pLine.GetType())
				{
					case Line.Type.TYPE_NONE:
					{
						m_lns = 0x0005;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_THIN:
					{
						m_lns = 0x0000;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_DASHED:
					{
						m_lns = 0x0001;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_DOTTED:
					{
						m_lns = 0x0002;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_DASH_DOT:
					{
						m_lns = 0x0003;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_DASH_DOT_DOT:
					{
						m_lns = 0x0004;
						m_we = -1;
						break;
					}

					case Line.Type.TYPE_MEDIUM:
					{
						m_lns = 0x0000;
						m_we = 0x0001;
						break;
					}

					case Line.Type.TYPE_MEDIUM_DASHED:
					{
						m_lns = 0x0001;
						m_we = 0x0001;
						break;
					}

					case Line.Type.TYPE_MEDIUM_DASH_DOT:
					{
						m_lns = 0x0003;
						m_we = 0x0001;
						break;
					}

					case Line.Type.TYPE_MEDIUM_DASH_DOT_DOT:
					{
						m_lns = 0x0004;
						m_we = 0x0001;
						break;
					}

					case Line.Type.TYPE_THICK:
					{
						m_lns = 0x0000;
						m_we = 0x0002;
						break;
					}

					default:
					{
						break;
					}

				}
				if (bAxis)
					m_fAxisOn = 0x1;
			}

			public void ModifyLine(Line pLine, BiffWorkbookGlobals pBiffWorkbookGlobals)
			{
				if (m_fAuto == 0x1)
					return;
				if (m_lns == 0x0005)
				{
					pLine.SetType(Line.Type.TYPE_NONE);
				}
				else
				{
					switch (m_we)
					{
						case 0x0002:
						{
							pLine.SetType(Line.Type.TYPE_THICK);
							break;
						}

						case 0x0001:
						{
							switch (m_lns)
							{
								case 0x0000:
								{
									pLine.SetType(Line.Type.TYPE_MEDIUM);
									break;
								}

								case 0x0001:
								{
									pLine.SetType(Line.Type.TYPE_MEDIUM_DASHED);
									break;
								}

								case 0x0003:
								{
									pLine.SetType(Line.Type.TYPE_MEDIUM_DASH_DOT);
									break;
								}

								case 0x0004:
								{
									pLine.SetType(Line.Type.TYPE_MEDIUM_DASH_DOT_DOT);
									break;
								}

								default:
								{
									pLine.SetType(Line.Type.TYPE_MEDIUM);
									break;
								}

							}
							break;
						}

						case 0x0000:
						case -1:
						{
							switch (m_lns)
							{
								case 0x0000:
								{
									pLine.SetType(Line.Type.TYPE_THIN);
									break;
								}

								case 0x0001:
								{
									pLine.SetType(Line.Type.TYPE_DASHED);
									break;
								}

								case 0x0002:
								{
									pLine.SetType(Line.Type.TYPE_DOTTED);
									break;
								}

								case 0x0003:
								{
									pLine.SetType(Line.Type.TYPE_DASH_DOT);
									break;
								}

								case 0x0004:
								{
									pLine.SetType(Line.Type.TYPE_DASH_DOT_DOT);
									break;
								}

								default:
								{
									pLine.SetType(Line.Type.TYPE_THIN);
									break;
								}

							}
							break;
						}

					}
				}
				if (m_icv != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_FOREGROUND && m_icv != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_BACKGROUND)
					pLine.GetColor().SetFromRgba(pBiffWorkbookGlobals.GetPaletteColorByIndex(m_icv));
			}

			public void PackForChecksum(BlobView pBlobView)
			{
				pBlobView.PackUint8((byte)(m_icv));
				pBlobView.PackUint8((byte)(m_lns));
				pBlobView.PackUint8((byte)(m_we));
				int nBitMask = 0;
				nBitMask = nBitMask | m_fAuto << 0;
				pBlobView.PackUint8((byte)(nBitMask));
				pBlobView.PackUint32(m_rgb);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LegendRecord : BiffRecord
		{
			public LegendRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Legend);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_x = pBlobView.UnpackUint32();
				m_y = pBlobView.UnpackUint32();
				m_dx = pBlobView.UnpackUint32();
				m_dy = pBlobView.UnpackUint32();
				m_unused = pBlobView.UnpackUint8();
				m_wSpace = pBlobView.UnpackUint8();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoPosition = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fAutoPosX = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fAutoPosY = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fVert = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fWasDataTable = (ushort)((nBitmask0 >> 5) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 6) & 0x3ff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_x);
				pBlobView.PackUint32(m_y);
				pBlobView.PackUint32(m_dx);
				pBlobView.PackUint32(m_dy);
				pBlobView.PackUint8(m_unused);
				pBlobView.PackUint8(m_wSpace);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoPosition << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fAutoPosX << 2;
				nBitmask0 += m_fAutoPosY << 3;
				nBitmask0 += m_fVert << 4;
				nBitmask0 += m_fWasDataTable << 5;
				nBitmask0 += m_reserved2 << 6;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Legend;
			protected void SetDefaults()
			{
				m_x = 0;
				m_y = 0;
				m_dx = 0;
				m_dy = 0;
				m_unused = 0x03;
				m_wSpace = 0x01;
				m_fAutoPosition = 1;
				m_reserved1 = 1;
				m_fAutoPosX = 1;
				m_fAutoPosY = 1;
				m_fVert = 1;
				m_fWasDataTable = 0;
				m_reserved2 = 0;
			}

			protected uint m_x;
			protected uint m_y;
			protected uint m_dx;
			protected uint m_dy;
			protected byte m_unused;
			protected byte m_wSpace;
			protected ushort m_fAutoPosition;
			protected ushort m_reserved1;
			protected ushort m_fAutoPosX;
			protected ushort m_fAutoPosY;
			protected ushort m_fVert;
			protected ushort m_fWasDataTable;
			protected ushort m_reserved2;
			public LegendRecord(Legend pLegend) : base(TYPE, SIZE)
			{
				SetDefaults();
				nbAssert.Assert(!pLegend.GetHidden());
			}

			public void ModifyLegend(Legend pLegend, BiffWorkbookGlobals pBiffWorkbookGlobals)
			{
				pLegend.SetHidden(false);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LeftMarginRecord : BiffRecord
		{
			public LeftMarginRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_LEFT_MARGIN);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_num = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackDouble(m_num);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_LEFT_MARGIN;
			protected void SetDefaults()
			{
				m_num = 0;
			}

			protected double m_num;
			public LeftMarginRecord(double num) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_num = num;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LabelSstRecord : BiffRecord
		{
			public LabelSstRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_LABELSST);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_nRowIndex = pBlobView.UnpackUint16();
				m_nColumnIndex = pBlobView.UnpackUint16();
				m_nXfIndex = pBlobView.UnpackUint16();
				m_nSstIndex = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_nRowIndex);
				pBlobView.PackUint16(m_nColumnIndex);
				pBlobView.PackUint16(m_nXfIndex);
				pBlobView.PackUint32(m_nSstIndex);
			}

			protected const ushort SIZE = 10;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_LABELSST;
			protected void SetDefaults()
			{
				m_nRowIndex = 0;
				m_nColumnIndex = 0;
				m_nXfIndex = 0;
				m_nSstIndex = 0;
			}

			protected ushort m_nRowIndex;
			protected ushort m_nColumnIndex;
			protected ushort m_nXfIndex;
			protected uint m_nSstIndex;
			public LabelSstRecord(ushort nX, ushort nY, ushort nXfIndex, uint nSstIndex) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_nRowIndex = nY;
				m_nColumnIndex = nX;
				m_nXfIndex = nXfIndex;
				m_nSstIndex = nSstIndex;
			}

			public ushort GetX()
			{
				return m_nColumnIndex;
			}

			public ushort GetY()
			{
				return m_nRowIndex;
			}

			public ushort GetXfIndex()
			{
				return m_nXfIndex;
			}

			public uint GetSstIndex()
			{
				return m_nSstIndex;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class InterfaceHdr : BiffRecord
		{
			public InterfaceHdr() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public InterfaceHdr(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_INTERFACE_HDR);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_codePage = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_codePage);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_INTERFACE_HDR;
			protected void SetDefaults()
			{
				m_codePage = 0x04B0;
			}

			protected ushort m_codePage;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class InterfaceEnd : BiffRecord
		{
			public InterfaceEnd() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public InterfaceEnd(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_INTERFACE_END);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_INTERFACE_END;
			protected void SetDefaults()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HideObj : BiffRecord
		{
			public HideObj() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public HideObj(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_HIDE_OBJ);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_hideObj = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_hideObj);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_HIDE_OBJ;
			protected void SetDefaults()
			{
				m_hideObj = 0;
			}

			protected ushort m_hideObj;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HeaderFooterRecord : BiffRecord
		{
			public HeaderFooterRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public HeaderFooterRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_HeaderFooter);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_guidSView_0 = pBlobView.UnpackUint8();
				m_guidSView_1 = pBlobView.UnpackUint8();
				m_guidSView_2 = pBlobView.UnpackUint8();
				m_guidSView_3 = pBlobView.UnpackUint8();
				m_guidSView_4 = pBlobView.UnpackUint8();
				m_guidSView_5 = pBlobView.UnpackUint8();
				m_guidSView_6 = pBlobView.UnpackUint8();
				m_guidSView_7 = pBlobView.UnpackUint8();
				m_guidSView_8 = pBlobView.UnpackUint8();
				m_guidSView_9 = pBlobView.UnpackUint8();
				m_guidSView_10 = pBlobView.UnpackUint8();
				m_guidSView_11 = pBlobView.UnpackUint8();
				m_guidSView_12 = pBlobView.UnpackUint8();
				m_guidSView_13 = pBlobView.UnpackUint8();
				m_guidSView_14 = pBlobView.UnpackUint8();
				m_guidSView_15 = pBlobView.UnpackUint8();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fHFDiffOddEven = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fHFDiffFirst = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fHFScaleWithDoc = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fHFAlignMargins = (ushort)((nBitmask0 >> 3) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 4) & 0xfff);
				m_cchHeaderEven = pBlobView.UnpackUint16();
				m_cchFooterEven = pBlobView.UnpackUint16();
				m_cchHeaderFirst = pBlobView.UnpackUint16();
				m_cchFooterFirst = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint8(m_guidSView_0);
				pBlobView.PackUint8(m_guidSView_1);
				pBlobView.PackUint8(m_guidSView_2);
				pBlobView.PackUint8(m_guidSView_3);
				pBlobView.PackUint8(m_guidSView_4);
				pBlobView.PackUint8(m_guidSView_5);
				pBlobView.PackUint8(m_guidSView_6);
				pBlobView.PackUint8(m_guidSView_7);
				pBlobView.PackUint8(m_guidSView_8);
				pBlobView.PackUint8(m_guidSView_9);
				pBlobView.PackUint8(m_guidSView_10);
				pBlobView.PackUint8(m_guidSView_11);
				pBlobView.PackUint8(m_guidSView_12);
				pBlobView.PackUint8(m_guidSView_13);
				pBlobView.PackUint8(m_guidSView_14);
				pBlobView.PackUint8(m_guidSView_15);
				int nBitmask0 = 0;
				nBitmask0 += m_fHFDiffOddEven << 0;
				nBitmask0 += m_fHFDiffFirst << 1;
				nBitmask0 += m_fHFScaleWithDoc << 2;
				nBitmask0 += m_fHFAlignMargins << 3;
				nBitmask0 += m_unused << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_cchHeaderEven);
				pBlobView.PackUint16(m_cchFooterEven);
				pBlobView.PackUint16(m_cchHeaderFirst);
				pBlobView.PackUint16(m_cchFooterFirst);
			}

			protected const ushort SIZE = 38;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_HeaderFooter;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_guidSView_0 = 0;
				m_guidSView_1 = 0;
				m_guidSView_2 = 0;
				m_guidSView_3 = 0;
				m_guidSView_4 = 0;
				m_guidSView_5 = 0;
				m_guidSView_6 = 0;
				m_guidSView_7 = 0;
				m_guidSView_8 = 0;
				m_guidSView_9 = 0;
				m_guidSView_10 = 0;
				m_guidSView_11 = 0;
				m_guidSView_12 = 0;
				m_guidSView_13 = 0;
				m_guidSView_14 = 0;
				m_guidSView_15 = 0;
				m_fHFDiffOddEven = 0;
				m_fHFDiffFirst = 0;
				m_fHFScaleWithDoc = 0x1;
				m_fHFAlignMargins = 0x1;
				m_unused = 819;
				m_cchHeaderEven = 0;
				m_cchFooterEven = 0;
				m_cchHeaderFirst = 0;
				m_cchFooterFirst = 0;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected byte m_guidSView_0;
			protected byte m_guidSView_1;
			protected byte m_guidSView_2;
			protected byte m_guidSView_3;
			protected byte m_guidSView_4;
			protected byte m_guidSView_5;
			protected byte m_guidSView_6;
			protected byte m_guidSView_7;
			protected byte m_guidSView_8;
			protected byte m_guidSView_9;
			protected byte m_guidSView_10;
			protected byte m_guidSView_11;
			protected byte m_guidSView_12;
			protected byte m_guidSView_13;
			protected byte m_guidSView_14;
			protected byte m_guidSView_15;
			protected ushort m_fHFDiffOddEven;
			protected ushort m_fHFDiffFirst;
			protected ushort m_fHFScaleWithDoc;
			protected ushort m_fHFAlignMargins;
			protected ushort m_unused;
			protected ushort m_cchHeaderEven;
			protected ushort m_cchFooterEven;
			protected ushort m_cchHeaderFirst;
			protected ushort m_cchFooterFirst;
			protected void PostSetDefaults()
			{
				m_frtHeader.m_rt = 2204;
			}

			~HeaderFooterRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HCenterRecord : BiffRecord
		{
			public HCenterRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_HCENTER);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_hcenter = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_hcenter);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_HCENTER;
			protected void SetDefaults()
			{
				m_hcenter = 0x0000;
			}

			protected ushort m_hcenter;
			public HCenterRecord(bool bCenter) : base(TYPE, SIZE)
			{
				SetDefaults();
				if (bCenter)
					m_hcenter = 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class GelFrameRecord : BiffRecord
		{
			public GelFrameRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_GelFrame);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_GelFrame;
			protected void SetDefaults()
			{
				PostSetDefaults();
			}

			protected OfficeArtFOPTRecord m_OPT1;
			protected OfficeArtTertiaryFOPTRecord m_OPT2;
			public GelFrameRecord(Fill pFill, WorkbookGlobals pWorkbookGlobals) : base(TYPE, SIZE)
			{
				SetDefaults();
				nbAssert.Assert(pFill != null);
				nbAssert.Assert(pWorkbookGlobals != null);
				m_OPT1 = new OfficeArtFOPTRecord();
				{
					Color pColor = pFill.GetForegroundColor();
					byte nR = pColor.GetRed();
					byte nG = pColor.GetGreen();
					byte nB = pColor.GetBlue();
					m_OPT1.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_COLOR), 0, nR << 0 | nG << 8 | nB << 16);
					m_OPT1.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_OPACITY), 0, 65536);
				}
				{
					Color pColor = pFill.GetBackgroundColor();
					byte nR = pColor.GetRed();
					byte nG = pColor.GetGreen();
					byte nB = pColor.GetBlue();
					m_OPT1.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_BACK_COLOR), 0, nR << 0 | nG << 8 | nB << 16);
					m_OPT1.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_BACK_OPACITY), 0, 65536);
				}
				m_OPT2 = new OfficeArtTertiaryFOPTRecord();
				m_pHeader.m_nSize = m_OPT1.GetSize() + m_OPT2.GetSize();
			}

			protected void PostSetDefaults()
			{
				m_OPT1 = null;
				m_OPT2 = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				OfficeArtRecord pOfficeArtRecord;
				nbAssert.Assert(m_OPT1 == null);
				pOfficeArtRecord = OfficeArtRecord.CreateOfficeArtRecord(pBlobView);
				nbAssert.Assert(pOfficeArtRecord != null);
				nbAssert.Assert(pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_FOPT);
				{
					NumberDuck.Secret.OfficeArtRecord __3533451309 = pOfficeArtRecord;
					pOfficeArtRecord = null;
					m_OPT1 = (OfficeArtFOPTRecord)(__3533451309);
				}
				nbAssert.Assert(m_OPT2 == null);
				pOfficeArtRecord = OfficeArtRecord.CreateOfficeArtRecord(pBlobView);
				nbAssert.Assert(pOfficeArtRecord != null);
				nbAssert.Assert(pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_TERTIARY_FOPT);
				{
					NumberDuck.Secret.OfficeArtRecord __3533451309 = pOfficeArtRecord;
					pOfficeArtRecord = null;
					m_OPT2 = (OfficeArtTertiaryFOPTRecord)(__3533451309);
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_OPT1.BlobWrite(pBlobView);
				m_OPT2.BlobWrite(pBlobView);
			}

			public void ModifyFill(Fill pFill, WorkbookGlobals pWorkbookGlobals)
			{
				OfficeArtFOPTEStruct pProperty;
				pProperty = m_OPT1.GetProperty(OfficeArtRecord.OPIDType.OPID_FILL_COLOR);
				if (pProperty != null)
					pFill.GetForegroundColor().Set((byte)(pProperty.m_op & 0xFF), (byte)((pProperty.m_op >> 8) & 0xFF), (byte)((pProperty.m_op >> 16) & 0xFF));
				pProperty = m_OPT1.GetProperty(OfficeArtRecord.OPIDType.OPID_FILL_BACK_COLOR);
				if (pProperty != null)
					pFill.GetBackgroundColor().Set((byte)(pProperty.m_op & 0xFF), (byte)((pProperty.m_op >> 8) & 0xFF), (byte)((pProperty.m_op >> 16) & 0xFF));
			}

			public void PackForChecksum(BlobView pBlobView)
			{
			}

			~GelFrameRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrameRecord : BiffRecord
		{
			public FrameRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Frame);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frt = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoSize = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fAutoPosition = (ushort)((nBitmask0 >> 1) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 2) & 0x3fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_frt);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoSize << 0;
				nBitmask0 += m_fAutoPosition << 1;
				nBitmask0 += m_reserved << 2;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 4;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Frame;
			protected void SetDefaults()
			{
				m_frt = 0x0000;
				m_fAutoSize = 0;
				m_fAutoPosition = 0;
				m_reserved = 0;
			}

			protected ushort m_frt;
			protected ushort m_fAutoSize;
			protected ushort m_fAutoPosition;
			protected ushort m_reserved;
			public FrameRecord(bool fAutoSize) : base(TYPE, SIZE)
			{
				SetDefaults();
				if (fAutoSize)
					m_fAutoSize = 0x1;
				m_fAutoPosition = 0x1;
			}

			public bool GetDropShadow()
			{
				return m_frt == 0x0004;
			}

			public bool GetAutoSize()
			{
				return m_fAutoSize == 0x1;
			}

			public bool GetAutoPosition()
			{
				return m_fAutoPosition == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FormulaRecord : BiffRecord
		{
			public FormulaRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_FORMULA);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cell.BlobRead(pBlobView);
				m_val.BlobRead(pBlobView);
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAlwaysCalc = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fFill = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fShrFmla = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fClearErrors = (ushort)((nBitmask0 >> 5) & 0x1);
				m_reserved3 = (ushort)((nBitmask0 >> 6) & 0x3ff);
				m_chn = pBlobView.UnpackUint32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_cell.BlobWrite(pBlobView);
				m_val.BlobWrite(pBlobView);
				int nBitmask0 = 0;
				nBitmask0 += m_fAlwaysCalc << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fFill << 2;
				nBitmask0 += m_fShrFmla << 3;
				nBitmask0 += m_reserved2 << 4;
				nBitmask0 += m_fClearErrors << 5;
				nBitmask0 += m_reserved3 << 6;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint32(m_chn);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_FORMULA;
			protected void SetDefaults()
			{
				m_cell = new CellStruct();
				m_val = new FormulaValueStruct();
				m_fAlwaysCalc = 0;
				m_reserved1 = 0;
				m_fFill = 0;
				m_fShrFmla = 0;
				m_reserved2 = 0;
				m_fClearErrors = 0;
				m_reserved3 = 0;
				m_chn = 0;
				PostSetDefaults();
			}

			protected CellStruct m_cell;
			protected FormulaValueStruct m_val;
			protected ushort m_fAlwaysCalc;
			protected ushort m_reserved1;
			protected ushort m_fFill;
			protected ushort m_fShrFmla;
			protected ushort m_reserved2;
			protected ushort m_fClearErrors;
			protected ushort m_reserved3;
			protected uint m_chn;
			protected CellParsedFormulaStruct m_formula;
			public FormulaRecord(ushort nX, ushort nY, ushort nXfIndex, Formula pFormula, WorkbookGlobals pWorkbookGlobals) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cell.m_col.m_col = nX;
				m_cell.m_rw.m_rw = nY;
				m_cell.m_ixfe.m_ixfe = nXfIndex;
				m_fAlwaysCalc = 0x1;
				nbAssert.Assert(m_formula == null);
				m_formula = new CellParsedFormulaStruct(pFormula, pWorkbookGlobals);
				m_pHeader.m_nSize += (ushort)(m_formula.GetSize());
			}

			protected void PostSetDefaults()
			{
				m_formula = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_formula == null);
				m_formula = new CellParsedFormulaStruct();
				m_formula.BlobRead(pBlobView);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_formula.BlobWrite(pBlobView);
			}

			public ushort GetX()
			{
				return m_cell.m_col.m_col;
			}

			public ushort GetY()
			{
				return m_cell.m_rw.m_rw;
			}

			public ushort GetXfIndex()
			{
				return m_cell.m_ixfe.m_ixfe;
			}

			public Formula GetFormula(WorkbookGlobals pWorkbookGlobals)
			{
				return new Formula(m_formula.m_rgce.m_pParsedExpressionRecordVector, pWorkbookGlobals);
			}

			~FormulaRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Format : BiffRecord
		{
			public Format(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_FORMAT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ifmt = pBlobView.UnpackUint16();
				m_stFormat.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ifmt);
				m_stFormat.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 5;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_FORMAT;
			protected void SetDefaults()
			{
				m_ifmt = 0;
				m_stFormat = new XLUnicodeStringStruct();
			}

			protected ushort m_ifmt;
			protected XLUnicodeStringStruct m_stFormat;
			public Format(ushort ifmt, string szFormat) : base(TYPE, SIZE)
			{
				SetDefaults();
				nbAssert.Assert(ifmt >= 5 && ifmt <= 8 || ifmt >= 23 && ifmt <= 26 || ifmt >= 41 && ifmt <= 44 || ifmt >= 63 && ifmt <= 66 || ifmt >= 164 && ifmt <= 382 || ifmt >= 383 && ifmt <= 392);
				m_ifmt = ifmt;
				m_stFormat.m_rgb.Set(szFormat);
				m_pHeader.m_nSize += (uint)(m_stFormat.GetDynamicSize());
			}

			public ushort GetFormatIndex()
			{
				return m_ifmt;
			}

			public string GetFormat()
			{
				return m_stFormat.m_rgb.GetExternalString();
			}

			~Format()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontXRecord : BiffRecord
		{
			public FontXRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public FontXRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_FontX);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_iFont = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_iFont);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_FontX;
			protected void SetDefaults()
			{
				m_iFont = 0x0001;
			}

			protected ushort m_iFont;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontRecord : BiffRecord
		{
			public FontRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_FONT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_dyHeight = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_unused1 = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fItalic = (ushort)((nBitmask0 >> 1) & 0x1);
				m_unused2 = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fStrikeOut = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fOutline = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fShadow = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fCondense = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fExtend = (ushort)((nBitmask0 >> 7) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 8) & 0xff);
				m_icv.BlobRead(pBlobView);
				m_bls = pBlobView.UnpackUint16();
				m_sss = pBlobView.UnpackUint16();
				m_uls = pBlobView.UnpackUint8();
				m_bFamily = pBlobView.UnpackUint8();
				m_bCharSet = pBlobView.UnpackUint8();
				m_unused3 = pBlobView.UnpackUint8();
				m_fontName.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_dyHeight);
				int nBitmask0 = 0;
				nBitmask0 += m_unused1 << 0;
				nBitmask0 += m_fItalic << 1;
				nBitmask0 += m_unused2 << 2;
				nBitmask0 += m_fStrikeOut << 3;
				nBitmask0 += m_fOutline << 4;
				nBitmask0 += m_fShadow << 5;
				nBitmask0 += m_fCondense << 6;
				nBitmask0 += m_fExtend << 7;
				nBitmask0 += m_reserved << 8;
				pBlobView.PackUint16((ushort)(nBitmask0));
				m_icv.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_bls);
				pBlobView.PackUint16(m_sss);
				pBlobView.PackUint8(m_uls);
				pBlobView.PackUint8(m_bFamily);
				pBlobView.PackUint8(m_bCharSet);
				pBlobView.PackUint8(m_unused3);
				m_fontName.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_FONT;
			protected void SetDefaults()
			{
				m_dyHeight = 0;
				m_unused1 = 0;
				m_fItalic = 0;
				m_unused2 = 0;
				m_fStrikeOut = 0;
				m_fOutline = 0;
				m_fShadow = 0;
				m_fCondense = 0;
				m_fExtend = 0;
				m_reserved = 0;
				m_icv = new IcvFontStruct();
				m_bls = 0;
				m_sss = 0;
				m_uls = 0;
				m_bFamily = 0;
				m_bCharSet = 0;
				m_unused3 = 0;
				m_fontName = new ShortXLUnicodeStringStruct();
			}

			protected ushort m_dyHeight;
			protected ushort m_unused1;
			protected ushort m_fItalic;
			protected ushort m_unused2;
			protected ushort m_fStrikeOut;
			protected ushort m_fOutline;
			protected ushort m_fShadow;
			protected ushort m_fCondense;
			protected ushort m_fExtend;
			protected ushort m_reserved;
			protected IcvFontStruct m_icv;
			protected ushort m_bls;
			protected ushort m_sss;
			protected byte m_uls;
			protected byte m_bFamily;
			protected byte m_bCharSet;
			protected byte m_unused3;
			protected ShortXLUnicodeStringStruct m_fontName;
			public FontRecord(string szFontName, ushort nSizeTwips, ushort nColourIndex, bool bBold, bool bItalic, Font.Underline eUnderline, byte unused1 = 0, byte unused3 = 181) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_dyHeight = nSizeTwips;
				m_unused1 = unused1;
				if (bItalic)
					m_fItalic = 0x1;
				nbAssert.Assert(nColourIndex >= BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY && nColourIndex < BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY || nColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_TOOL_TIP_TEXT || nColourIndex == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_FONT_AUTOMATIC);
				m_icv.m_icv = nColourIndex;
				if (bBold)
					m_bls = 700;
				else
					m_bls = 400;
				m_uls = 0x0;
				switch (eUnderline)
				{
					case Font.Underline.UNDERLINE_SINGLE:
					{
						m_uls = 0x1;
						break;
					}

					case Font.Underline.UNDERLINE_DOUBLE:
					{
						m_uls = 0x2;
						break;
					}

					case Font.Underline.UNDERLINE_SINGLE_ACCOUNTING:
					{
						m_uls = 0x21;
						break;
					}

					case Font.Underline.UNDERLINE_DOUBLE_ACCOUNTING:
					{
						m_uls = 0x22;
						break;
					}

					case Font.Underline.UNDERLINE_NONE:
					{
						m_uls = 0x0;
						break;
					}

				}
				m_bFamily = 2;
				m_unused3 = unused3;
				m_fontName.m_rgb.Set(szFontName);
				m_pHeader.m_nSize += (uint)(m_fontName.GetDynamicSize());
			}

			public ushort GetColourIndex()
			{
				return m_icv.m_icv;
			}

			public string GetName()
			{
				return m_fontName.m_rgb.GetExternalString();
			}

			public ushort GetSizeTwips()
			{
				return m_dyHeight;
			}

			public bool GetBold()
			{
				return m_bls >= 700;
			}

			public bool GetItalic()
			{
				return m_fItalic == 0x1;
			}

			public Font.Underline GetUnderline()
			{
				switch (m_uls)
				{
					case 0x1:
					{
						return Font.Underline.UNDERLINE_SINGLE;
					}

					case 0x2:
					{
						return Font.Underline.UNDERLINE_DOUBLE;
					}

					case 0x21:
					{
						return Font.Underline.UNDERLINE_SINGLE_ACCOUNTING;
					}

					case 0x22:
					{
						return Font.Underline.UNDERLINE_DOUBLE_ACCOUNTING;
					}

				}
				return Font.Underline.UNDERLINE_NONE;
			}

			~FontRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ExternSheetRecord : BiffRecord
		{
			public ExternSheetRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_EXTERN_SHEET);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cXTI = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_cXTI);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_EXTERN_SHEET;
			protected void SetDefaults()
			{
				m_cXTI = 0;
				PostSetDefaults();
			}

			protected ushort m_cXTI;
			protected OwnedVector<XTIStruct> m_pXTIVector;
			public ExternSheetRecord(OwnedVector<WorksheetRange> pWorksheetRangeVector) : base(TYPE, (uint)(SIZE + XTIStruct.SIZE * pWorksheetRangeVector.GetSize()))
			{
				SetDefaults();
				m_cXTI = (ushort)(pWorksheetRangeVector.GetSize());
				nbAssert.Assert(m_pXTIVector == null);
				m_pXTIVector = new OwnedVector<XTIStruct>();
				for (ushort i = 0; i < m_cXTI; i++)
				{
					WorksheetRange pWorksheetRange = pWorksheetRangeVector.Get(i);
					XTIStruct pXTI = new XTIStruct();
					pXTI.m_iSupBook = 0;
					pXTI.m_itabFirst = (short)(pWorksheetRange.m_nFirst);
					pXTI.m_itabLast = (short)(pWorksheetRange.m_nLast);
					{
						NumberDuck.Secret.XTIStruct __2160298696 = pXTI;
						pXTI = null;
						m_pXTIVector.PushBack(__2160298696);
					}
				}
			}

			protected void PostSetDefaults()
			{
				m_pXTIVector = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_pXTIVector == null);
				m_pXTIVector = new OwnedVector<XTIStruct>();
				for (ushort i = 0; i < m_cXTI; i++)
				{
					XTIStruct pXTI = new XTIStruct();
					pXTI.BlobRead(pBlobView);
					{
						NumberDuck.Secret.XTIStruct __2160298696 = pXTI;
						pXTI = null;
						m_pXTIVector.PushBack(__2160298696);
					}
				}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				for (ushort i = 0; i < m_cXTI; i++)
					m_pXTIVector.Get(i).BlobWrite(pBlobView);
			}

			public ushort GetNumXTI()
			{
				return m_cXTI;
			}

			public XTIStruct GetXTIByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex < m_cXTI);
				return m_pXTIVector.Get(nIndex);
			}

			~ExternSheetRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Excel9File : BiffRecord
		{
			public Excel9File() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Excel9File(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_EXCEL9_FILE);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_EXCEL9_FILE;
			protected void SetDefaults()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class EndBlockRecord : BiffRecord
		{
			public EndBlockRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_EndBlock);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobRead(pBlobView);
				m_iObjectKind = pBlobView.UnpackUint16();
				m_unused1 = pBlobView.UnpackUint16();
				m_unused2 = pBlobView.UnpackUint16();
				m_unused3 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_iObjectKind);
				pBlobView.PackUint16(m_unused1);
				pBlobView.PackUint16(m_unused2);
				pBlobView.PackUint16(m_unused3);
			}

			protected const ushort SIZE = 12;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_EndBlock;
			protected void SetDefaults()
			{
				m_frtHeaderOld = new FrtHeaderOldStruct();
				m_iObjectKind = 0;
				m_unused1 = 0;
				m_unused2 = 0;
				m_unused3 = 0;
			}

			protected FrtHeaderOldStruct m_frtHeaderOld;
			protected ushort m_iObjectKind;
			protected ushort m_unused1;
			protected ushort m_unused2;
			protected ushort m_unused3;
			public EndBlockRecord(ushort iObjectKind) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_frtHeaderOld.m_rt = 0x0853;
				m_iObjectKind = iObjectKind;
			}

			~EndBlockRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DimensionRecord : BiffRecord
		{
			public DimensionRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DIMENSION);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_nFirstUsedRow = pBlobView.UnpackUint32();
				m_nLastUsedRow = pBlobView.UnpackUint32();
				m_nFirstUsedColumn = pBlobView.UnpackUint16();
				m_nLastUsedColumn = pBlobView.UnpackUint16();
				m_nUnused = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_nFirstUsedRow);
				pBlobView.PackUint32(m_nLastUsedRow);
				pBlobView.PackUint16(m_nFirstUsedColumn);
				pBlobView.PackUint16(m_nLastUsedColumn);
				pBlobView.PackUint16(m_nUnused);
			}

			protected const ushort SIZE = 14;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DIMENSION;
			protected void SetDefaults()
			{
				m_nFirstUsedRow = 0;
				m_nLastUsedRow = 0;
				m_nFirstUsedColumn = 0;
				m_nLastUsedColumn = 0;
				m_nUnused = 0;
			}

			protected uint m_nFirstUsedRow;
			protected uint m_nLastUsedRow;
			protected ushort m_nFirstUsedColumn;
			protected ushort m_nLastUsedColumn;
			protected ushort m_nUnused;
			public DimensionRecord(uint nFirstUsedRow, uint nLastUsedRow, ushort nFirstUsedColumn, ushort nLastUsedColumn, ushort nUnused) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_nFirstUsedRow = nFirstUsedRow;
				m_nLastUsedRow = nLastUsedRow;
				m_nFirstUsedColumn = nFirstUsedColumn;
				m_nLastUsedColumn = nLastUsedColumn;
				m_nUnused = nUnused;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefaultTextRecord : BiffRecord
		{
			public DefaultTextRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DefaultText);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_id = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_id);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DefaultText;
			protected void SetDefaults()
			{
				m_id = 0;
			}

			protected ushort m_id;
			public DefaultTextRecord(ushort id) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_id = id;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefaultRowHeight : BiffRecord
		{
			public DefaultRowHeight(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DEFAULTROWHEIGHT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fUnsynced = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fDyZero = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fExAsc = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fExDsc = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 4) & 0xfff);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fUnsynced << 0;
				nBitmask0 += m_fDyZero << 1;
				nBitmask0 += m_fExAsc << 2;
				nBitmask0 += m_fExDsc << 3;
				nBitmask0 += m_reserved << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DEFAULTROWHEIGHT;
			protected void SetDefaults()
			{
				m_fUnsynced = 0;
				m_fDyZero = 0;
				m_fExAsc = 0;
				m_fExDsc = 0;
				m_reserved = 0;
			}

			protected ushort m_fUnsynced;
			protected ushort m_fDyZero;
			protected ushort m_fExAsc;
			protected ushort m_fExDsc;
			protected ushort m_reserved;
			protected short m_miyRw;
			protected short m_miyRwHidden;
			public DefaultRowHeight(short nRowHeight) : base(TYPE, SIZE + 2)
			{
				nbAssert.Assert(nRowHeight > 0);
				SetDefaults();
				m_fUnsynced = 1;
				m_miyRw = nRowHeight;
			}

			public void PostBlobRead(BlobView pBlobView)
			{
				m_miyRw = 0;
				m_miyRwHidden = 0;
				if (m_fDyZero == 0)
					m_miyRw = pBlobView.UnpackInt16();
				else
					m_miyRwHidden = pBlobView.UnpackInt16();
			}

			public void PostBlobWrite(BlobView pBlobView)
			{
				if (m_fDyZero == 0)
					pBlobView.PackInt16(m_miyRw);
				else
					pBlobView.PackInt16(m_miyRwHidden);
			}

			public short GetRowHeight()
			{
				return m_miyRw;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DefColWidthRecord : BiffRecord
		{
			public DefColWidthRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DEFCOLWIDTH);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cchdefColWidth = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_cchdefColWidth);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DEFCOLWIDTH;
			protected void SetDefaults()
			{
				m_cchdefColWidth = 0;
			}

			protected ushort m_cchdefColWidth;
			public DefColWidthRecord(ushort cchdefColWidth) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cchdefColWidth = cchdefColWidth;
			}

			public ushort GetColumnWidth()
			{
				return m_cchdefColWidth;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Date1904 : BiffRecord
		{
			public Date1904() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Date1904(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DATE_1904);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_f1904DateSystem = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_f1904DateSystem);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DATE_1904;
			protected void SetDefaults()
			{
				m_f1904DateSystem = 0;
			}

			protected ushort m_f1904DateSystem;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DataFormatRecord : BiffRecord
		{
			public DataFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DataFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_xi = pBlobView.UnpackUint16();
				m_yi = pBlobView.UnpackUint16();
				m_iss = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fXL4iss = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 1) & 0x7fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_xi);
				pBlobView.PackUint16(m_yi);
				pBlobView.PackUint16(m_iss);
				int nBitmask0 = 0;
				nBitmask0 += m_fXL4iss << 0;
				nBitmask0 += m_reserved << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DataFormat;
			protected void SetDefaults()
			{
				m_xi = 0xFFFF;
				m_yi = 0;
				m_iss = 0;
				m_fXL4iss = 0;
				m_reserved = 0;
			}

			protected ushort m_xi;
			protected ushort m_yi;
			protected ushort m_iss;
			protected ushort m_fXL4iss;
			protected ushort m_reserved;
			public DataFormatRecord(ushort nIndex) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_yi = nIndex;
				m_iss = nIndex;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class DSF : BiffRecord
		{
			public DSF() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public DSF(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_DSF);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_reserved = pBlobView.UnpackInt16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_reserved);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_DSF;
			protected void SetDefaults()
			{
				m_reserved = 0;
			}

			protected short m_reserved;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtMlFrtRecord : BiffRecord
		{
			public CrtMlFrtRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CrtMlFrt);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_cb = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint32(m_cb);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CrtMlFrt;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_cb = 0;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected uint m_cb;
			protected void PostSetDefaults()
			{
				m_frtHeader.m_rt = 0x08A7;
			}

			~CrtMlFrtRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtLinkRecord : BiffRecord
		{
			public CrtLinkRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CrtLinkRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CrtLink);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_unused0 = pBlobView.UnpackUint16();
				m_unused1 = pBlobView.UnpackUint16();
				m_unused2 = pBlobView.UnpackUint16();
				m_unused3 = pBlobView.UnpackUint16();
				m_unused4 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_unused0);
				pBlobView.PackUint16(m_unused1);
				pBlobView.PackUint16(m_unused2);
				pBlobView.PackUint16(m_unused3);
				pBlobView.PackUint16(m_unused4);
			}

			protected const ushort SIZE = 10;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CrtLink;
			protected void SetDefaults()
			{
				m_unused0 = 0;
				m_unused1 = 0;
				m_unused2 = 0;
				m_unused3 = 0;
				m_unused4 = 0;
			}

			protected ushort m_unused0;
			protected ushort m_unused1;
			protected ushort m_unused2;
			protected ushort m_unused3;
			protected ushort m_unused4;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CrtLayout12ARecord : BiffRecord
		{
			public CrtLayout12ARecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CrtLayout12ARecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CrtLayout12A);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_dwCheckSum = pBlobView.UnpackUint32();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fLayoutTargetInner = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x7fff);
				m_xTL = pBlobView.UnpackInt16();
				m_yTL = pBlobView.UnpackInt16();
				m_xBR = pBlobView.UnpackInt16();
				m_yBR = pBlobView.UnpackInt16();
				m_wXMode = pBlobView.UnpackUint16();
				m_wYMode = pBlobView.UnpackUint16();
				m_wWidthMode = pBlobView.UnpackUint16();
				m_wHeightMode = pBlobView.UnpackUint16();
				m_x = pBlobView.UnpackDouble();
				m_y = pBlobView.UnpackDouble();
				m_dx = pBlobView.UnpackDouble();
				m_dy = pBlobView.UnpackDouble();
				m_reserved2 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint32(m_dwCheckSum);
				int nBitmask0 = 0;
				nBitmask0 += m_fLayoutTargetInner << 0;
				nBitmask0 += m_reserved1 << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackInt16(m_xTL);
				pBlobView.PackInt16(m_yTL);
				pBlobView.PackInt16(m_xBR);
				pBlobView.PackInt16(m_yBR);
				pBlobView.PackUint16(m_wXMode);
				pBlobView.PackUint16(m_wYMode);
				pBlobView.PackUint16(m_wWidthMode);
				pBlobView.PackUint16(m_wHeightMode);
				pBlobView.PackDouble(m_x);
				pBlobView.PackDouble(m_y);
				pBlobView.PackDouble(m_dx);
				pBlobView.PackDouble(m_dy);
				pBlobView.PackUint16(m_reserved2);
			}

			protected const ushort SIZE = 68;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CrtLayout12A;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_dwCheckSum = 0x00000001;
				m_fLayoutTargetInner = 0;
				m_reserved1 = 0;
				m_xTL = 23;
				m_yTL = -34;
				m_xBR = 3877;
				m_yBR = 4038;
				m_wXMode = 0;
				m_wYMode = 0;
				m_wWidthMode = 0;
				m_wHeightMode = 0;
				m_x = 0;
				m_y = 0;
				m_dx = 0;
				m_dy = 0;
				m_reserved2 = 0;
				PostSetDefaults();
			}

			protected FrtHeaderStruct m_frtHeader;
			protected uint m_dwCheckSum;
			protected ushort m_fLayoutTargetInner;
			protected ushort m_reserved1;
			protected short m_xTL;
			protected short m_yTL;
			protected short m_xBR;
			protected short m_yBR;
			protected ushort m_wXMode;
			protected ushort m_wYMode;
			protected ushort m_wWidthMode;
			protected ushort m_wHeightMode;
			protected double m_x;
			protected double m_y;
			protected double m_dx;
			protected double m_dy;
			protected ushort m_reserved2;
			public void PostSetDefaults()
			{
				m_frtHeader.m_rt = 0x08A7;
			}

			~CrtLayout12ARecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColInfoRecord : BiffRecord
		{
			public ColInfoRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_COLINFO);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_colFirst = pBlobView.UnpackUint16();
				m_colLast = pBlobView.UnpackUint16();
				m_coldx = pBlobView.UnpackUint16();
				m_ixfe = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fHidden = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fUserSet = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fBestFit = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fPhonetic = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 4) & 0xf);
				m_iOutLevel = (ushort)((nBitmask0 >> 8) & 0x7);
				m_unused1 = (ushort)((nBitmask0 >> 11) & 0x1);
				m_fCollapsed = (ushort)((nBitmask0 >> 12) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 13) & 0x7);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_colFirst);
				pBlobView.PackUint16(m_colLast);
				pBlobView.PackUint16(m_coldx);
				pBlobView.PackUint16(m_ixfe);
				int nBitmask0 = 0;
				nBitmask0 += m_fHidden << 0;
				nBitmask0 += m_fUserSet << 1;
				nBitmask0 += m_fBestFit << 2;
				nBitmask0 += m_fPhonetic << 3;
				nBitmask0 += m_reserved1 << 4;
				nBitmask0 += m_iOutLevel << 8;
				nBitmask0 += m_unused1 << 11;
				nBitmask0 += m_fCollapsed << 12;
				nBitmask0 += m_reserved2 << 13;
				pBlobView.PackUint16((ushort)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 10;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_COLINFO;
			protected void SetDefaults()
			{
				m_colFirst = 0;
				m_colLast = 0;
				m_coldx = 0;
				m_ixfe = 15;
				m_fHidden = 0;
				m_fUserSet = 0;
				m_fBestFit = 0;
				m_fPhonetic = 0;
				m_reserved1 = 0;
				m_iOutLevel = 0;
				m_unused1 = 0;
				m_fCollapsed = 0;
				m_reserved2 = 0;
			}

			protected ushort m_colFirst;
			protected ushort m_colLast;
			protected ushort m_coldx;
			protected ushort m_ixfe;
			protected ushort m_fHidden;
			protected ushort m_fUserSet;
			protected ushort m_fBestFit;
			protected ushort m_fPhonetic;
			protected ushort m_reserved1;
			protected ushort m_iOutLevel;
			protected ushort m_unused1;
			protected ushort m_fCollapsed;
			protected ushort m_reserved2;
			public ColInfoRecord(ushort nFirstColumn, ushort nLastColumn, ushort nColumnWidth, bool bHidden) : base(TYPE, SIZE + 2)
			{
				SetDefaults();
				m_colFirst = nFirstColumn;
				m_colLast = nLastColumn;
				m_coldx = nColumnWidth;
				if (bHidden)
					m_fHidden = 0x1;
			}

			public void PostBlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(0);
			}

			public ushort GetFirstColumn()
			{
				return m_colFirst;
			}

			public ushort GetLastColumn()
			{
				return m_colLast;
			}

			public ushort GetColumnWidth()
			{
				return m_coldx;
			}

			public bool GetHidden()
			{
				return m_fHidden == 0x1;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CodePage : BiffRecord
		{
			public CodePage() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CodePage(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CODE_PAGE);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cv = pBlobView.UnpackInt16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_cv);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CODE_PAGE;
			protected void SetDefaults()
			{
				m_cv = 1200;
			}

			protected short m_cv;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartRecord : BiffRecord
		{
			public ChartRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ChartRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Chart);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_x = pBlobView.UnpackUint32();
				m_y = pBlobView.UnpackUint32();
				m_dx = pBlobView.UnpackUint32();
				m_dy = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_x);
				pBlobView.PackUint32(m_y);
				pBlobView.PackUint32(m_dx);
				pBlobView.PackUint32(m_dy);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Chart;
			protected void SetDefaults()
			{
				m_x = 0;
				m_y = 0;
				m_dx = 23642112;
				m_dy = 14204928;
			}

			protected uint m_x;
			protected uint m_y;
			protected uint m_dx;
			protected uint m_dy;
			public ushort GetX()
			{
				return (ushort)(m_x >> 16);
			}

			public ushort GetY()
			{
				return (ushort)(m_y >> 16);
			}

			public ushort GetWidth()
			{
				return (ushort)(m_dx >> 16);
			}

			public ushort GetHeight()
			{
				return (ushort)(m_dy >> 16);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartFrtInfoRecord : BiffRecord
		{
			public ChartFrtInfoRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ChartFrtInfoRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ChartFrtInfo);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobRead(pBlobView);
				m_verOriginator = pBlobView.UnpackUint8();
				m_verWriter = pBlobView.UnpackUint8();
				m_cCFRTID = pBlobView.UnpackUint16();
				m_hax1 = pBlobView.UnpackUint16();
				m_hax2 = pBlobView.UnpackUint16();
				m_hax3 = pBlobView.UnpackUint16();
				m_hax4 = pBlobView.UnpackUint16();
				m_hax5 = pBlobView.UnpackUint16();
				m_hax6 = pBlobView.UnpackUint16();
				m_hax7 = pBlobView.UnpackUint16();
				m_hax8 = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobWrite(pBlobView);
				pBlobView.PackUint8(m_verOriginator);
				pBlobView.PackUint8(m_verWriter);
				pBlobView.PackUint16(m_cCFRTID);
				pBlobView.PackUint16(m_hax1);
				pBlobView.PackUint16(m_hax2);
				pBlobView.PackUint16(m_hax3);
				pBlobView.PackUint16(m_hax4);
				pBlobView.PackUint16(m_hax5);
				pBlobView.PackUint16(m_hax6);
				pBlobView.PackUint16(m_hax7);
				pBlobView.PackUint16(m_hax8);
			}

			protected const ushort SIZE = 24;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ChartFrtInfo;
			protected void SetDefaults()
			{
				m_frtHeaderOld = new FrtHeaderOldStruct();
				m_verOriginator = 0xE;
				m_verWriter = 0xE;
				m_cCFRTID = 0x0004;
				m_hax1 = 0x0850;
				m_hax2 = 0x085a;
				m_hax3 = 0x0861;
				m_hax4 = 0x0861;
				m_hax5 = 0x086a;
				m_hax6 = 0x086b;
				m_hax7 = 0x089d;
				m_hax8 = 0x08a6;
				PostSetDefaults();
			}

			protected FrtHeaderOldStruct m_frtHeaderOld;
			protected byte m_verOriginator;
			protected byte m_verWriter;
			protected ushort m_cCFRTID;
			protected ushort m_hax1;
			protected ushort m_hax2;
			protected ushort m_hax3;
			protected ushort m_hax4;
			protected ushort m_hax5;
			protected ushort m_hax6;
			protected ushort m_hax7;
			protected ushort m_hax8;
			protected void PostSetDefaults()
			{
				m_frtHeaderOld.m_rt = 2128;
			}

			~ChartFrtInfoRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartFormatRecord : BiffRecord
		{
			public ChartFormatRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public ChartFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_ChartFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_reserved1 = pBlobView.UnpackUint32();
				m_reserved2 = pBlobView.UnpackUint32();
				m_reserved3 = pBlobView.UnpackUint32();
				m_reserved4 = pBlobView.UnpackUint32();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fVaried = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved5 = (ushort)((nBitmask0 >> 1) & 0x7fff);
				m_icrt = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_reserved1);
				pBlobView.PackUint32(m_reserved2);
				pBlobView.PackUint32(m_reserved3);
				pBlobView.PackUint32(m_reserved4);
				int nBitmask0 = 0;
				nBitmask0 += m_fVaried << 0;
				nBitmask0 += m_reserved5 << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icrt);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_ChartFormat;
			protected void SetDefaults()
			{
				m_reserved1 = 0;
				m_reserved2 = 0;
				m_reserved3 = 0;
				m_reserved4 = 0;
				m_fVaried = 0;
				m_reserved5 = 0;
				m_icrt = 0;
			}

			protected uint m_reserved1;
			protected uint m_reserved2;
			protected uint m_reserved3;
			protected uint m_reserved4;
			protected ushort m_fVaried;
			protected ushort m_reserved5;
			protected ushort m_icrt;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Chart3DBarShapeRecord : BiffRecord
		{
			public Chart3DBarShapeRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Chart3DBarShapeRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Chart3DBarShape);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_riser = pBlobView.UnpackUint8();
				m_taper = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_riser);
				pBlobView.PackUint8(m_taper);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Chart3DBarShape;
			protected void SetDefaults()
			{
				m_riser = 0;
				m_taper = 0;
			}

			protected byte m_riser;
			protected byte m_taper;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CatSerRangeRecord : BiffRecord
		{
			public CatSerRangeRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CatSerRangeRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CatSerRange);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_catCross = pBlobView.UnpackInt16();
				m_catLabel = pBlobView.UnpackInt16();
				m_catMark = pBlobView.UnpackInt16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fBetween = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fMaxCross = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fReverse = (ushort)((nBitmask0 >> 2) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 3) & 0x1fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_catCross);
				pBlobView.PackInt16(m_catLabel);
				pBlobView.PackInt16(m_catMark);
				int nBitmask0 = 0;
				nBitmask0 += m_fBetween << 0;
				nBitmask0 += m_fMaxCross << 1;
				nBitmask0 += m_fReverse << 2;
				nBitmask0 += m_reserved << 3;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CatSerRange;
			protected void SetDefaults()
			{
				m_catCross = 1;
				m_catLabel = 1;
				m_catMark = 1;
				m_fBetween = 1;
				m_fMaxCross = 0;
				m_fReverse = 0;
				m_reserved = 0;
			}

			protected short m_catCross;
			protected short m_catLabel;
			protected short m_catMark;
			protected ushort m_fBetween;
			protected ushort m_fMaxCross;
			protected ushort m_fReverse;
			protected ushort m_reserved;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CatLabRecord : BiffRecord
		{
			public CatLabRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CatLabRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CatLab);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobRead(pBlobView);
				m_wOffset = pBlobView.UnpackUint16();
				m_at = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_cAutoCatLabelReal = (ushort)((nBitmask0 >> 0) & 0x1);
				m_unused = (ushort)((nBitmask0 >> 1) & 0x7fff);
				m_reserved = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeaderOld.BlobWrite(pBlobView);
				pBlobView.PackUint16(m_wOffset);
				pBlobView.PackUint16(m_at);
				int nBitmask0 = 0;
				nBitmask0 += m_cAutoCatLabelReal << 0;
				nBitmask0 += m_unused << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_reserved);
			}

			protected const ushort SIZE = 12;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CatLab;
			protected void SetDefaults()
			{
				m_frtHeaderOld = new FrtHeaderOldStruct();
				m_wOffset = 100;
				m_at = 0x0002;
				m_cAutoCatLabelReal = 0;
				m_unused = 158;
				m_reserved = 0;
				PostSetDefaults();
			}

			protected FrtHeaderOldStruct m_frtHeaderOld;
			protected ushort m_wOffset;
			protected ushort m_at;
			protected ushort m_cAutoCatLabelReal;
			protected ushort m_unused;
			protected ushort m_reserved;
			protected void PostSetDefaults()
			{
				m_frtHeaderOld.m_rt = 0x0856;
			}

			~CatLabRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CalcPrecision : BiffRecord
		{
			public CalcPrecision() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CalcPrecision(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CALC_PRECISION);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fFullPrec = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fFullPrec);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CALC_PRECISION;
			protected void SetDefaults()
			{
				m_fFullPrec = 1;
			}

			protected ushort m_fFullPrec;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CalcCountRecord : BiffRecord
		{
			public CalcCountRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public CalcCountRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_CALCCOUNT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_nMaxNumIteration = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_nMaxNumIteration);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_CALCCOUNT;
			protected void SetDefaults()
			{
				m_nMaxNumIteration = 100;
			}

			protected ushort m_nMaxNumIteration;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BuiltInFnGroupCount : BiffRecord
		{
			public BuiltInFnGroupCount() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public BuiltInFnGroupCount(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BUILT_IN_FN_GROUP_COUNT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_count = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_count);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BUILT_IN_FN_GROUP_COUNT;
			protected void SetDefaults()
			{
				m_count = 17;
			}

			protected ushort m_count;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BraiRecord : BiffRecord
		{
			public BraiRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BRAI);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_id = pBlobView.UnpackUint8();
				m_rt = pBlobView.UnpackUint8();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fUnlinkedIfmt = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 1) & 0x7fff);
				m_ifmt = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_id);
				pBlobView.PackUint8(m_rt);
				int nBitmask0 = 0;
				nBitmask0 += m_fUnlinkedIfmt << 0;
				nBitmask0 += m_reserved << 1;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_ifmt);
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 6;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BRAI;
			protected void SetDefaults()
			{
				m_id = 0;
				m_rt = 0;
				m_fUnlinkedIfmt = 0;
				m_reserved = 0;
				m_ifmt = 0;
				PostSetDefaults();
			}

			protected byte m_id;
			protected byte m_rt;
			protected ushort m_fUnlinkedIfmt;
			protected ushort m_reserved;
			protected ushort m_ifmt;
			protected CellParsedFormulaStruct m_formula;
			public BraiRecord(byte id, byte rt, bool fUnlinkedIfmt, ushort ifmt, Formula pFormula, WorkbookGlobals pWorkbookGlobals) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_id = id;
				m_rt = rt;
				m_fUnlinkedIfmt = 0x0;
				if (fUnlinkedIfmt)
					m_fUnlinkedIfmt = 0x1;
				nbAssert.Assert(m_formula == null);
				m_formula = new CellParsedFormulaStruct(pFormula, pWorkbookGlobals);
				m_pHeader.m_nSize += (ushort)(m_formula.GetSize());
			}

			protected void PostSetDefaults()
			{
				m_formula = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_formula == null);
				m_formula = new CellParsedFormulaStruct();
				m_formula.BlobRead(pBlobView);
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_formula.BlobWrite(pBlobView);
			}

			public byte GetId()
			{
				return m_id;
			}

			public byte GetRt()
			{
				return m_rt;
			}

			public Formula GetFormula(WorkbookGlobals pWorkbookGlobals)
			{
				return new Formula(m_formula.m_rgce.m_pParsedExpressionRecordVector, pWorkbookGlobals);
			}

			~BraiRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoundSheet8Record : BiffRecord
		{
			public BoundSheet8Record(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOUND_SHEET_8);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_lbPlyPos = pBlobView.UnpackUint32();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_hsState = (byte)((nBitmask0 >> 0) & 0x3);
				m_reserved = (byte)((nBitmask0 >> 2) & 0x3f);
				m_dt = pBlobView.UnpackUint8();
				m_stName.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_lbPlyPos);
				int nBitmask0 = 0;
				nBitmask0 += m_hsState << 0;
				nBitmask0 += m_reserved << 2;
				pBlobView.PackUint8((byte)(nBitmask0));
				pBlobView.PackUint8(m_dt);
				m_stName.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOUND_SHEET_8;
			protected void SetDefaults()
			{
				m_lbPlyPos = 0;
				m_hsState = 0;
				m_reserved = 0;
				m_dt = 0;
				m_stName = new ShortXLUnicodeStringStruct();
			}

			protected uint m_lbPlyPos;
			protected byte m_hsState;
			protected byte m_reserved;
			protected byte m_dt;
			protected ShortXLUnicodeStringStruct m_stName;
			public BoundSheet8Record(string sxName) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_stName.m_rgb.Set(sxName);
				m_pHeader.m_nSize += (uint)(m_stName.GetDynamicSize());
			}

			public string GetName()
			{
				return m_stName.m_rgb.GetExternalString();
			}

			public void SetStreamOffset(uint nStreamOffset)
			{
				m_lbPlyPos = nStreamOffset;
			}

			~BoundSheet8Record()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BottomMarginRecord : BiffRecord
		{
			public BottomMarginRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOTTOM_MARGIN);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_num = pBlobView.UnpackDouble();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackDouble(m_num);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOTTOM_MARGIN;
			protected void SetDefaults()
			{
				m_num = 0.0;
			}

			protected double m_num;
			public BottomMarginRecord(double num) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_num = num;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoolErrRecord : BiffRecord
		{
			public BoolErrRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOOLERR);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cell.BlobRead(pBlobView);
				m_bBoolErr = pBlobView.UnpackUint8();
				m_fError = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_cell.BlobWrite(pBlobView);
				pBlobView.PackUint8(m_bBoolErr);
				pBlobView.PackUint8(m_fError);
			}

			protected const ushort SIZE = 8;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOOLERR;
			protected void SetDefaults()
			{
				m_cell = new CellStruct();
				m_bBoolErr = 0;
				m_fError = 0;
			}

			protected CellStruct m_cell;
			protected byte m_bBoolErr;
			protected byte m_fError;
			public BoolErrRecord(ushort nX, ushort nY, ushort nXfIndex, bool bBoolean) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cell.m_rw.m_rw = nY;
				m_cell.m_col.m_col = nX;
				m_cell.m_ixfe.m_ixfe = nXfIndex;
				m_bBoolErr = 0x00;
				if (bBoolean)
					m_bBoolErr = 0x01;
			}

			public ushort GetX()
			{
				return m_cell.m_col.m_col;
			}

			public ushort GetY()
			{
				return m_cell.m_rw.m_rw;
			}

			public ushort GetXfIndex()
			{
				return m_cell.m_ixfe.m_ixfe;
			}

			public bool IsBoolean()
			{
				return m_fError == 0x00;
			}

			public bool GetBoolean()
			{
				return m_bBoolErr == 0x01;
			}

			~BoolErrRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BookExtRecord : BiffRecord
		{
			public BookExtRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOOK_EXT);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_frtHeader.BlobRead(pBlobView);
				m_cb = pBlobView.UnpackUint32();
				uint nBitmask0 = pBlobView.UnpackUint32();
				m_fDontAutoRecover = (uint)((nBitmask0 >> 0) & 0x1);
				m_fHidePivotList = (uint)((nBitmask0 >> 1) & 0x1);
				m_fFilterPrivacy = (uint)((nBitmask0 >> 2) & 0x1);
				m_fEmbedFactoids = (uint)((nBitmask0 >> 3) & 0x1);
				m_mdFactoidDisplay = (uint)((nBitmask0 >> 4) & 0x3);
				m_fSavedDuringRecovery = (uint)((nBitmask0 >> 6) & 0x1);
				m_fCreatedViaMinimalSave = (uint)((nBitmask0 >> 7) & 0x1);
				m_fOpenedViaDataRecovery = (uint)((nBitmask0 >> 8) & 0x1);
				m_fOpenedViaSafeLoad = (uint)((nBitmask0 >> 9) & 0x1);
				m_reserved = (uint)((nBitmask0 >> 10) & 0x3fffff);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_frtHeader.BlobWrite(pBlobView);
				pBlobView.PackUint32(m_cb);
				uint nBitmask0 = 0;
				nBitmask0 += m_fDontAutoRecover << 0;
				nBitmask0 += m_fHidePivotList << 1;
				nBitmask0 += m_fFilterPrivacy << 2;
				nBitmask0 += m_fEmbedFactoids << 3;
				nBitmask0 += m_mdFactoidDisplay << 4;
				nBitmask0 += m_fSavedDuringRecovery << 6;
				nBitmask0 += m_fCreatedViaMinimalSave << 7;
				nBitmask0 += m_fOpenedViaDataRecovery << 8;
				nBitmask0 += m_fOpenedViaSafeLoad << 9;
				nBitmask0 += m_reserved << 10;
				pBlobView.PackUint32((uint)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			protected const ushort SIZE = 20;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOOK_EXT;
			protected void SetDefaults()
			{
				m_frtHeader = new FrtHeaderStruct();
				m_cb = 0;
				m_fDontAutoRecover = 0;
				m_fHidePivotList = 0;
				m_fFilterPrivacy = 0;
				m_fEmbedFactoids = 0;
				m_mdFactoidDisplay = 0;
				m_fSavedDuringRecovery = 0;
				m_fCreatedViaMinimalSave = 0;
				m_fOpenedViaDataRecovery = 0;
				m_fOpenedViaSafeLoad = 0;
				m_reserved = 0;
			}

			protected FrtHeaderStruct m_frtHeader;
			protected uint m_cb;
			protected uint m_fDontAutoRecover;
			protected uint m_fHidePivotList;
			protected uint m_fFilterPrivacy;
			protected uint m_fEmbedFactoids;
			protected uint m_mdFactoidDisplay;
			protected uint m_fSavedDuringRecovery;
			protected uint m_fCreatedViaMinimalSave;
			protected uint m_fOpenedViaDataRecovery;
			protected uint m_fOpenedViaSafeLoad;
			protected uint m_reserved;
			protected byte m_grbit1;
			protected byte m_grbit2;
			public BookExtRecord() : base(TYPE, SIZE + 2)
			{
				SetDefaults();
				m_frtHeader.m_rt = (ushort)(BiffRecord.Type.TYPE_BOOK_EXT);
				m_cb = m_pHeader.m_nSize;
				m_grbit1 = 2;
				m_grbit2 = 0;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_cb > 20)
					m_grbit1 = pBlobView.UnpackUint8();
				if (m_cb > 21)
					m_grbit2 = pBlobView.UnpackUint8();
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_cb > 20)
					pBlobView.PackUint8(m_grbit1);
				if (m_cb > 21)
					pBlobView.PackUint8(m_grbit2);
			}

			~BookExtRecord()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BookBool : BiffRecord
		{
			public BookBool() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public BookBool(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOOK_BOOL);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fNoSaveSup = (ushort)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fHasEnvelope = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fEnvelopeVisible = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fEnvelopeInitDone = (ushort)((nBitmask0 >> 4) & 0x1);
				m_grUpdateLinks = (ushort)((nBitmask0 >> 5) & 0x3);
				m_unused = (ushort)((nBitmask0 >> 7) & 0x1);
				m_fHideBorderUnselLists = (ushort)((nBitmask0 >> 8) & 0x1);
				m_reserved2 = (ushort)((nBitmask0 >> 9) & 0x7f);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fNoSaveSup << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fHasEnvelope << 2;
				nBitmask0 += m_fEnvelopeVisible << 3;
				nBitmask0 += m_fEnvelopeInitDone << 4;
				nBitmask0 += m_grUpdateLinks << 5;
				nBitmask0 += m_unused << 7;
				nBitmask0 += m_fHideBorderUnselLists << 8;
				nBitmask0 += m_reserved2 << 9;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOOK_BOOL;
			protected void SetDefaults()
			{
				m_fNoSaveSup = 0;
				m_reserved1 = 0;
				m_fHasEnvelope = 0;
				m_fEnvelopeVisible = 0;
				m_fEnvelopeInitDone = 0;
				m_grUpdateLinks = 0;
				m_unused = 0;
				m_fHideBorderUnselLists = 0;
				m_reserved2 = 0;
			}

			protected ushort m_fNoSaveSup;
			protected ushort m_reserved1;
			protected ushort m_fHasEnvelope;
			protected ushort m_fEnvelopeVisible;
			protected ushort m_fEnvelopeInitDone;
			protected ushort m_grUpdateLinks;
			protected ushort m_unused;
			protected ushort m_fHideBorderUnselLists;
			protected ushort m_reserved2;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BofRecord : BiffRecord
		{
			public BofRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BOF);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_vers = pBlobView.UnpackUint16();
				m_dt = pBlobView.UnpackUint16();
				m_rupBuild = pBlobView.UnpackUint16();
				m_rupYear = pBlobView.UnpackUint16();
				m_nFileHistoryFlags = pBlobView.UnpackUint32();
				m_nMinimumExcelVersion = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_vers);
				pBlobView.PackUint16(m_dt);
				pBlobView.PackUint16(m_rupBuild);
				pBlobView.PackUint16(m_rupYear);
				pBlobView.PackUint32(m_nFileHistoryFlags);
				pBlobView.PackUint32(m_nMinimumExcelVersion);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BOF;
			protected void SetDefaults()
			{
				m_vers = 0x0600;
				m_dt = 0;
				m_rupBuild = 12902;
				m_rupYear = 1997;
				m_nFileHistoryFlags = 98505;
				m_nMinimumExcelVersion = 1542;
			}

			protected ushort m_vers;
			protected ushort m_dt;
			protected ushort m_rupBuild;
			protected ushort m_rupYear;
			protected uint m_nFileHistoryFlags;
			protected uint m_nMinimumExcelVersion;
			public enum BofType
			{
				BOF_TYPE_WORKBOOK_GLOBALS = 0x005,
				BOF_TYPE_VISUAL_BASIC_MODULE = 0x0006,
				BOF_TYPE_SHEET = 0x0010,
				BOF_TYPE_CHART = 0x0020,
				BOF_TYPE_MACRO_SHEET = 0x0040,
			}

			public BofRecord(BofType eType) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_dt = (ushort)(eType);
			}

			public BofRecord.BofType GetBofType()
			{
				return (BofType)(m_dt);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Blank : BiffRecord
		{
			public Blank(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BLANK);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cell.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_cell.BlobWrite(pBlobView);
			}

			protected const ushort SIZE = 6;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BLANK;
			protected void SetDefaults()
			{
				m_cell = new CellStruct();
			}

			protected CellStruct m_cell;
			public Blank(ushort nX, ushort nY, ushort nXfIndex) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_cell.m_rw.m_rw = nY;
				m_cell.m_col.m_col = nX;
				m_cell.m_ixfe.m_ixfe = nXfIndex;
			}

			public ushort GetX()
			{
				return m_cell.m_col.m_col;
			}

			public ushort GetY()
			{
				return m_cell.m_rw.m_rw;
			}

			public ushort GetXfIndex()
			{
				return m_cell.m_ixfe.m_ixfe;
			}

			~Blank()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffRecord_ContinueInfo
		{
			public int m_nOffset;
			public int m_nType;
			public BiffRecord_ContinueInfo(int nOffset, int nType)
			{
				m_nOffset = nOffset;
				m_nType = nType;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BeginRecord : BiffRecord
		{
			public BeginRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public BeginRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Begin);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
			}

			public override void BlobWrite(BlobView pBlobView)
			{
			}

			protected const ushort SIZE = 0;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Begin;
			protected void SetDefaults()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BarRecord : BiffRecord
		{
			public BarRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Bar);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_pcOverlap = pBlobView.UnpackInt16();
				m_pcGap = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fTranspose = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fStacked = (ushort)((nBitmask0 >> 1) & 0x1);
				m_f100 = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fHasShadow = (ushort)((nBitmask0 >> 3) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 4) & 0xfff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackInt16(m_pcOverlap);
				pBlobView.PackUint16(m_pcGap);
				int nBitmask0 = 0;
				nBitmask0 += m_fTranspose << 0;
				nBitmask0 += m_fStacked << 1;
				nBitmask0 += m_f100 << 2;
				nBitmask0 += m_fHasShadow << 3;
				nBitmask0 += m_reserved << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 6;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Bar;
			protected void SetDefaults()
			{
				m_pcOverlap = 0;
				m_pcGap = 150;
				m_fTranspose = 0;
				m_fStacked = 0;
				m_f100 = 0;
				m_fHasShadow = 0;
				m_reserved = 0;
			}

			protected short m_pcOverlap;
			protected ushort m_pcGap;
			protected ushort m_fTranspose;
			protected ushort m_fStacked;
			protected ushort m_f100;
			protected ushort m_fHasShadow;
			protected ushort m_reserved;
			public BarRecord(Chart.Type eType) : base(TYPE, SIZE)
			{
				nbAssert.Assert(eType == Chart.Type.TYPE_COLUMN || eType == Chart.Type.TYPE_COLUMN_STACKED || eType == Chart.Type.TYPE_COLUMN_STACKED_100 || eType == Chart.Type.TYPE_BAR || eType == Chart.Type.TYPE_BAR_STACKED || eType == Chart.Type.TYPE_BAR_STACKED_100);
				SetDefaults();
				if (eType == Chart.Type.TYPE_BAR || eType == Chart.Type.TYPE_BAR_STACKED || eType == Chart.Type.TYPE_BAR_STACKED_100)
					m_fTranspose = 0x1;
				if (eType == Chart.Type.TYPE_COLUMN_STACKED || eType == Chart.Type.TYPE_COLUMN_STACKED_100 || eType == Chart.Type.TYPE_BAR_STACKED || eType == Chart.Type.TYPE_BAR_STACKED_100)
				{
					m_fStacked = 0x1;
				}
				if (eType == Chart.Type.TYPE_COLUMN_STACKED_100 || eType == Chart.Type.TYPE_BAR_STACKED_100)
					m_f100 = 0x1;
			}

			public Chart.Type GetChartType()
			{
				if (m_fTranspose == 0x1)
				{
					if (m_fStacked == 0x1)
						if (m_f100 == 0x1)
							return Chart.Type.TYPE_BAR_STACKED_100;
						else
							return Chart.Type.TYPE_BAR_STACKED;
					return Chart.Type.TYPE_BAR;
				}
				if (m_fStacked == 0x1)
					if (m_f100 == 0x1)
						return Chart.Type.TYPE_COLUMN_STACKED_100;
					else
						return Chart.Type.TYPE_COLUMN_STACKED;
				return Chart.Type.TYPE_COLUMN;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Backup : BiffRecord
		{
			public Backup() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public Backup(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_BACKUP);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fBackup = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fBackup);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_BACKUP;
			protected void SetDefaults()
			{
				m_fBackup = 0;
			}

			protected ushort m_fBackup;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisRecord : BiffRecord
		{
			public AxisRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Axis);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_wType = pBlobView.UnpackUint16();
				m_reserved1 = pBlobView.UnpackUint32();
				m_reserved2 = pBlobView.UnpackUint32();
				m_reserved3 = pBlobView.UnpackUint32();
				m_reserved4 = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_wType);
				pBlobView.PackUint32(m_reserved1);
				pBlobView.PackUint32(m_reserved2);
				pBlobView.PackUint32(m_reserved3);
				pBlobView.PackUint32(m_reserved4);
			}

			protected const ushort SIZE = 18;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Axis;
			protected void SetDefaults()
			{
				m_wType = 0x0000;
				m_reserved1 = 0;
				m_reserved2 = 0;
				m_reserved3 = 0;
				m_reserved4 = 0;
			}

			protected ushort m_wType;
			protected uint m_reserved1;
			protected uint m_reserved2;
			protected uint m_reserved3;
			protected uint m_reserved4;
			public AxisRecord(ushort wType) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_wType = wType;
			}

			public new ushort GetType()
			{
				return m_wType;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisParentRecord : BiffRecord
		{
			public AxisParentRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public AxisParentRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_AxisParent);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_iax = pBlobView.UnpackUint16();
				m_unused1 = pBlobView.UnpackUint32();
				m_unused2 = pBlobView.UnpackUint32();
				m_unused3 = pBlobView.UnpackUint32();
				m_unused4 = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_iax);
				pBlobView.PackUint32(m_unused1);
				pBlobView.PackUint32(m_unused2);
				pBlobView.PackUint32(m_unused3);
				pBlobView.PackUint32(m_unused4);
			}

			protected const ushort SIZE = 18;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_AxisParent;
			protected void SetDefaults()
			{
				m_iax = 0x0000;
				m_unused1 = 0x000000f6;
				m_unused2 = 0x00000081;
				m_unused3 = 0x00000e44;
				m_unused4 = 0x00000d63;
			}

			protected ushort m_iax;
			protected uint m_unused1;
			protected uint m_unused2;
			protected uint m_unused3;
			protected uint m_unused4;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxisLineRecord : BiffRecord
		{
			public AxisLineRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_AxisLine);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_id = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_id);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_AxisLine;
			protected void SetDefaults()
			{
				m_id = 0;
			}

			protected ushort m_id;
			public AxisLineRecord(ushort id) : base(TYPE, SIZE)
			{
				m_id = id;
			}

			public ushort GetId()
			{
				SetDefaults();
				return m_id;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxesUsedRecord : BiffRecord
		{
			public AxesUsedRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public AxesUsedRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_AxesUsed);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cAxes = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_cAxes);
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_AxesUsed;
			protected void SetDefaults()
			{
				m_cAxes = 0x0001;
			}

			protected ushort m_cAxes;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AxcExtRecord : BiffRecord
		{
			public AxcExtRecord() : base(TYPE, SIZE)
			{
				SetDefaults();
			}

			public AxcExtRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_AxcExt);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_catMin = pBlobView.UnpackUint16();
				m_catMax = pBlobView.UnpackUint16();
				m_catMajor = pBlobView.UnpackUint16();
				m_duMajor = pBlobView.UnpackUint16();
				m_catMinor = pBlobView.UnpackUint16();
				m_duMinor = pBlobView.UnpackUint16();
				m_duBase = pBlobView.UnpackUint16();
				m_catCrossDate = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoMin = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fAutoMax = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fAutoMajor = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fAutoMinor = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fDateAxis = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fAutoBase = (ushort)((nBitmask0 >> 5) & 0x1);
				m_fAutoCross = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fAutoDate = (ushort)((nBitmask0 >> 7) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 8) & 0xff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_catMin);
				pBlobView.PackUint16(m_catMax);
				pBlobView.PackUint16(m_catMajor);
				pBlobView.PackUint16(m_duMajor);
				pBlobView.PackUint16(m_catMinor);
				pBlobView.PackUint16(m_duMinor);
				pBlobView.PackUint16(m_duBase);
				pBlobView.PackUint16(m_catCrossDate);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoMin << 0;
				nBitmask0 += m_fAutoMax << 1;
				nBitmask0 += m_fAutoMajor << 2;
				nBitmask0 += m_fAutoMinor << 3;
				nBitmask0 += m_fDateAxis << 4;
				nBitmask0 += m_fAutoBase << 5;
				nBitmask0 += m_fAutoCross << 6;
				nBitmask0 += m_fAutoDate << 7;
				nBitmask0 += m_reserved << 8;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 18;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_AxcExt;
			protected void SetDefaults()
			{
				m_catMin = 0;
				m_catMax = 0;
				m_catMajor = 0x0001;
				m_duMajor = 0;
				m_catMinor = 0x0001;
				m_duMinor = 0;
				m_duBase = 0;
				m_catCrossDate = 0;
				m_fAutoMin = 0x1;
				m_fAutoMax = 0x1;
				m_fAutoMajor = 0x1;
				m_fAutoMinor = 0x1;
				m_fDateAxis = 0;
				m_fAutoBase = 0x1;
				m_fAutoCross = 0x1;
				m_fAutoDate = 0x1;
				m_reserved = 0;
			}

			protected ushort m_catMin;
			protected ushort m_catMax;
			protected ushort m_catMajor;
			protected ushort m_duMajor;
			protected ushort m_catMinor;
			protected ushort m_duMinor;
			protected ushort m_duBase;
			protected ushort m_catCrossDate;
			protected ushort m_fAutoMin;
			protected ushort m_fAutoMax;
			protected ushort m_fAutoMajor;
			protected ushort m_fAutoMinor;
			protected ushort m_fDateAxis;
			protected ushort m_fAutoBase;
			protected ushort m_fAutoCross;
			protected ushort m_fAutoDate;
			protected ushort m_reserved;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaRecord : BiffRecord
		{
			public AreaRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_Area);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fStacked = (ushort)((nBitmask0 >> 0) & 0x1);
				m_f100 = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fHasShadow = (ushort)((nBitmask0 >> 2) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 3) & 0x1fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fStacked << 0;
				nBitmask0 += m_f100 << 1;
				nBitmask0 += m_fHasShadow << 2;
				nBitmask0 += m_reserved << 3;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			protected const ushort SIZE = 2;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_Area;
			protected void SetDefaults()
			{
				m_fStacked = 0;
				m_f100 = 0;
				m_fHasShadow = 0;
				m_reserved = 0;
			}

			protected ushort m_fStacked;
			protected ushort m_f100;
			protected ushort m_fHasShadow;
			protected ushort m_reserved;
			public AreaRecord(Chart.Type eType) : base(TYPE, SIZE)
			{
				nbAssert.Assert(eType == Chart.Type.TYPE_AREA || eType == Chart.Type.TYPE_AREA_STACKED || eType == Chart.Type.TYPE_AREA_STACKED_100);
				SetDefaults();
				if (eType == Chart.Type.TYPE_AREA_STACKED || eType == Chart.Type.TYPE_AREA_STACKED_100)
					m_fStacked = 1;
				if (eType == Chart.Type.TYPE_AREA_STACKED_100)
					m_f100 = 1;
			}

			public Chart.Type GetChartType()
			{
				if (m_fStacked == 1)
					if (m_f100 == 1)
						return Chart.Type.TYPE_AREA_STACKED_100;
					else
						return Chart.Type.TYPE_AREA_STACKED;
				return Chart.Type.TYPE_AREA;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaFormatRecord : BiffRecord
		{
			public AreaFormatRecord(BiffHeader pHeader, Stream pStream) : base(pHeader, pStream)
			{
				nbAssert.Assert((BiffRecord.Type)(m_pHeader.m_nType) == BiffRecord.Type.TYPE_AreaFormat);
				SetDefaults();
				BlobRead(m_pBlobView);
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rgbFore = pBlobView.UnpackUint32();
				m_rgbBack = pBlobView.UnpackUint32();
				m_fls = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAuto = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fInvertNeg = (ushort)((nBitmask0 >> 1) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 2) & 0x3fff);
				m_icvFore = pBlobView.UnpackUint16();
				m_icvBack = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_rgbFore);
				pBlobView.PackUint32(m_rgbBack);
				pBlobView.PackUint16(m_fls);
				int nBitmask0 = 0;
				nBitmask0 += m_fAuto << 0;
				nBitmask0 += m_fInvertNeg << 1;
				nBitmask0 += m_reserved << 2;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_icvFore);
				pBlobView.PackUint16(m_icvBack);
			}

			protected const ushort SIZE = 16;
			protected const BiffRecord.Type TYPE = BiffRecord.Type.TYPE_AreaFormat;
			protected void SetDefaults()
			{
				m_rgbFore = 0;
				m_rgbBack = 0;
				m_fls = 0;
				m_fAuto = 0;
				m_fInvertNeg = 0;
				m_reserved = 0;
				m_icvFore = 0;
				m_icvBack = 0;
			}

			protected uint m_rgbFore;
			protected uint m_rgbBack;
			protected ushort m_fls;
			protected ushort m_fAuto;
			protected ushort m_fInvertNeg;
			protected ushort m_reserved;
			protected ushort m_icvFore;
			protected ushort m_icvBack;
			public AreaFormatRecord(uint rgbFore, uint rgbBack, ushort fls, bool fAuto, bool fInvertNeg, ushort icvFore, ushort icvBack) : base(TYPE, SIZE)
			{
				SetDefaults();
				m_rgbFore = rgbFore;
				m_rgbBack = rgbBack;
				m_fls = fls;
				m_fAuto = 0x0;
				if (fAuto)
					m_fAuto = 0x1;
				m_fInvertNeg = 0x0;
				if (fInvertNeg)
					m_fInvertNeg = 0x1;
				m_icvFore = icvFore;
				m_icvBack = icvBack;
			}

			public AreaFormatRecord(Fill pFill) : base(TYPE, SIZE)
			{
				if (pFill.GetType() == Fill.Type.TYPE_NONE)
				{
					m_rgbFore = 0x00ffffff;
					m_rgbBack = 0x00000000;
					m_fls = 0x0000;
					m_icvFore = 0x004e;
					m_icvBack = 0x004d;
				}
				else
				{
					m_icvFore = BiffWorkbookGlobals.SnapToPalette(pFill.GetForegroundColor());
					m_rgbFore = BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(m_icvFore) & 0xFFFFFF;
					m_icvBack = BiffWorkbookGlobals.SnapToPalette(pFill.GetBackgroundColor());
					m_rgbBack = BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(m_icvBack) & 0xFFFFFF;
					m_fls = 0x0001;
				}
				m_fAuto = 0;
				m_fInvertNeg = 0;
				m_reserved = 0;
			}

			public void ModifyFill(Fill pFill, BiffWorkbookGlobals pBiffWorkbookGlobals)
			{
				if (m_fAuto == 1 && m_fls == 0x0000)
					return;
				switch (m_fls)
				{
					case 0x0000:
					{
						pFill.SetType(Fill.Type.TYPE_NONE);
						break;
					}

					default:
					{
						pFill.SetType(Fill.Type.TYPE_SOLID);
						break;
					}

				}
				pFill.GetForegroundColor().Set((byte)(m_rgbFore & 0xFF), (byte)((m_rgbFore >> 8) & 0xFF), (byte)((m_rgbFore >> 16) & 0xFF));
				pFill.GetBackgroundColor().Set((byte)(m_rgbBack & 0xFF), (byte)((m_rgbBack >> 8) & 0xFF), (byte)((m_rgbBack >> 16) & 0xFF));
				if (m_icvFore != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_FOREGROUND && m_icvFore != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_BACKGROUND)
					pFill.GetForegroundColor().SetFromRgba(pBiffWorkbookGlobals.GetPaletteColorByIndex(m_icvFore));
				if (m_icvBack != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_FOREGROUND && m_icvBack != BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_CHART_BACKGROUND)
					pFill.GetBackgroundColor().SetFromRgba(pBiffWorkbookGlobals.GetPaletteColorByIndex(m_icvBack));
			}

			public void PackForChecksum(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_rgbFore);
				pBlobView.PackUint32(m_rgbBack);
				pBlobView.PackUint8((byte)(m_fls));
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XTIStruct : BiffStruct
		{
			public XTIStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_iSupBook = pBlobView.UnpackUint16();
				m_itabFirst = pBlobView.UnpackInt16();
				m_itabLast = pBlobView.UnpackInt16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_iSupBook);
				pBlobView.PackInt16(m_itabFirst);
				pBlobView.PackInt16(m_itabLast);
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_iSupBook = 0;
				m_itabFirst = 0;
				m_itabLast = 0;
			}

			public ushort m_iSupBook;
			public short m_itabFirst;
			public short m_itabLast;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeStringStruct : BiffStruct
		{
			public XLUnicodeStringStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cch = pBlobView.UnpackUint16();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_fHighByte = (byte)((nBitmask0 >> 0) & 0x1);
				m_reserved = (byte)((nBitmask0 >> 1) & 0x7f);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint16(m_cch);
				int nBitmask0 = 0;
				nBitmask0 += m_fHighByte << 0;
				nBitmask0 += m_reserved << 1;
				pBlobView.PackUint8((byte)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 3;
			public void SetDefaults()
			{
				m_cch = 0;
				m_fHighByte = 0;
				m_reserved = 0;
				PostSetDefaults();
			}

			public ushort m_cch;
			public byte m_fHighByte;
			public byte m_reserved;
			public InternalString m_rgb;
			protected void PostSetDefaults()
			{
				m_rgb = new InternalString("");
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_fHighByte == 0)
				{
					for (ushort i = 0; i < m_cch; i++)
						m_rgb.AppendChar((char)(pBlobView.UnpackUint8()));
				}
				else
				{
					for (ushort i = 0; i < m_cch; i++)
						m_rgb.AppendChar((char)(pBlobView.UnpackUint16()));
				}
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				m_cch = (ushort)(m_rgb.GetLength());
				if (m_rgb.IsAscii())
					m_fHighByte = 0x0;
				else
					m_fHighByte = 0x1;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_fHighByte > 0)
					m_rgb.BlobWrite16Bit(pBlobView, false);
				else
					m_rgb.BlobWriteUtf8(pBlobView, false);
			}

			public int GetDynamicSize()
			{
				int nSize = m_rgb.GetLength();
				if (!m_rgb.IsAscii())
					nSize = nSize * 2;
				return nSize;
			}

			~XLUnicodeStringStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeRichExtendedString_ContinueInfo
		{
			public int m_nOffset;
			public byte m_fHighByte;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XLUnicodeRichExtendedString : BiffStruct
		{
			public XLUnicodeRichExtendedString()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cch = pBlobView.UnpackUint16();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_fHighByte = (byte)((nBitmask0 >> 0) & 0x1);
				m_reserved1 = (byte)((nBitmask0 >> 1) & 0x1);
				m_fExtSt = (byte)((nBitmask0 >> 2) & 0x1);
				m_fRichSt = (byte)((nBitmask0 >> 3) & 0x1);
				m_reserved2 = (byte)((nBitmask0 >> 4) & 0xf);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint16(m_cch);
				int nBitmask0 = 0;
				nBitmask0 += m_fHighByte << 0;
				nBitmask0 += m_reserved1 << 1;
				nBitmask0 += m_fExtSt << 2;
				nBitmask0 += m_fRichSt << 3;
				nBitmask0 += m_reserved2 << 4;
				pBlobView.PackUint8((byte)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 3;
			public void SetDefaults()
			{
				m_cch = 0;
				m_fHighByte = 0;
				m_reserved1 = 0;
				m_fExtSt = 0;
				m_fRichSt = 0;
				m_reserved2 = 0;
				PostSetDefaults();
			}

			public ushort m_cch;
			public byte m_fHighByte;
			public byte m_reserved1;
			public byte m_fExtSt;
			public byte m_fRichSt;
			public byte m_reserved2;
			public ushort m_cRun;
			public int m_cbExtRst;
			public InternalString m_rgb;
			public OwnedVector<XLUnicodeRichExtendedString_ContinueInfo> m_pContinueInfoVector;
			public XLUnicodeRichExtendedString(string szString)
			{
				SetDefaults();
				m_rgb.Set(szString);
				m_cch = (ushort)(m_rgb.GetLength());
				m_fHighByte = m_rgb.IsAscii() ? (byte)(0x0) : (byte)(0x1);
			}

			protected void PostSetDefaults()
			{
				m_cRun = 0;
				m_cbExtRst = 0;
				m_rgb = new InternalString("");
				m_pContinueInfoVector = new OwnedVector<XLUnicodeRichExtendedString_ContinueInfo>();
			}

			public void ContinueAwareBlobRead(BlobView pBlobView, OwnedVector<BiffRecord_ContinueInfo> pContinueInfoVector)
			{
				BlobRead(pBlobView);
				int nStringLength = 0;
				byte fHighByte = m_fHighByte;
				int nNextContinueIndex = 0;
				while (true)
				{
					if (nNextContinueIndex >= pContinueInfoVector.GetSize())
						break;
					int nNextContinueOffset = pContinueInfoVector.Get(nNextContinueIndex).m_nOffset;
					if (nNextContinueOffset >= pBlobView.GetOffset())
						break;
					nNextContinueIndex++;
				}
				while (nStringLength < m_cch)
				{
					if (pContinueInfoVector.GetSize() > nNextContinueIndex)
					{
						uint nCurrentOffset = (uint)(pBlobView.GetOffset());
						uint nNextContinueOffset = (uint)(pContinueInfoVector.Get(nNextContinueIndex).m_nOffset);
						if (nCurrentOffset == nNextContinueOffset)
						{
							if (nStringLength < m_cch)
							{
								XLUnicodeRichExtendedString_ContinueInfo pContinueInfo = new XLUnicodeRichExtendedString_ContinueInfo();
								pContinueInfo.m_nOffset = pBlobView.GetOffset();
								pContinueInfo.m_fHighByte = pBlobView.UnpackUint8();
								fHighByte = pContinueInfo.m_fHighByte;
								{
									NumberDuck.Secret.XLUnicodeRichExtendedString_ContinueInfo __2394809829 = pContinueInfo;
									pContinueInfo = null;
									m_pContinueInfoVector.PushBack(__2394809829);
								}
								nNextContinueIndex++;
							}
						}
					}
					int nReadSize = m_cch - nStringLength;
					if (fHighByte == 1)
						nReadSize = nReadSize << 1;
					{
						int nCurrentOffset = pBlobView.GetOffset();
						int nNextContinueOffset = pBlobView.GetSize();
						while (pContinueInfoVector.GetSize() > nNextContinueIndex)
						{
							nNextContinueOffset = pContinueInfoVector.Get(nNextContinueIndex).m_nOffset;
							if (nNextContinueOffset > nCurrentOffset)
								break;
							nNextContinueIndex++;
							if (nNextContinueIndex >= pContinueInfoVector.GetSize())
								nNextContinueOffset = pBlobView.GetSize();
						}
						if (nNextContinueOffset < nCurrentOffset + nReadSize)
							nReadSize = nNextContinueOffset - nCurrentOffset;
					}
					if (fHighByte == 0)
					{
						for (int j = 0; j < nReadSize; j++)
							m_rgb.AppendChar((char)(pBlobView.UnpackUint8()));
						nStringLength += nReadSize;
					}
					else
					{
						for (int j = 0; j < nReadSize >> 1; j++)
							m_rgb.AppendChar((char)(pBlobView.UnpackUint16()));
						nStringLength += nReadSize >> 1;
					}
				}
				if (m_fRichSt == 0x1)
				{
					for (ushort i = 0; i < m_cRun; i++)
					{
						pBlobView.SetOffset(pBlobView.GetOffset() + 4);
					}
				}
				if (m_fExtSt == 0x1)
				{
					pBlobView.SetOffset(pBlobView.GetOffset() + m_cbExtRst);
				}
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_fRichSt == 0x1)
					m_cRun = pBlobView.UnpackUint16();
				if (m_fExtSt == 0x1)
					m_cbExtRst = pBlobView.UnpackInt32();
			}

			public void ContinueAwareBlobWrite(BlobView pBlobView, OwnedVector<BiffRecord_ContinueInfo> pContinueInfoVector)
			{
				int nOffset = pBlobView.GetOffset() + 8;
				int nContinueOffset = nOffset;
				if (pContinueInfoVector.GetSize() > 0)
					nContinueOffset -= pContinueInfoVector.Get(pContinueInfoVector.GetSize() - 1).m_nOffset;
				if (nContinueOffset + SIZE > BiffRecord.MAX_DATA_SIZE)
				{
					pContinueInfoVector.PushBack(new BiffRecord_ContinueInfo(nOffset, 0));
					nbAssert.Assert(nContinueOffset <= BiffRecord.MAX_DATA_SIZE);
					nContinueOffset = 0;
				}
				BlobWrite(pBlobView);
				nOffset = nOffset + (int)(SIZE);
				nContinueOffset = nContinueOffset + (int)(SIZE);
				nbAssert.Assert(m_cch == m_rgb.GetLength());
				if (m_cch > 0)
				{
					Blob pDataBlob = new Blob(true);
					BlobView pDataBlobView = pDataBlob.GetBlobView();
					if (m_fHighByte > 0)
						m_rgb.BlobWrite16Bit(pDataBlobView, false);
					else
						m_rgb.BlobWriteUtf8(pDataBlobView, false);
					pDataBlobView.SetOffset(0);
					while (pDataBlobView.GetOffset() != pDataBlobView.GetSize())
					{
						if (nContinueOffset == 0)
						{
							pBlobView.PackUint8(m_fHighByte);
							nOffset += 1;
							nContinueOffset += 1;
						}
						int nWriteSize = pDataBlobView.GetSize() - pDataBlobView.GetOffset();
						if (nContinueOffset + nWriteSize > BiffRecord.MAX_DATA_SIZE)
						{
							nWriteSize = BiffRecord.MAX_DATA_SIZE - nContinueOffset;
							if (m_fHighByte > 0 && nWriteSize % 2 > 0)
								nWriteSize--;
						}
						if (nWriteSize > 0)
						{
							pBlobView.Pack(pDataBlobView, nWriteSize);
							nOffset += nWriteSize;
							nContinueOffset += nWriteSize;
						}
						if (pDataBlobView.GetOffset() < pDataBlobView.GetSize())
						{
							pContinueInfoVector.PushBack(new BiffRecord_ContinueInfo(nOffset, 0));
							nbAssert.Assert(nContinueOffset <= BiffRecord.MAX_DATA_SIZE);
							nContinueOffset = 0;
						}
					}
					{
						pDataBlob = null;
					}
				}
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
			}

			~XLUnicodeRichExtendedString()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ShortXLUnicodeStringStruct : BiffStruct
		{
			public ShortXLUnicodeStringStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cch = pBlobView.UnpackUint8();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_fHighByte = (byte)((nBitmask0 >> 0) & 0x1);
				m_reserved = (byte)((nBitmask0 >> 1) & 0x7f);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint8(m_cch);
				int nBitmask0 = 0;
				nBitmask0 += m_fHighByte << 0;
				nBitmask0 += m_reserved << 1;
				pBlobView.PackUint8((byte)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_cch = 0;
				m_fHighByte = 1;
				m_reserved = 0;
				PostSetDefaults();
			}

			public byte m_cch;
			public byte m_fHighByte;
			public byte m_reserved;
			public InternalString m_rgb;
			protected void PostSetDefaults()
			{
				m_rgb = new InternalString("");
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_fHighByte == 0)
				{
					for (ushort i = 0; i < m_cch; i++)
						m_rgb.AppendChar((char)(pBlobView.UnpackUint8()));
				}
				else
				{
					for (ushort i = 0; i < m_cch; i++)
						m_rgb.AppendChar((char)(pBlobView.UnpackUint16()));
				}
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				m_cch = (byte)(m_rgb.GetLength());
				if (m_rgb.IsAscii())
					m_fHighByte = 0x0;
				else
					m_fHighByte = 0x1;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_fHighByte > 0)
					m_rgb.BlobWrite16Bit(pBlobView, false);
				else
					m_rgb.BlobWriteUtf8(pBlobView, false);
			}

			public int GetDynamicSize()
			{
				int nSize = m_rgb.GetLength();
				if (!m_rgb.IsAscii())
					nSize = nSize * 2;
				return nSize;
			}

			~ShortXLUnicodeStringStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RwUStruct : BiffStruct
		{
			public RwUStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_rw);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_rw = 0;
			}

			public ushort m_rw;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RwStruct : BiffStruct
		{
			public RwStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_rw);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_rw = 0;
			}

			public ushort m_rw;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RkRecStruct : BiffStruct
		{
			public RkRecStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ixfe = pBlobView.UnpackUint16();
				m_RK = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ixfe);
				pBlobView.PackUint32(m_RK);
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_ixfe = 0;
				m_RK = 0;
			}

			public ushort m_ixfe;
			public uint m_RK;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceStruct : BiffStruct
		{
			public RgceStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 0;
			public void SetDefaults()
			{
				PostSetDefaults();
			}

			public OwnedVector<ParsedExpressionRecord> m_pParsedExpressionRecordVector;
			protected void PostSetDefaults()
			{
				m_pParsedExpressionRecordVector = new OwnedVector<ParsedExpressionRecord>();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				while (pBlobView.GetOffset() < pBlobView.GetSize())
				{
					ParsedExpressionRecord pParsedExpressionRecord = ParsedExpressionRecord.CreateParsedExpressionRecord(pBlobView);
					ParsedExpressionRecord.Type eTemp = pParsedExpressionRecord.GetType();
					{
						NumberDuck.Secret.ParsedExpressionRecord __3596419756 = pParsedExpressionRecord;
						pParsedExpressionRecord = null;
						m_pParsedExpressionRecordVector.PushBack(__3596419756);
					}
					if (eTemp == ParsedExpressionRecord.Type.TYPE_UNKNOWN)
					{
						break;
					}
				}
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				int i;
				for (i = 0; i < m_pParsedExpressionRecordVector.GetSize(); i++)
				{
					ParsedExpressionRecord pParsedExpressionRecord = m_pParsedExpressionRecordVector.Get(i);
					pParsedExpressionRecord.BlobWrite(pBlobView);
				}
			}

			public int GetSize()
			{
				int i;
				int nSize = 0;
				for (i = 0; i < m_pParsedExpressionRecordVector.GetSize(); i++)
				{
					ParsedExpressionRecord pParsedExpressionRecord = m_pParsedExpressionRecordVector.Get(i);
					nSize = nSize + (int)(pParsedExpressionRecord.GetDataSize());
				}
				return nSize;
			}

			~RgceStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceLocStruct : BiffStruct
		{
			public RgceLocStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_row.BlobRead(pBlobView);
				m_column.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_row.BlobWrite(pBlobView);
				m_column.BlobWrite(pBlobView);
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_row = new RwUStruct();
				m_column = new ColRelUStruct();
			}

			public RwUStruct m_row;
			public ColRelUStruct m_column;
			~RgceLocStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RgceAreaStruct : BiffStruct
		{
			public RgceAreaStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rowFirst.BlobRead(pBlobView);
				m_rowLast.BlobRead(pBlobView);
				m_columnFirst.BlobRead(pBlobView);
				m_columnLast.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rowFirst.BlobWrite(pBlobView);
				m_rowLast.BlobWrite(pBlobView);
				m_columnFirst.BlobWrite(pBlobView);
				m_columnLast.BlobWrite(pBlobView);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_rowFirst = new RwUStruct();
				m_rowLast = new RwUStruct();
				m_columnFirst = new ColRelUStruct();
				m_columnLast = new ColRelUStruct();
			}

			public RwUStruct m_rowFirst;
			public RwUStruct m_rowLast;
			public ColRelUStruct m_columnFirst;
			public ColRelUStruct m_columnLast;
			~RgceAreaStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Ref8Struct : BiffStruct
		{
			public Ref8Struct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rwFirst = pBlobView.UnpackUint16();
				m_rwLast = pBlobView.UnpackUint16();
				m_colFirst = pBlobView.UnpackUint16();
				m_colLast = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_rwFirst);
				pBlobView.PackUint16(m_rwLast);
				pBlobView.PackUint16(m_colFirst);
				pBlobView.PackUint16(m_colLast);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_rwFirst = 0;
				m_rwLast = 0;
				m_colFirst = 0;
				m_colLast = 0;
			}

			public ushort m_rwFirst;
			public ushort m_rwLast;
			public ushort m_colFirst;
			public ushort m_colLast;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PtgAttrSpaceTypeStruct : BiffStruct
		{
			public PtgAttrSpaceTypeStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_type = pBlobView.UnpackUint8();
				m_cch = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_type);
				pBlobView.PackUint8(m_cch);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_type = 0;
				m_cch = 0;
			}

			public byte m_type;
			public byte m_cch;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtRecordHeaderStruct : BiffStruct
		{
			public OfficeArtRecordHeaderStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_recVer = (ushort)((nBitmask0 >> 0) & 0xf);
				m_recInstance = (ushort)((nBitmask0 >> 4) & 0xfff);
				m_recType = pBlobView.UnpackUint16();
				m_recLen = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_recVer << 0;
				nBitmask0 += m_recInstance << 4;
				pBlobView.PackUint16((ushort)(nBitmask0));
				pBlobView.PackUint16(m_recType);
				pBlobView.PackUint32(m_recLen);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_recVer = 0;
				m_recInstance = 0;
				m_recType = 0;
				m_recLen = 0;
			}

			public ushort m_recVer;
			public ushort m_recInstance;
			public ushort m_recType;
			public uint m_recLen;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtIDCLStruct : BiffStruct
		{
			public OfficeArtIDCLStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_dgid = pBlobView.UnpackUint32();
				m_cspidCur = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_dgid);
				pBlobView.PackUint32(m_cspidCur);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_dgid = 0;
				m_cspidCur = 0;
			}

			public uint m_dgid;
			public uint m_cspidCur;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFRITStruct : BiffStruct
		{
			public OfficeArtFRITStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_fridNew = pBlobView.UnpackUint16();
				m_fridOld = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_fridNew);
				pBlobView.PackUint16(m_fridOld);
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_fridNew = 0;
				m_fridOld = 0;
			}

			public ushort m_fridNew;
			public ushort m_fridOld;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTEStruct : BiffStruct
		{
			public OfficeArtFOPTEStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_opid.BlobRead(pBlobView);
				m_op = pBlobView.UnpackInt32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_opid.BlobWrite(pBlobView);
				pBlobView.PackInt32(m_op);
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_opid = new OfficeArtFOPTEOPIDStruct();
				m_op = 0;
				PostSetDefaults();
			}

			public OfficeArtFOPTEOPIDStruct m_opid;
			public int m_op;
			public Blob m_pComplexData;
			protected void PostSetDefaults()
			{
				m_pComplexData = null;
			}

			~OfficeArtFOPTEStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtFOPTEOPIDStruct : BiffStruct
		{
			public OfficeArtFOPTEOPIDStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_opid = (ushort)((nBitmask0 >> 0) & 0x3fff);
				m_fBid = (ushort)((nBitmask0 >> 14) & 0x1);
				m_fComplex = (ushort)((nBitmask0 >> 15) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_opid << 0;
				nBitmask0 += m_fBid << 14;
				nBitmask0 += m_fComplex << 15;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_opid = 0;
				m_fBid = 0;
				m_fComplex = 0;
			}

			public ushort m_opid;
			public ushort m_fBid;
			public ushort m_fComplex;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MSOCRStruct : BiffStruct
		{
			public MSOCRStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_red = pBlobView.UnpackUint8();
				m_green = pBlobView.UnpackUint8();
				m_blue = pBlobView.UnpackUint8();
				byte nBitmask0 = pBlobView.UnpackUint8();
				m_unused1 = (byte)((nBitmask0 >> 0) & 0x7);
				m_fSchemeIndex = (byte)((nBitmask0 >> 3) & 0x1);
				m_unused2 = (byte)((nBitmask0 >> 4) & 0xf);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_red);
				pBlobView.PackUint8(m_green);
				pBlobView.PackUint8(m_blue);
				int nBitmask0 = 0;
				nBitmask0 += m_unused1 << 0;
				nBitmask0 += m_fSchemeIndex << 3;
				nBitmask0 += m_unused2 << 4;
				pBlobView.PackUint8((byte)(nBitmask0));
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_red = 0;
				m_green = 0;
				m_blue = 0;
				m_unused1 = 0;
				m_fSchemeIndex = 0;
				m_unused2 = 0;
			}

			public byte m_red;
			public byte m_green;
			public byte m_blue;
			public byte m_unused1;
			public byte m_fSchemeIndex;
			public byte m_unused2;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MD4DigestStruct : BiffStruct
		{
			public MD4DigestStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rgbUid1_0 = pBlobView.UnpackUint8();
				m_rgbUid1_1 = pBlobView.UnpackUint8();
				m_rgbUid1_2 = pBlobView.UnpackUint8();
				m_rgbUid1_3 = pBlobView.UnpackUint8();
				m_rgbUid1_4 = pBlobView.UnpackUint8();
				m_rgbUid1_5 = pBlobView.UnpackUint8();
				m_rgbUid1_6 = pBlobView.UnpackUint8();
				m_rgbUid1_7 = pBlobView.UnpackUint8();
				m_rgbUid1_8 = pBlobView.UnpackUint8();
				m_rgbUid1_9 = pBlobView.UnpackUint8();
				m_rgbUid1_10 = pBlobView.UnpackUint8();
				m_rgbUid1_11 = pBlobView.UnpackUint8();
				m_rgbUid1_12 = pBlobView.UnpackUint8();
				m_rgbUid1_13 = pBlobView.UnpackUint8();
				m_rgbUid1_14 = pBlobView.UnpackUint8();
				m_rgbUid1_15 = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_rgbUid1_0);
				pBlobView.PackUint8(m_rgbUid1_1);
				pBlobView.PackUint8(m_rgbUid1_2);
				pBlobView.PackUint8(m_rgbUid1_3);
				pBlobView.PackUint8(m_rgbUid1_4);
				pBlobView.PackUint8(m_rgbUid1_5);
				pBlobView.PackUint8(m_rgbUid1_6);
				pBlobView.PackUint8(m_rgbUid1_7);
				pBlobView.PackUint8(m_rgbUid1_8);
				pBlobView.PackUint8(m_rgbUid1_9);
				pBlobView.PackUint8(m_rgbUid1_10);
				pBlobView.PackUint8(m_rgbUid1_11);
				pBlobView.PackUint8(m_rgbUid1_12);
				pBlobView.PackUint8(m_rgbUid1_13);
				pBlobView.PackUint8(m_rgbUid1_14);
				pBlobView.PackUint8(m_rgbUid1_15);
			}

			public const ushort SIZE = 16;
			public void SetDefaults()
			{
				m_rgbUid1_0 = 0;
				m_rgbUid1_1 = 0;
				m_rgbUid1_2 = 0;
				m_rgbUid1_3 = 0;
				m_rgbUid1_4 = 0;
				m_rgbUid1_5 = 0;
				m_rgbUid1_6 = 0;
				m_rgbUid1_7 = 0;
				m_rgbUid1_8 = 0;
				m_rgbUid1_9 = 0;
				m_rgbUid1_10 = 0;
				m_rgbUid1_11 = 0;
				m_rgbUid1_12 = 0;
				m_rgbUid1_13 = 0;
				m_rgbUid1_14 = 0;
				m_rgbUid1_15 = 0;
			}

			public byte m_rgbUid1_0;
			public byte m_rgbUid1_1;
			public byte m_rgbUid1_2;
			public byte m_rgbUid1_3;
			public byte m_rgbUid1_4;
			public byte m_rgbUid1_5;
			public byte m_rgbUid1_6;
			public byte m_rgbUid1_7;
			public byte m_rgbUid1_8;
			public byte m_rgbUid1_9;
			public byte m_rgbUid1_10;
			public byte m_rgbUid1_11;
			public byte m_rgbUid1_12;
			public byte m_rgbUid1_13;
			public byte m_rgbUid1_14;
			public byte m_rgbUid1_15;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IcvFontStruct : BiffStruct
		{
			public IcvFontStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_icv = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_icv);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_icv = 0;
			}

			public ushort m_icv;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IXFCellStruct : BiffStruct
		{
			public IXFCellStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ixfe = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ixfe);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_ixfe = 0;
			}

			public ushort m_ixfe;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IHlinkStruct : BiffStruct
		{
			public IHlinkStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_CLSID_StdHlink_0 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_1 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_2 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_3 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_4 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_5 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_6 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_7 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_8 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_9 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_10 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_11 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_12 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_13 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_14 = pBlobView.UnpackUint8();
				m_CLSID_StdHlink_15 = pBlobView.UnpackUint8();
				m_hyperlink.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_CLSID_StdHlink_0);
				pBlobView.PackUint8(m_CLSID_StdHlink_1);
				pBlobView.PackUint8(m_CLSID_StdHlink_2);
				pBlobView.PackUint8(m_CLSID_StdHlink_3);
				pBlobView.PackUint8(m_CLSID_StdHlink_4);
				pBlobView.PackUint8(m_CLSID_StdHlink_5);
				pBlobView.PackUint8(m_CLSID_StdHlink_6);
				pBlobView.PackUint8(m_CLSID_StdHlink_7);
				pBlobView.PackUint8(m_CLSID_StdHlink_8);
				pBlobView.PackUint8(m_CLSID_StdHlink_9);
				pBlobView.PackUint8(m_CLSID_StdHlink_10);
				pBlobView.PackUint8(m_CLSID_StdHlink_11);
				pBlobView.PackUint8(m_CLSID_StdHlink_12);
				pBlobView.PackUint8(m_CLSID_StdHlink_13);
				pBlobView.PackUint8(m_CLSID_StdHlink_14);
				pBlobView.PackUint8(m_CLSID_StdHlink_15);
				m_hyperlink.BlobWrite(pBlobView);
			}

			public const ushort SIZE = 24;
			public void SetDefaults()
			{
				m_CLSID_StdHlink_0 = 0xD0;
				m_CLSID_StdHlink_1 = 0xC9;
				m_CLSID_StdHlink_2 = 0xEA;
				m_CLSID_StdHlink_3 = 0x79;
				m_CLSID_StdHlink_4 = 0xF9;
				m_CLSID_StdHlink_5 = 0xBA;
				m_CLSID_StdHlink_6 = 0xCE;
				m_CLSID_StdHlink_7 = 0x11;
				m_CLSID_StdHlink_8 = 0x8C;
				m_CLSID_StdHlink_9 = 0x82;
				m_CLSID_StdHlink_10 = 0x00;
				m_CLSID_StdHlink_11 = 0xAA;
				m_CLSID_StdHlink_12 = 0x00;
				m_CLSID_StdHlink_13 = 0x4B;
				m_CLSID_StdHlink_14 = 0xA9;
				m_CLSID_StdHlink_15 = 0x0B;
				m_hyperlink = new HyperlinkObjectStruct();
			}

			public byte m_CLSID_StdHlink_0;
			public byte m_CLSID_StdHlink_1;
			public byte m_CLSID_StdHlink_2;
			public byte m_CLSID_StdHlink_3;
			public byte m_CLSID_StdHlink_4;
			public byte m_CLSID_StdHlink_5;
			public byte m_CLSID_StdHlink_6;
			public byte m_CLSID_StdHlink_7;
			public byte m_CLSID_StdHlink_8;
			public byte m_CLSID_StdHlink_9;
			public byte m_CLSID_StdHlink_10;
			public byte m_CLSID_StdHlink_11;
			public byte m_CLSID_StdHlink_12;
			public byte m_CLSID_StdHlink_13;
			public byte m_CLSID_StdHlink_14;
			public byte m_CLSID_StdHlink_15;
			public HyperlinkObjectStruct m_hyperlink;
			~IHlinkStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HyperlinkStringStruct : BiffStruct
		{
			public HyperlinkStringStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_length = pBlobView.UnpackUint32();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint32(m_length);
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_length = 0;
				PostSetDefaults();
			}

			public uint m_length;
			protected InternalString m_string;
			protected void PostSetDefaults()
			{
				m_string = new InternalString("");
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				for (uint i = 0; i < m_length; i++)
					m_string.AppendChar((char)(pBlobView.UnpackUint16()));
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				m_length = (uint)(m_string.GetLength());
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_string.BlobWrite16Bit(pBlobView, false);
			}

			~HyperlinkStringStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class HyperlinkObjectStruct : BiffStruct
		{
			public HyperlinkObjectStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_streamVersion = pBlobView.UnpackUint32();
				uint nBitmask0 = pBlobView.UnpackUint32();
				m_hlstmfHasMoniker = (uint)((nBitmask0 >> 0) & 0x1);
				m_hlstmfIsAbsolute = (uint)((nBitmask0 >> 1) & 0x1);
				m_hlstmfSiteGaveDisplayName = (uint)((nBitmask0 >> 2) & 0x1);
				m_hlstmfHasLocationStr = (uint)((nBitmask0 >> 3) & 0x1);
				m_hlstmfHasDisplayName = (uint)((nBitmask0 >> 4) & 0x1);
				m_hlstmfHasGUID = (uint)((nBitmask0 >> 5) & 0x1);
				m_hlstmfHasCreationTime = (uint)((nBitmask0 >> 6) & 0x1);
				m_hlstmfHasFrameName = (uint)((nBitmask0 >> 7) & 0x1);
				m_hlstmfMonikerSavedAsStr = (uint)((nBitmask0 >> 8) & 0x1);
				m_hlstmfAbsFromGetdataRel = (uint)((nBitmask0 >> 9) & 0x1);
				m_reserved = (uint)((nBitmask0 >> 10) & 0x3fffff);
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint32(m_streamVersion);
				uint nBitmask0 = 0;
				nBitmask0 += m_hlstmfHasMoniker << 0;
				nBitmask0 += m_hlstmfIsAbsolute << 1;
				nBitmask0 += m_hlstmfSiteGaveDisplayName << 2;
				nBitmask0 += m_hlstmfHasLocationStr << 3;
				nBitmask0 += m_hlstmfHasDisplayName << 4;
				nBitmask0 += m_hlstmfHasGUID << 5;
				nBitmask0 += m_hlstmfHasCreationTime << 6;
				nBitmask0 += m_hlstmfHasFrameName << 7;
				nBitmask0 += m_hlstmfMonikerSavedAsStr << 8;
				nBitmask0 += m_hlstmfAbsFromGetdataRel << 9;
				nBitmask0 += m_reserved << 10;
				pBlobView.PackUint32((uint)(nBitmask0));
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_streamVersion = 0;
				m_hlstmfHasMoniker = 0;
				m_hlstmfIsAbsolute = 0;
				m_hlstmfSiteGaveDisplayName = 0;
				m_hlstmfHasLocationStr = 0;
				m_hlstmfHasDisplayName = 0;
				m_hlstmfHasGUID = 0;
				m_hlstmfHasCreationTime = 0;
				m_hlstmfHasFrameName = 0;
				m_hlstmfMonikerSavedAsStr = 0;
				m_hlstmfAbsFromGetdataRel = 0;
				m_reserved = 0;
				PostSetDefaults();
			}

			public uint m_streamVersion;
			public uint m_hlstmfHasMoniker;
			public uint m_hlstmfIsAbsolute;
			public uint m_hlstmfSiteGaveDisplayName;
			public uint m_hlstmfHasLocationStr;
			public uint m_hlstmfHasDisplayName;
			public uint m_hlstmfHasGUID;
			public uint m_hlstmfHasCreationTime;
			public uint m_hlstmfHasFrameName;
			public uint m_hlstmfMonikerSavedAsStr;
			public uint m_hlstmfAbsFromGetdataRel;
			public uint m_reserved;
			protected HyperlinkStringStruct m_displayName;
			protected HyperlinkStringStruct m_targetFrameName;
			protected HyperlinkStringStruct m_moniker;
			public InternalString m_haxUrl;
			protected void PostSetDefaults()
			{
				m_displayName = null;
				m_targetFrameName = null;
				m_moniker = null;
				m_haxUrl = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				if (m_hlstmfHasDisplayName > 0)
				{
					m_displayName = new HyperlinkStringStruct();
					m_displayName.BlobRead(pBlobView);
				}
				if (m_hlstmfHasFrameName > 0)
				{
					m_targetFrameName = new HyperlinkStringStruct();
					m_targetFrameName.BlobRead(pBlobView);
				}
				if (m_hlstmfHasMoniker > 0 && m_hlstmfMonikerSavedAsStr > 0)
				{
					m_moniker = new HyperlinkStringStruct();
					m_moniker.BlobRead(pBlobView);
				}
				if (m_hlstmfHasMoniker > 0 && m_hlstmfMonikerSavedAsStr == 0)
					if (pBlobView.UnpackUint8() == 0xE0 && pBlobView.UnpackUint8() == 0xC9 && pBlobView.UnpackUint8() == 0xEA && pBlobView.UnpackUint8() == 0x79 && pBlobView.UnpackUint8() == 0xF9 && pBlobView.UnpackUint8() == 0xBA && pBlobView.UnpackUint8() == 0xCE && pBlobView.UnpackUint8() == 0x11 && pBlobView.UnpackUint8() == 0x8C && pBlobView.UnpackUint8() == 0x82 && pBlobView.UnpackUint8() == 0x00 && pBlobView.UnpackUint8() == 0xAA && pBlobView.UnpackUint8() == 0x00 && pBlobView.UnpackUint8() == 0x4B && pBlobView.UnpackUint8() == 0xA9 && pBlobView.UnpackUint8() == 0x0B)
					{
						uint nLength = pBlobView.UnpackUint32();
						nLength = nLength >> 1;
						m_haxUrl = new InternalString("");
						for (uint i = 0; i < nLength; i++)
							m_haxUrl.AppendChar((char)(pBlobView.UnpackUint16()));
					}
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(m_haxUrl != null);
				nbAssert.Assert(m_hlstmfHasMoniker > 0 && m_hlstmfMonikerSavedAsStr == 0);
				pBlobView.PackUint8(0xE0);
				pBlobView.PackUint8(0xC9);
				pBlobView.PackUint8(0xEA);
				pBlobView.PackUint8(0x79);
				pBlobView.PackUint8(0xF9);
				pBlobView.PackUint8(0xBA);
				pBlobView.PackUint8(0xCE);
				pBlobView.PackUint8(0x11);
				pBlobView.PackUint8(0x8C);
				pBlobView.PackUint8(0x82);
				pBlobView.PackUint8(0x00);
				pBlobView.PackUint8(0xAA);
				pBlobView.PackUint8(0x00);
				pBlobView.PackUint8(0x4B);
				pBlobView.PackUint8(0xA9);
				pBlobView.PackUint8(0x0B);
				uint URLMoniker_nSize = (uint)((m_haxUrl.GetLength() + 1) * 2 + 16 + 4 + 4);
				pBlobView.PackUint32(URLMoniker_nSize);
				m_haxUrl.BlobWrite16Bit(pBlobView, true);
				pBlobView.PackUint8(0x79);
				pBlobView.PackUint8(0x58);
				pBlobView.PackUint8(0x81);
				pBlobView.PackUint8(0xF4);
				pBlobView.PackUint8(0x3B);
				pBlobView.PackUint8(0x1D);
				pBlobView.PackUint8(0x7F);
				pBlobView.PackUint8(0x48);
				pBlobView.PackUint8(0xAF);
				pBlobView.PackUint8(0x2C);
				pBlobView.PackUint8(0x82);
				pBlobView.PackUint8(0x5D);
				pBlobView.PackUint8(0xC4);
				pBlobView.PackUint8(0x85);
				pBlobView.PackUint8(0x27);
				pBlobView.PackUint8(0x63);
				pBlobView.PackUint32(0);
				pBlobView.PackUint32(43941);
			}

			~HyperlinkObjectStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FullColorExtStruct : BiffStruct
		{
			public FullColorExtStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_xclrType = pBlobView.UnpackUint16();
				m_nTintShade = pBlobView.UnpackInt16();
				m_xclrValue = pBlobView.UnpackUint32();
				m_unusedA = pBlobView.UnpackUint8();
				m_unusedB = pBlobView.UnpackUint8();
				m_unusedC = pBlobView.UnpackUint8();
				m_unusedD = pBlobView.UnpackUint8();
				m_unusedE = pBlobView.UnpackUint8();
				m_unusedF = pBlobView.UnpackUint8();
				m_unusedG = pBlobView.UnpackUint8();
				m_unusedH = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_xclrType);
				pBlobView.PackInt16(m_nTintShade);
				pBlobView.PackUint32(m_xclrValue);
				pBlobView.PackUint8(m_unusedA);
				pBlobView.PackUint8(m_unusedB);
				pBlobView.PackUint8(m_unusedC);
				pBlobView.PackUint8(m_unusedD);
				pBlobView.PackUint8(m_unusedE);
				pBlobView.PackUint8(m_unusedF);
				pBlobView.PackUint8(m_unusedG);
				pBlobView.PackUint8(m_unusedH);
			}

			public const ushort SIZE = 16;
			public void SetDefaults()
			{
				m_xclrType = 0;
				m_nTintShade = 0;
				m_xclrValue = 0;
				m_unusedA = 0;
				m_unusedB = 0;
				m_unusedC = 0;
				m_unusedD = 0;
				m_unusedE = 0;
				m_unusedF = 0;
				m_unusedG = 0;
				m_unusedH = 0;
			}

			public ushort m_xclrType;
			public short m_nTintShade;
			public uint m_xclrValue;
			public byte m_unusedA;
			public byte m_unusedB;
			public byte m_unusedC;
			public byte m_unusedD;
			public byte m_unusedE;
			public byte m_unusedF;
			public byte m_unusedG;
			public byte m_unusedH;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtPioGrbitStruct : BiffStruct
		{
			public FtPioGrbitStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ft = pBlobView.UnpackUint16();
				m_cb = pBlobView.UnpackUint16();
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fAutoPict = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fDde = (ushort)((nBitmask0 >> 1) & 0x1);
				m_fPrintCalc = (ushort)((nBitmask0 >> 2) & 0x1);
				m_fIcon = (ushort)((nBitmask0 >> 3) & 0x1);
				m_fCtl = (ushort)((nBitmask0 >> 4) & 0x1);
				m_fPrstm = (ushort)((nBitmask0 >> 5) & 0x1);
				m_unused1 = (ushort)((nBitmask0 >> 6) & 0x1);
				m_fCamera = (ushort)((nBitmask0 >> 7) & 0x1);
				m_fDefaultSize = (ushort)((nBitmask0 >> 8) & 0x1);
				m_fAutoLoad = (ushort)((nBitmask0 >> 9) & 0x1);
				m_unused2 = (ushort)((nBitmask0 >> 10) & 0x3f);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ft);
				pBlobView.PackUint16(m_cb);
				int nBitmask0 = 0;
				nBitmask0 += m_fAutoPict << 0;
				nBitmask0 += m_fDde << 1;
				nBitmask0 += m_fPrintCalc << 2;
				nBitmask0 += m_fIcon << 3;
				nBitmask0 += m_fCtl << 4;
				nBitmask0 += m_fPrstm << 5;
				nBitmask0 += m_unused1 << 6;
				nBitmask0 += m_fCamera << 7;
				nBitmask0 += m_fDefaultSize << 8;
				nBitmask0 += m_fAutoLoad << 9;
				nBitmask0 += m_unused2 << 10;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_ft = 0x0008;
				m_cb = 0x0002;
				m_fAutoPict = 0;
				m_fDde = 0;
				m_fPrintCalc = 0;
				m_fIcon = 0;
				m_fCtl = 0;
				m_fPrstm = 0;
				m_unused1 = 0;
				m_fCamera = 0;
				m_fDefaultSize = 0;
				m_fAutoLoad = 0;
				m_unused2 = 0;
			}

			public ushort m_ft;
			public ushort m_cb;
			public ushort m_fAutoPict;
			public ushort m_fDde;
			public ushort m_fPrintCalc;
			public ushort m_fIcon;
			public ushort m_fCtl;
			public ushort m_fPrstm;
			public ushort m_unused1;
			public ushort m_fCamera;
			public ushort m_fDefaultSize;
			public ushort m_fAutoLoad;
			public ushort m_unused2;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FtCfStruct : BiffStruct
		{
			public FtCfStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_ft = pBlobView.UnpackUint16();
				m_cb = pBlobView.UnpackUint16();
				m_cf = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_ft);
				pBlobView.PackUint16(m_cb);
				pBlobView.PackUint16(m_cf);
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_ft = 0x0007;
				m_cb = 0x0002;
				m_cf = 0;
			}

			public ushort m_ft;
			public ushort m_cb;
			public ushort m_cf;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtHeaderStruct : BiffStruct
		{
			public FrtHeaderStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rt = pBlobView.UnpackUint16();
				m_grbitFrt.BlobRead(pBlobView);
				m_reservedA = pBlobView.UnpackUint32();
				m_reservedB = pBlobView.UnpackUint32();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_rt);
				m_grbitFrt.BlobWrite(pBlobView);
				pBlobView.PackUint32(m_reservedA);
				pBlobView.PackUint32(m_reservedB);
			}

			public const ushort SIZE = 12;
			public void SetDefaults()
			{
				m_rt = 0;
				m_grbitFrt = new FrtFlagsStruct();
				m_reservedA = 0;
				m_reservedB = 0;
			}

			public ushort m_rt;
			public FrtFlagsStruct m_grbitFrt;
			public uint m_reservedA;
			public uint m_reservedB;
			~FrtHeaderStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtHeaderOldStruct : BiffStruct
		{
			public FrtHeaderOldStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rt = pBlobView.UnpackUint16();
				m_grbitFrt.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_rt);
				m_grbitFrt.BlobWrite(pBlobView);
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_rt = 0;
				m_grbitFrt = new FrtFlagsStruct();
			}

			public ushort m_rt;
			public FrtFlagsStruct m_grbitFrt;
			~FrtHeaderOldStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FrtFlagsStruct : BiffStruct
		{
			public FrtFlagsStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_fFrtRef = (ushort)((nBitmask0 >> 0) & 0x1);
				m_fFrtAlert = (ushort)((nBitmask0 >> 1) & 0x1);
				m_reserved = (ushort)((nBitmask0 >> 2) & 0x3fff);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_fFrtRef << 0;
				nBitmask0 += m_fFrtAlert << 1;
				nBitmask0 += m_reserved << 2;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_fFrtRef = 0;
				m_fFrtAlert = 0;
				m_reserved = 0;
			}

			public ushort m_fFrtRef;
			public ushort m_fFrtAlert;
			public ushort m_reserved;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FormulaValueStruct : BiffStruct
		{
			public FormulaValueStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_byte1 = pBlobView.UnpackUint8();
				m_byte2 = pBlobView.UnpackUint8();
				m_byte3 = pBlobView.UnpackUint8();
				m_byte4 = pBlobView.UnpackUint8();
				m_byte5 = pBlobView.UnpackUint8();
				m_byte6 = pBlobView.UnpackUint8();
				m_fExprO = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_byte1);
				pBlobView.PackUint8(m_byte2);
				pBlobView.PackUint8(m_byte3);
				pBlobView.PackUint8(m_byte4);
				pBlobView.PackUint8(m_byte5);
				pBlobView.PackUint8(m_byte6);
				pBlobView.PackUint16(m_fExprO);
			}

			public const ushort SIZE = 8;
			public void SetDefaults()
			{
				m_byte1 = 0;
				m_byte2 = 0;
				m_byte3 = 0;
				m_byte4 = 0;
				m_byte5 = 0;
				m_byte6 = 0;
				m_fExprO = 0;
			}

			public byte m_byte1;
			public byte m_byte2;
			public byte m_byte3;
			public byte m_byte4;
			public byte m_byte5;
			public byte m_byte6;
			public ushort m_fExprO;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ExtPropStruct : BiffStruct
		{
			public ExtPropStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_extType = pBlobView.UnpackUint16();
				m_cb = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint16(m_extType);
				pBlobView.PackUint16(m_cb);
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 4;
			public void SetDefaults()
			{
				m_extType = 0;
				m_cb = 0;
				PostSetDefaults();
			}

			public ushort m_extType;
			public ushort m_cb;
			public FullColorExtStruct m_pFullColorExt;
			protected void PostSetDefaults()
			{
				m_pFullColorExt = null;
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				nbAssert.Assert(m_pFullColorExt == null);
				switch (m_extType)
				{
					case 0x0004:
					case 0x0005:
					case 0x0007:
					case 0x0008:
					case 0x0009:
					case 0x000A:
					case 0x000B:
					case 0x000D:
					{
						m_pFullColorExt = new FullColorExtStruct();
						m_pFullColorExt.BlobRead(pBlobView);
						break;
					}

					default:
					{
						pBlobView.SetOffset(pBlobView.GetOffset() + m_cb - SIZE);
						break;
					}

				}
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				nbAssert.Assert(m_pFullColorExt != null);
				m_cb = SIZE + FullColorExtStruct.SIZE;
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				if (m_pFullColorExt != null)
					m_pFullColorExt.BlobWrite(pBlobView);
			}

			~ExtPropStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColStruct : BiffStruct
		{
			public ColStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_col = pBlobView.UnpackUint16();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint16(m_col);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_col = 0;
			}

			public ushort m_col;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColRelUStruct : BiffStruct
		{
			public ColRelUStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				ushort nBitmask0 = pBlobView.UnpackUint16();
				m_col = (ushort)((nBitmask0 >> 0) & 0x3fff);
				m_colRelative = (ushort)((nBitmask0 >> 14) & 0x1);
				m_rowRelative = (ushort)((nBitmask0 >> 15) & 0x1);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				int nBitmask0 = 0;
				nBitmask0 += m_col << 0;
				nBitmask0 += m_colRelative << 14;
				nBitmask0 += m_rowRelative << 15;
				pBlobView.PackUint16((ushort)(nBitmask0));
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_col = 0;
				m_colRelative = 0;
				m_rowRelative = 0;
			}

			public ushort m_col;
			public ushort m_colRelative;
			public ushort m_rowRelative;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellStruct : BiffStruct
		{
			public CellStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_rw.BlobRead(pBlobView);
				m_col.BlobRead(pBlobView);
				m_ixfe.BlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				m_rw.BlobWrite(pBlobView);
				m_col.BlobWrite(pBlobView);
				m_ixfe.BlobWrite(pBlobView);
			}

			public const ushort SIZE = 6;
			public void SetDefaults()
			{
				m_rw = new RwStruct();
				m_col = new ColStruct();
				m_ixfe = new IXFCellStruct();
			}

			public RwStruct m_rw;
			public ColStruct m_col;
			public IXFCellStruct m_ixfe;
			~CellStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellParsedFormulaStruct : BiffStruct
		{
			public CellParsedFormulaStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_cce = pBlobView.UnpackUint16();
				PostBlobRead(pBlobView);
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				PreBlobWrite(pBlobView);
				pBlobView.PackUint16(m_cce);
				PostBlobWrite(pBlobView);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_cce = 0;
				PostSetDefaults();
			}

			public ushort m_cce;
			public RgceStruct m_rgce;
			public CellParsedFormulaStruct(Formula pFormula, WorkbookGlobals pWorkbookGlobals)
			{
				SetDefaults();
				if (pFormula != null)
				{
					nbAssert.Assert(pWorkbookGlobals != null);
					pFormula.ToRgce(m_rgce, pWorkbookGlobals);
				}
			}

			protected void PostSetDefaults()
			{
				m_rgce = new RgceStruct();
			}

			protected void PostBlobRead(BlobView pBlobView)
			{
				int nStart = pBlobView.GetStart() + pBlobView.GetOffset();
				int nEnd = nStart + m_cce;
				nbAssert.Assert(nEnd <= pBlobView.GetEnd());
				BlobView pTempBlobView = new BlobView(pBlobView.GetBlob(), nStart, nEnd);
				m_rgce.BlobRead(pTempBlobView);
				pBlobView.SetOffset(pBlobView.GetOffset() + pTempBlobView.GetOffset());
			}

			protected void PreBlobWrite(BlobView pBlobView)
			{
				m_cce = (ushort)(m_rgce.GetSize());
			}

			protected void PostBlobWrite(BlobView pBlobView)
			{
				m_rgce.BlobWrite(pBlobView);
			}

			public int GetSize()
			{
				return SIZE + m_rgce.GetSize();
			}

			~CellParsedFormulaStruct()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BuiltInStyleStruct : BiffStruct
		{
			public BuiltInStyleStruct()
			{
				SetDefaults();
			}

			public override void BlobRead(BlobView pBlobView)
			{
				m_istyBuiltIn = pBlobView.UnpackUint8();
				m_iLevel = pBlobView.UnpackUint8();
			}

			public override void BlobWrite(BlobView pBlobView)
			{
				pBlobView.PackUint8(m_istyBuiltIn);
				pBlobView.PackUint8(m_iLevel);
			}

			public const ushort SIZE = 2;
			public void SetDefaults()
			{
				m_istyBuiltIn = 0;
				m_iLevel = 0;
			}

			public byte m_istyBuiltIn;
			public byte m_iLevel;
		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OfficeArtDimensions
		{
			public ushort m_nCellX1;
			public ushort m_nSubCellX1;
			public ushort m_nCellY1;
			public ushort m_nSubCellY1;
			public ushort m_nCellX2;
			public ushort m_nSubCellX2;
			public ushort m_nCellY2;
			public ushort m_nSubCellY2;
		}
		class BiffWorksheet : Worksheet
		{
			protected BiffRecordContainer m_pBiffRecordContainer;
			public BiffWorksheet(Workbook pWorkbook) : base(pWorkbook)
			{
				m_pBiffRecordContainer = null;
			}

			public BiffWorksheet(Workbook pWorkbook, BiffWorkbookGlobals pBiffWorkbookGlobals, BiffRecord pInitialBiffRecord, Stream pStream) : base(pWorkbook)
			{
				m_pBiffRecordContainer = new BiffRecordContainer(pInitialBiffRecord, pStream);
				SetName(pBiffWorkbookGlobals.GetNextWorksheetName());
				for (int i = 0; i < m_pBiffRecordContainer.m_pBiffRecordVector.GetSize(); i++)
				{
					BiffRecord pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
					switch (pBiffRecord.GetType())
					{
						case BiffRecord.Type.TYPE_PrintGrid:
						{
							PrintGridRecord pPrintGridRecord = (PrintGridRecord)(pBiffRecord);
							SetPrintGridlines(pPrintGridRecord.GetPrintGridlines());
							break;
						}

						case BiffRecord.Type.TYPE_DEFAULTROWHEIGHT:
						{
							DefaultRowHeight pDefaultRowHeight = (DefaultRowHeight)(pBiffRecord);
							m_pImpl.m_nDefaultRowHeight = (ushort)(pDefaultRowHeight.GetRowHeight());
							break;
						}

						case BiffRecord.Type.TYPE_SETUP:
						{
							SetupRecord pSetupRecord = (SetupRecord)(pBiffRecord);
							if (pSetupRecord.GetPortrait())
								SetOrientation(Orientation.ORIENTATION_PORTRAIT);
							else
								SetOrientation(Orientation.ORIENTATION_LANDSCAPE);
							break;
						}

						case BiffRecord.Type.TYPE_COLINFO:
						{
							ColInfoRecord pColInfoRecord = (ColInfoRecord)(pBiffRecord);
							for (ushort j = pColInfoRecord.GetFirstColumn(); j <= pColInfoRecord.GetLastColumn(); j++)
							{
								SetColumnWidth(j, (ushort)(((uint)(pColInfoRecord.GetColumnWidth()) * 1789 + 65426 / 2) / 65426));
								SetColumnHidden(j, pColInfoRecord.GetHidden());
							}
							break;
						}

						case BiffRecord.Type.TYPE_ROW:
						{
							RowRecord pRowRecord = (RowRecord)(pBiffRecord);
							SetRowHeight(pRowRecord.GetRow(), (ushort)(((uint)(pRowRecord.GetHeight()) * 546 + 8190 / 2) / 8190));
							break;
						}

						case BiffRecord.Type.TYPE_LABELSST:
						{
							LabelSstRecord pLabelSstRecord = (LabelSstRecord)(pBiffRecord);
							Cell pCell = GetCell(pLabelSstRecord.GetX(), pLabelSstRecord.GetY());
							pCell.SetString(pBiffWorkbookGlobals.GetSharedStringByIndex(pLabelSstRecord.GetSstIndex()));
							pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pLabelSstRecord.GetXfIndex()));
							break;
						}

						case BiffRecord.Type.TYPE_RK:
						{
							RkRecord pRkRecord = (RkRecord)(pBiffRecord);
							Cell pCell = GetCell(pRkRecord.GetX(), pRkRecord.GetY());
							pCell.SetFloat(BiffUtils.RkValueDecode(pRkRecord.GetRkValue()));
							pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pRkRecord.GetXfIndex()));
							break;
						}

						case BiffRecord.Type.TYPE_MULRK:
						{
							MulRkRecord pMulRkRecord = (MulRkRecord)(pBiffRecord);
							ushort nX = pMulRkRecord.GetX();
							ushort nY = pMulRkRecord.GetY();
							ushort nNumRk = pMulRkRecord.GetNumRk();
							for (ushort j = 0; j < nNumRk; j++)
							{
								Cell pCell = GetCell((ushort)(nX + j), nY);
								pCell.SetFloat(BiffUtils.RkValueDecode(pMulRkRecord.GetRkValueByIndex(j)));
								pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pMulRkRecord.GetXfIndexByIndex(j)));
							}
							break;
						}

						case BiffRecord.Type.TYPE_BLANK:
						{
							Blank pBlank = (Blank)(pBiffRecord);
							Cell pCell = GetCell(pBlank.GetX(), pBlank.GetY());
							pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pBlank.GetXfIndex()));
							break;
						}

						case BiffRecord.Type.TYPE_MULBLANK:
						{
							MulBlank pMulBlank = (MulBlank)(pBiffRecord);
							ushort nX = pMulBlank.GetX();
							ushort nY = pMulBlank.GetY();
							ushort nNumColumn = pMulBlank.GetNumColumn();
							for (ushort j = 0; j < nNumColumn; j++)
							{
								Cell pCell = GetCell((ushort)(nX + j), nY);
								pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pMulBlank.GetXfIndexByIndex(j)));
							}
							break;
						}

						case BiffRecord.Type.TYPE_NUMBER:
						{
							NumberRecord pNumberRecord = (NumberRecord)(pBiffRecord);
							Cell pCell = GetCell(pNumberRecord.GetX(), pNumberRecord.GetY());
							pCell.SetFloat(pNumberRecord.GetNumber());
							pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pNumberRecord.GetXfIndex()));
							break;
						}

						case BiffRecord.Type.TYPE_BOOLERR:
						{
							BoolErrRecord pBoolErrRecord = (BoolErrRecord)(pBiffRecord);
							if (pBoolErrRecord.IsBoolean())
							{
								Cell pCell = GetCell(pBoolErrRecord.GetX(), pBoolErrRecord.GetY());
								pCell.SetBoolean(pBoolErrRecord.GetBoolean());
								pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pBoolErrRecord.GetXfIndex()));
							}
							break;
						}

						case BiffRecord.Type.TYPE_FORMULA:
						{
							FormulaRecord pFormulaRecord = (FormulaRecord)(pBiffRecord);
							ushort nX = pFormulaRecord.GetX();
							ushort nY = pFormulaRecord.GetY();
							Cell pCell = GetCell(nX, nY);
							pCell.m_pImpl.SetFormula(pFormulaRecord.GetFormula(pBiffWorkbookGlobals));
							pCell.SetStyle(pBiffWorkbookGlobals.GetStyleByXfIndex(pFormulaRecord.GetXfIndex()));
							break;
						}

						case BiffRecord.Type.TYPE_OBJ:
						{
							ObjRecord pObjRecord = (ObjRecord)(pBiffRecord);
							if (pObjRecord.GetType() == FtCmoStruct.ObjType.OBJ_TYPE_PICTURE || pObjRecord.GetType() == FtCmoStruct.ObjType.OBJ_TYPE_CHART)
							{
								if (m_pBiffRecordContainer.m_pBiffRecordVector.Get(i - 1).GetType() == BiffRecord.Type.TYPE_MSO_DRAWING)
								{
									MsoDrawingRecord pMsoDrawingRecord = (MsoDrawingRecord)(m_pBiffRecordContainer.m_pBiffRecordVector.Get(i - 1));
									MsoDrawingRecord_Position pPosition = pMsoDrawingRecord.GetPosition();
									if (pPosition != null)
									{
										ushort nX = pPosition.m_nCellX1;
										int nColumnWidth = GetColumnWidth(nX);
										int nSubX = (pPosition.m_nSubCellX1 * nColumnWidth + 1024 / 2) / 1024;
										ushort nY = pPosition.m_nCellY1;
										int nRowHeight = GetRowHeight(nY);
										int nSubY = (pPosition.m_nSubCellY1 * nRowHeight + 256 / 2) / 256;
										int nSubX2 = (pPosition.m_nSubCellX2 * GetColumnWidth(pPosition.m_nCellX2) + 1024 / 2) / 1024;
										int nSubY2 = (pPosition.m_nSubCellY2 * GetRowHeight(pPosition.m_nCellY2) + 256 / 2) / 256;
										int nWidth = 0;
										if (pPosition.m_nCellX1 == pPosition.m_nCellX2)
										{
											nWidth = nSubX2 - nSubX;
										}
										else
										{
											nWidth = nColumnWidth - nSubX;
											for (ushort nColumn = (ushort)(pPosition.m_nCellX1 + 1); nColumn < pPosition.m_nCellX2; nColumn++)
												nWidth = nWidth + GetColumnWidth(nColumn);
											nWidth = nWidth + nSubX2;
										}
										int nHeight = 0;
										if (pPosition.m_nCellY1 == pPosition.m_nCellY2)
										{
											nHeight = nSubY2 - nSubY;
										}
										else
										{
											nHeight = nRowHeight - nSubY;
											for (ushort nRow = (ushort)(pPosition.m_nCellY1 + 1); nRow < pPosition.m_nCellY2; nRow++)
												nHeight = nHeight + GetRowHeight(nRow);
											nHeight = nHeight + nSubY2;
										}
										switch (pObjRecord.GetType())
										{
											case FtCmoStruct.ObjType.OBJ_TYPE_PICTURE:
											{
												OfficeArtFOPTEStruct pPib = pMsoDrawingRecord.GetProperty(OfficeArtRecord.OPIDType.OPID_PIB);
												if (pPib != null)
												{
													if (pPib.m_opid.m_fComplex == 0x0)
													{
														uint nBlipIndex = Utils.ByteConvertInt32ToUint32(pPib.m_op) - 1;
														MsoDrawingGroupRecord pMsoDrawingGroupRecord = pBiffWorkbookGlobals.GetMsoDrawingGroupRecord();
														OfficeArtDggContainerRecord pOfficeArtDggContainerRecord = pMsoDrawingGroupRecord.GetOfficeArtDggContainerRecord();
														OfficeArtBStoreContainerRecord pOfficeArtBStoreContainerRecord = (OfficeArtBStoreContainerRecord)(pOfficeArtDggContainerRecord.FindOfficeArtRecordByType(OfficeArtRecord.Type.TYPE_OFFICE_ART_B_STORE_CONTAINER));
														if (nBlipIndex < pOfficeArtBStoreContainerRecord.GetNumOfficeArtRecord())
														{
															OfficeArtRecord pOfficeArtRecord = pOfficeArtBStoreContainerRecord.GetOfficeArtRecordByIndex((ushort)(nBlipIndex));
															if (pOfficeArtRecord != null && pOfficeArtRecord.GetType() == OfficeArtRecord.Type.TYPE_OFFICE_ART_FBSE)
															{
																OfficeArtFBSERecord pOfficeArtFBSERecord = (OfficeArtFBSERecord)(pOfficeArtRecord);
																OfficeArtBlipRecord pEmbeddedBlip = pOfficeArtFBSERecord.GetEmbeddedBlip();
																Picture.Format eFormat = Picture.Format.FORMAT_JPEG;
																switch (pEmbeddedBlip.GetType())
																{
																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_EMF:
																	{
																		eFormat = Picture.Format.FORMAT_EMF;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_WMF:
																	{
																		eFormat = Picture.Format.FORMAT_WMF;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_PICT:
																	{
																		eFormat = Picture.Format.FORMAT_PICT;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_JPEG:
																	{
																		eFormat = Picture.Format.FORMAT_JPEG;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_PNG:
																	{
																		eFormat = Picture.Format.FORMAT_PNG;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_DIB:
																	{
																		eFormat = Picture.Format.FORMAT_DIB;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_TIFF:
																	{
																		eFormat = Picture.Format.FORMAT_TIFF;
																		break;
																	}

																	case OfficeArtRecord.Type.TYPE_OFFICE_ART_BLIP_JPEG_CMYK:
																	{
																		eFormat = Picture.Format.FORMAT_JPEG;
																		break;
																	}

																	default:
																	{
																		nbAssert.Assert(false);
																		break;
																	}

																}
																InternalString sUrl = new InternalString("");
																OfficeArtFOPTEStruct pHyperlink = pMsoDrawingRecord.GetProperty(OfficeArtRecord.OPIDType.OPID_HYPERLINK);
																if (pHyperlink != null && pHyperlink.m_pComplexData != null)
																{
																	BlobView pBlobView = pHyperlink.m_pComplexData.GetBlobView();
																	pBlobView.SetOffset(0);
																	IHlinkStruct pIHlink = new IHlinkStruct();
																	pIHlink.BlobRead(pBlobView);
																	if (pIHlink.m_hyperlink.m_haxUrl != null)
																		sUrl.Set(pIHlink.m_hyperlink.m_haxUrl.GetExternalString());
																	{
																		pIHlink = null;
																	}
																}
																Picture pPicture = new Picture(pEmbeddedBlip.GetBlob(), eFormat);
																pPicture.SetX(nX);
																pPicture.SetSubX((ushort)(nSubX));
																pPicture.SetY(nY);
																pPicture.SetSubY((ushort)(nSubY));
																pPicture.SetWidth((ushort)(nWidth));
																pPicture.SetHeight((ushort)(nHeight));
																pPicture.SetUrl(sUrl.GetExternalString());
																{
																	NumberDuck.Picture __417512960 = pPicture;
																	pPicture = null;
																	m_pImpl.m_pPictureVector.PushBack(__417512960);
																}
																{
																	sUrl = null;
																}
															}
														}
													}
												}
												break;
											}

											case FtCmoStruct.ObjType.OBJ_TYPE_CHART:
											{
												if (m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_BOF)
												{
													BofRecord pBofRecord = (BofRecord)(m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1));
													if (pBofRecord.GetBofType() == BofRecord.BofType.BOF_TYPE_CHART)
													{
														Chart pChart = null;
														LineFormatRecord pDefaultLineFormatRecord = null;
														AreaFormatRecord pDefaultAreaFormatRecord = null;
														MarkerFormatRecord pDefaultMarkerFormatRecord = null;
														int j = i + 2;
														pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(j);
														while (pBiffRecord.GetType() != BiffRecord.Type.TYPE_EOF)
														{
															j++;
															pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(j);
															if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_ChartFormat && m_pBiffRecordContainer.m_pBiffRecordVector.Get(j + 1).GetType() == BiffRecord.Type.TYPE_Begin)
															{
																j++;
																while (true)
																{
																	j++;
																	pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(j);
																	if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Bar)
																		pChart = new Chart(this, ((BarRecord)(pBiffRecord)).GetChartType());
																	else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Line)
																		pChart = new Chart(this, ((LineRecord)(pBiffRecord)).GetChartType());
																	else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Area)
																		pChart = new Chart(this, ((AreaRecord)(pBiffRecord)).GetChartType());
																	else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Scatter)
																		pChart = new Chart(this, ((ScatterRecord)(pBiffRecord)).GetChartType());
																	else if (pChart != null && pBiffRecord.GetType() == BiffRecord.Type.TYPE_DataFormat && m_pBiffRecordContainer.m_pBiffRecordVector.Get(j + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																	{
																		j++;
																		while (true)
																		{
																			j++;
																			pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(j);
																			if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_LineFormat)
																				pDefaultLineFormatRecord = (LineFormatRecord)(pBiffRecord);
																			if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AreaFormat)
																				pDefaultAreaFormatRecord = (AreaFormatRecord)(pBiffRecord);
																			else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_MarkerFormat)
																				pDefaultMarkerFormatRecord = (MarkerFormatRecord)(pBiffRecord);
																			else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																				j = LoopToEnd(j, m_pBiffRecordContainer);
																			else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																				break;
																		}
																	}
																	else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																		j = LoopToEnd(j, m_pBiffRecordContainer);
																	else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																		break;
																}
															}
														}
														if (pChart != null)
														{
															pChart.SetX(nX);
															pChart.SetSubX((ushort)(nSubX));
															pChart.SetY(nY);
															pChart.SetSubY((ushort)(nSubY));
															pChart.SetWidth((ushort)(nWidth));
															pChart.SetHeight((ushort)(nHeight));
															pChart.m_pImpl.SetClassicStyle();
															while (true)
															{
																i++;
																pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Chart && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																{
																	i++;
																	while (true)
																	{
																		i++;
																		pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																		if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Frame && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																		{
																			i++;
																			LineFormatRecord pLineFormat = null;
																			AreaFormatRecord pAreaFormat = null;
																			while (true)
																			{
																				i++;
																				pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																				if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_LineFormat)
																				{
																					pLineFormat = (LineFormatRecord)(pBiffRecord);
																					pLineFormat.ModifyLine(pChart.GetFrameBorderLine(), pBiffWorkbookGlobals);
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AreaFormat)
																				{
																					pAreaFormat = (AreaFormatRecord)(pBiffRecord);
																					pAreaFormat.ModifyFill(pChart.GetFrameFill(), pBiffWorkbookGlobals);
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																					i = LoopToEnd(i, m_pBiffRecordContainer);
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_ShapePropsStream)
																				{
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																					break;
																			}
																		}
																		else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Series && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																		{
																			i++;
																			Formula pNameFormula = null;
																			Formula pValuesFormula = null;
																			Formula pCategoriesFormula = null;
																			int k = i;
																			while (true)
																			{
																				k++;
																				pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(k);
																				if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_BRAI)
																				{
																					BraiRecord pBraiRecord = (BraiRecord)(pBiffRecord);
																					if (pBraiRecord.GetRt() == 0x02)
																					{
																						if (pBraiRecord.GetId() == 0x00)
																							pNameFormula = pBraiRecord.GetFormula(pBiffWorkbookGlobals);
																						else if (pBraiRecord.GetId() == 0x01)
																							pValuesFormula = pBraiRecord.GetFormula(pBiffWorkbookGlobals);
																						else if (pBraiRecord.GetId() == 0x02)
																							pCategoriesFormula = pBraiRecord.GetFormula(pBiffWorkbookGlobals);
																					}
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																					k = LoopToEnd(k, m_pBiffRecordContainer);
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																					break;
																			}
																			if (pValuesFormula != null)
																			{
																				Series pSeries;
																				{
																					NumberDuck.Secret.Formula __310527988 = pValuesFormula;
																					pValuesFormula = null;
																					pSeries = pChart.m_pImpl.CreateSeries(__310527988);
																				}
																				if (pNameFormula != null)
																				{
																					NumberDuck.Secret.Formula __4039636097 = pNameFormula;
																					pNameFormula = null;
																					pSeries.m_pImpl.SetNameFormula(__4039636097);
																				}
																				if (pCategoriesFormula != null)
																				{
																					NumberDuck.Secret.Formula __3225269424 = pCategoriesFormula;
																					pCategoriesFormula = null;
																					pChart.m_pImpl.SetCategoriesFormula(__3225269424);
																				}
																				pSeries.m_pImpl.SetClassicStyle(pChart.GetType(), (ushort)(pChart.GetNumSeries() - 1));
																				if (pDefaultLineFormatRecord != null)
																					pDefaultLineFormatRecord.ModifyLine(pSeries.GetLine(), pBiffWorkbookGlobals);
																				if (pDefaultAreaFormatRecord != null)
																					pDefaultAreaFormatRecord.ModifyFill(pSeries.GetFill(), pBiffWorkbookGlobals);
																				if (pDefaultMarkerFormatRecord != null)
																					pDefaultMarkerFormatRecord.ModifyMarker(pSeries.GetMarker(), pBiffWorkbookGlobals);
																				while (true)
																				{
																					i++;
																					pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_DataFormat && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																					{
																						i++;
																						while (true)
																						{
																							i++;
																							pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																							if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_LineFormat)
																								((LineFormatRecord)(pBiffRecord)).ModifyLine(pSeries.GetLine(), pBiffWorkbookGlobals);
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AreaFormat)
																								((AreaFormatRecord)(pBiffRecord)).ModifyFill(pSeries.GetFill(), pBiffWorkbookGlobals);
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_MarkerFormat)
																								((MarkerFormatRecord)(pBiffRecord)).ModifyMarker(pSeries.GetMarker(), pBiffWorkbookGlobals);
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																								i = LoopToEnd(i, m_pBiffRecordContainer);
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																								break;
																						}
																					}
																					else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																						i = LoopToEnd(i, m_pBiffRecordContainer);
																					else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																						break;
																				}
																			}
																			else
																			{
																				{
																					pNameFormula = null;
																				}
																				{
																					pValuesFormula = null;
																				}
																				{
																					pCategoriesFormula = null;
																				}
																				i = LoopToEnd(i, m_pBiffRecordContainer);
																			}
																		}
																		else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AxisParent && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																		{
																			i++;
																			while (true)
																			{
																				i++;
																				pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																				if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Axis && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																				{
																					i++;
																					ushort nType = ((AxisRecord)(pBiffRecord)).GetType();
																					if (nType == 0x0000 || nType == 0x0001)
																					{
																						Line pAxisLine = pChart.GetHorizontalAxisLine();
																						Line pGridLine = pChart.GetHorizontalGridLine();
																						if (((AxisRecord)(pBiffRecord)).GetType() == 0x0001)
																						{
																							pAxisLine = pChart.GetVerticalAxisLine();
																							pGridLine = pChart.GetVerticalGridLine();
																						}
																						while (true)
																						{
																							i++;
																							pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																							if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AxisLine && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_LineFormat)
																							{
																								switch (((AxisLineRecord)(pBiffRecord)).GetId())
																								{
																									case 0x0000:
																									{
																										i++;
																										pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																										((LineFormatRecord)(pBiffRecord)).ModifyLine(pAxisLine, pBiffWorkbookGlobals);
																										break;
																									}

																									case 0x0001:
																									{
																										i++;
																										pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																										((LineFormatRecord)(pBiffRecord)).ModifyLine(pGridLine, pBiffWorkbookGlobals);
																										break;
																									}

																								}
																							}
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																								i = LoopToEnd(i, m_pBiffRecordContainer);
																							else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																								break;
																						}
																					}
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_PlotArea && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Frame && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 2).GetType() == BiffRecord.Type.TYPE_Begin)
																				{
																					i += 2;
																					while (true)
																					{
																						i++;
																						pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																						if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_LineFormat)
																							((LineFormatRecord)(pBiffRecord)).ModifyLine(pChart.GetPlotBorderLine(), pBiffWorkbookGlobals);
																						else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AreaFormat)
																							((AreaFormatRecord)(pBiffRecord)).ModifyFill(pChart.GetPlotFill(), pBiffWorkbookGlobals);
																						else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																							i = LoopToEnd(i, m_pBiffRecordContainer);
																						else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																							break;
																					}
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_ChartFormat && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																				{
																					i++;
																					while (true)
																					{
																						i++;
																						pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																						if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Legend && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																						{
																							((LegendRecord)(pBiffRecord)).ModifyLegend(pChart.GetLegend(), pBiffWorkbookGlobals);
																							i++;
																							while (true)
																							{
																								i++;
																								pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																								if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Frame && m_pBiffRecordContainer.m_pBiffRecordVector.Get(i + 1).GetType() == BiffRecord.Type.TYPE_Begin)
																								{
																									i++;
																									while (true)
																									{
																										i++;
																										pBiffRecord = m_pBiffRecordContainer.m_pBiffRecordVector.Get(i);
																										if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_LineFormat)
																											((LineFormatRecord)(pBiffRecord)).ModifyLine(pChart.GetLegend().GetBorderLine(), pBiffWorkbookGlobals);
																										else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_AreaFormat)
																											((AreaFormatRecord)(pBiffRecord)).ModifyFill(pChart.GetLegend().GetFill(), pBiffWorkbookGlobals);
																										else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																											i = LoopToEnd(i, m_pBiffRecordContainer);
																										else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																											break;
																									}
																								}
																								else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																									i = LoopToEnd(i, m_pBiffRecordContainer);
																								else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																									break;
																							}
																						}
																						else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																							i = LoopToEnd(i, m_pBiffRecordContainer);
																						else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																							break;
																					}
																				}
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																					i = LoopToEnd(i, m_pBiffRecordContainer);
																				else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																					break;
																			}
																		}
																		else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
																			i = LoopToEnd(i, m_pBiffRecordContainer);
																		else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
																			break;
																	}
																}
																else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_EOF)
																	break;
															}
															{
																NumberDuck.Chart __477487748 = pChart;
																pChart = null;
																m_pImpl.m_pChartVector.PushBack(__477487748);
															}
														}
													}
												}
												break;
											}

										}
									}
								}
							}
							break;
						}

						case BiffRecord.Type.TYPE_WINDOW2:
						{
							Window2Record pWindow2Record = (Window2Record)(pBiffRecord);
							SetShowGridlines(pWindow2Record.GetShowGridlines());
							break;
						}

						case BiffRecord.Type.TYPE_MergeCells:
						{
							MergeCellsRecord pMergeCellsRecord = (MergeCellsRecord)(pBiffRecord);
							for (ushort j = 0; j < pMergeCellsRecord.GetNumMergedCell(); j++)
							{
								Ref8Struct pRef8 = pMergeCellsRecord.GetMergedCell(j);
								MergedCell pMergedCell = CreateMergedCell(pRef8.m_colFirst, pRef8.m_rwFirst, (ushort)(pRef8.m_colLast - pRef8.m_colFirst + 1), (ushort)(pRef8.m_rwLast - pRef8.m_rwFirst + 1));
							}
							break;
						}

					}
				}
			}

			public static void Write(Worksheet pWorksheet, WorkbookGlobals pWorkbookGlobals, ushort nWorksheetIndex, BiffRecordContainer pBiffRecordContainer)
			{
				pBiffRecordContainer.AddBiffRecord(new BofRecord(BofRecord.BofType.BOF_TYPE_SHEET));
				pBiffRecordContainer.AddBiffRecord(new PrintRowColRecord());
				pBiffRecordContainer.AddBiffRecord(new PrintGridRecord(pWorksheet.m_pImpl.m_bPrintGridlines));
				pBiffRecordContainer.AddBiffRecord(new DefaultRowHeight((short)(pWorksheet.m_pImpl.m_nDefaultRowHeight)));
				pBiffRecordContainer.AddBiffRecord(new SetupRecord(pWorksheet.m_pImpl.m_eOrientation == Orientation.ORIENTATION_PORTRAIT));
				pBiffRecordContainer.AddBiffRecord(new HeaderFooterRecord());
				int nLastUsedRow = 0;
				{
					int nSize = pWorksheet.m_pImpl.m_pRowInfoTable.GetSize();
					if (nSize > 0)
					{
						TableElement<RowInfo> pLastElement = pWorksheet.m_pImpl.m_pRowInfoTable.GetByIndex(nSize - 1);
						nLastUsedRow = pLastElement.m_nRow;
					}
				}
				{
					int nSize = pWorksheet.m_pImpl.m_pCellTable.GetSize();
					if (nSize > 0)
					{
						TableElement<Cell> pElement = pWorksheet.m_pImpl.m_pCellTable.GetByIndex(nSize - 1);
						if (pElement.m_nRow > nLastUsedRow)
							nLastUsedRow = pElement.m_nRow;
					}
				}
				DefColWidthRecord pDefColWidth = new DefColWidthRecord(8);
				{
					NumberDuck.Secret.DefColWidthRecord __2009045423 = pDefColWidth;
					pDefColWidth = null;
					pBiffRecordContainer.AddBiffRecord(__2009045423);
				}
				{
					int i = 0;
					while (true)
					{
						if (i >= pWorksheet.m_pImpl.m_pColumnInfoTable.GetSize())
							break;
						TableElement<ColumnInfo> pCurrent = pWorksheet.m_pImpl.m_pColumnInfoTable.GetByIndex(i);
						TableElement<ColumnInfo> pLast = pWorksheet.m_pImpl.m_pColumnInfoTable.GetByIndex(i);
						int nNextIndex = i;
						while (true)
						{
							if (i == pWorksheet.m_pImpl.m_pColumnInfoTable.GetSize() - 1)
								break;
							TableElement<ColumnInfo> pNext = pWorksheet.m_pImpl.m_pColumnInfoTable.GetByIndex(i + 1);
							if (pNext.m_xObject.m_nWidth != pCurrent.m_xObject.m_nWidth)
								break;
							if (pNext.m_xObject.m_bHidden != pCurrent.m_xObject.m_bHidden)
								break;
							pLast = pNext;
							i++;
						}
						ColInfoRecord pColInfo = new ColInfoRecord((ushort)(pCurrent.m_nColumn), (ushort)(pLast.m_nColumn), (ushort)((uint)(pCurrent.m_xObject.m_nWidth) * 65426 / 1789), pCurrent.m_xObject.m_bHidden);
						{
							NumberDuck.Secret.ColInfoRecord __4169651391 = pColInfo;
							pColInfo = null;
							pBiffRecordContainer.AddBiffRecord(__4169651391);
						}
						i++;
					}
				}
				pBiffRecordContainer.AddBiffRecord(new DimensionRecord(0, (uint)(nLastUsedRow + 1), 0, 0, 0));
				int nCurrentRowBlock = 0;
				RowRecord pLastRow = null;
				RowRecord pCurrentRow = null;
				OwnedVector<RowRecord> pRowRecordVector = new OwnedVector<RowRecord>();
				OwnedVector<BiffRecord> pCellRecordVector = new OwnedVector<BiffRecord>();
				int nCellIndex = 0;
				int nRowIndex = 0;
				while (true)
				{
					if (nCellIndex < pWorksheet.m_pImpl.m_pCellTable.GetSize())
					{
						TableElement<Cell> pElement = pWorksheet.m_pImpl.m_pCellTable.GetByIndex(nCellIndex);
						if (pElement.m_nRow / 32 == (int)(nCurrentRowBlock))
						{
							Cell pCell = pElement.m_xObject;
							if (pCurrentRow == null || pCurrentRow.GetRow() < pElement.m_nRow)
							{
								while (true)
								{
									if (nRowIndex == pWorksheet.m_pImpl.m_pRowInfoTable.GetSize())
										break;
									TableElement<RowInfo> pRowElement = pWorksheet.m_pImpl.m_pRowInfoTable.GetByIndex(nRowIndex);
									if (pRowElement.m_nRow > pElement.m_nRow)
										break;
									if (pRowElement.m_xObject.m_nHeight > 0)
									{
										pLastRow = pCurrentRow;
										RowRecord pOwnedRowRecord = new RowRecord((ushort)(pRowElement.m_nRow), (ushort)((uint)(pRowElement.m_xObject.m_nHeight) * 8190 / 546));
										pCurrentRow = pOwnedRowRecord;
										{
											NumberDuck.Secret.RowRecord __3618893028 = pOwnedRowRecord;
											pOwnedRowRecord = null;
											pRowRecordVector.PushBack(__3618893028);
										}
										if (pLastRow != null && pCurrentRow.GetRow() == pLastRow.GetRow() + 1 && pLastRow.GetBottomThick())
											pCurrentRow.SetTopThick();
									}
									nRowIndex++;
								}
								if (pCurrentRow == null || pCurrentRow.GetRow() < pElement.m_nRow)
								{
									pLastRow = pCurrentRow;
									RowRecord pOwnedRowRecord = new RowRecord((ushort)(pElement.m_nRow), (ushort)(DEFAULT_ROW_HEIGHT * 8190 / 546));
									pCurrentRow = pOwnedRowRecord;
									{
										NumberDuck.Secret.RowRecord __3618893028 = pOwnedRowRecord;
										pOwnedRowRecord = null;
										pRowRecordVector.PushBack(__3618893028);
									}
									if (pLastRow != null && pCurrentRow.GetRow() == pLastRow.GetRow() + 1 && pLastRow.GetBottomThick())
										pCurrentRow.SetTopThick();
								}
							}
							Style pStyle = pCell.GetStyle();
							Line.Type eType = pStyle.GetTopBorderLine().GetType();
							switch (eType)
							{
								case Line.Type.TYPE_THICK:
								case Line.Type.TYPE_DOUBLE:
								{
									if (pLastRow != null && pCurrentRow.GetRow() == pLastRow.GetRow() + 1)
										pLastRow.SetBottomThick();
									pCurrentRow.SetTopThick();
									break;
								}

								case Line.Type.TYPE_MEDIUM:
								case Line.Type.TYPE_MEDIUM_DASHED:
								case Line.Type.TYPE_MEDIUM_DASH_DOT:
								case Line.Type.TYPE_MEDIUM_DASH_DOT_DOT:
								case Line.Type.TYPE_SLANT_DASH_DOT_DOT:
								{
									if (pLastRow != null && pCurrentRow.GetRow() == pLastRow.GetRow() + 1)
										pLastRow.SetBottomMedium();
									pCurrentRow.SetTopMedium();
									break;
								}

							}
							eType = pStyle.GetBottomBorderLine().GetType();
							switch (eType)
							{
								case Line.Type.TYPE_THICK:
								case Line.Type.TYPE_DOUBLE:
								{
									pCurrentRow.SetBottomThick();
									break;
								}

								case Line.Type.TYPE_MEDIUM:
								case Line.Type.TYPE_MEDIUM_DASHED:
								case Line.Type.TYPE_MEDIUM_DASH_DOT:
								case Line.Type.TYPE_MEDIUM_DASH_DOT_DOT:
								case Line.Type.TYPE_SLANT_DASH_DOT_DOT:
								{
									pCurrentRow.SetBottomMedium();
									break;
								}

							}
							switch (pCell.GetType())
							{
								case Value.Type.TYPE_EMPTY:
								{
									Vector<int> pXFIndexVector = new Vector<int>();
									TableElement<Cell> pFirstElement = pElement;
									int nNextCellIndex = nCellIndex;
									nNextCellIndex++;
									pXFIndexVector.PushBack(pWorkbookGlobals.GetStyleIndex(pCell.GetStyle()));
									while (nNextCellIndex < pWorksheet.m_pImpl.m_pCellTable.GetSize())
									{
										TableElement<Cell> pNextElement = pWorksheet.m_pImpl.m_pCellTable.GetByIndex(nNextCellIndex);
										if (pNextElement.m_nRow != pElement.m_nRow || pNextElement.m_nColumn != pElement.m_nColumn + 1 || pNextElement.m_xObject.GetType() != Value.Type.TYPE_EMPTY)
											break;
										pXFIndexVector.PushBack(pWorkbookGlobals.GetStyleIndex(pNextElement.m_xObject.GetStyle()));
										nCellIndex++;
										pElement = pWorksheet.m_pImpl.m_pCellTable.GetByIndex(nCellIndex);
										pCell = pElement.m_xObject;
										nNextCellIndex++;
									}
									if (pXFIndexVector.GetSize() > 1)
										pCellRecordVector.PushBack(new MulBlank((ushort)(pFirstElement.m_nColumn), (ushort)(pFirstElement.m_nRow), pXFIndexVector));
									else
										pCellRecordVector.PushBack(new Blank((ushort)(pFirstElement.m_nColumn), (ushort)(pFirstElement.m_nRow), (ushort)(pXFIndexVector.Get(0))));
									{
										pXFIndexVector = null;
									}
									{
										break;
									}
								}

								case Value.Type.TYPE_STRING:
								{
									pCellRecordVector.PushBack(new LabelSstRecord((ushort)(pElement.m_nColumn), (ushort)(pElement.m_nRow), pWorkbookGlobals.GetStyleIndex(pStyle), pWorkbookGlobals.GetSharedStringIndex(pCell.GetString())));
									break;
								}

								case Value.Type.TYPE_FLOAT:
								{
									pCellRecordVector.PushBack(new NumberRecord((ushort)(pElement.m_nColumn), (ushort)(pElement.m_nRow), pWorkbookGlobals.GetStyleIndex(pStyle), pCell.GetFloat()));
									break;
								}

								case Value.Type.TYPE_BOOLEAN:
								{
									pCellRecordVector.PushBack(new BoolErrRecord((ushort)(pElement.m_nColumn), (ushort)(pElement.m_nRow), pWorkbookGlobals.GetStyleIndex(pStyle), pCell.GetBoolean()));
									break;
								}

								case Value.Type.TYPE_FORMULA:
								{
									Formula pFormula = new Formula(pCell.GetFormula(), pWorksheet.m_pImpl);
									pCellRecordVector.PushBack(new FormulaRecord((ushort)(pElement.m_nColumn), (ushort)(pElement.m_nRow), pWorkbookGlobals.GetStyleIndex(pStyle), pFormula, pWorkbookGlobals));
									{
										pFormula = null;
									}
									{
										break;
									}
								}

							}
							nCellIndex++;
						}
					}
					if (nCellIndex >= pWorksheet.m_pImpl.m_pCellTable.GetSize() || pWorksheet.m_pImpl.m_pCellTable.GetByIndex(nCellIndex).m_nRow / 32 != (int)(nCurrentRowBlock))
					{
						while (true)
						{
							if (nRowIndex == pWorksheet.m_pImpl.m_pRowInfoTable.GetSize())
								break;
							TableElement<RowInfo> pRowElement = pWorksheet.m_pImpl.m_pRowInfoTable.GetByIndex(nRowIndex);
							if (pRowElement.m_nRow / 32 != nCurrentRowBlock)
								break;
							if (pRowElement.m_xObject.m_nHeight > 0)
							{
								pLastRow = pCurrentRow;
								RowRecord pOwnedRowRecord = new RowRecord((ushort)(pRowElement.m_nRow), (ushort)((uint)(pRowElement.m_xObject.m_nHeight) * 8190 / 546));
								pCurrentRow = pOwnedRowRecord;
								{
									NumberDuck.Secret.RowRecord __3618893028 = pOwnedRowRecord;
									pOwnedRowRecord = null;
									pRowRecordVector.PushBack(__3618893028);
								}
								if (pLastRow != null && pCurrentRow.GetRow() == pLastRow.GetRow() + 1 && pLastRow.GetBottomThick())
									pCurrentRow.SetTopThick();
							}
							nRowIndex++;
						}
						if (nRowIndex == pWorksheet.m_pImpl.m_pRowInfoTable.GetSize() || pWorksheet.m_pImpl.m_pRowInfoTable.GetByIndex(nRowIndex).m_nRow / 32 != nCurrentRowBlock)
						{
							while (pRowRecordVector.GetSize() > 0)
							{
								pBiffRecordContainer.AddBiffRecord(pRowRecordVector.PopFront());
							}
							while (pCellRecordVector.GetSize() > 0)
							{
								pBiffRecordContainer.AddBiffRecord(pCellRecordVector.PopFront());
							}
							nCurrentRowBlock++;
							if (nCellIndex >= pWorksheet.m_pImpl.m_pCellTable.GetSize() && nRowIndex == pWorksheet.m_pImpl.m_pRowInfoTable.GetSize())
								break;
						}
					}
				}
				if (pWorksheet.GetNumPicture() > 0 || pWorksheet.GetNumChart() > 0)
				{
					ushort nIndex = 0;
					uint nFirstOffset = 0;
					Vector<OfficeArtSpContainerRecord> pOfficeArtSpContainerRecordVector = new Vector<OfficeArtSpContainerRecord>();
					OfficeArtDgContainerRecord pOfficeArtDgContainerRecord = new OfficeArtDgContainerRecord();
					pOfficeArtDgContainerRecord.AddOfficeArtRecord(new OfficeArtFDGRecord(1, 2, 1026));
					OfficeArtSpgrContainerRecord pOfficeArtSpgrContainerRecord = new OfficeArtSpgrContainerRecord();
					OfficeArtSpContainerRecord pOfficeArtSpContainerRecord = new OfficeArtSpContainerRecord();
					pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtFSPGRRecord());
					pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtFSPRecord(0, 1024, 1, 1, 0, 0));
					{
						NumberDuck.Secret.OfficeArtSpContainerRecord __1049470179 = pOfficeArtSpContainerRecord;
						pOfficeArtSpContainerRecord = null;
						pOfficeArtSpgrContainerRecord.AddOfficeArtRecord(__1049470179);
					}
					nIndex = 0;
					for (ushort i = 0; i < pWorksheet.GetNumPicture(); i++)
					{
						Picture pPicture = pWorksheet.GetPictureByIndex(i);
						uint nImageDataIndex = pWorkbookGlobals.PushPicture(pPicture);
						string szUrl = pPicture.GetUrl();
						pOfficeArtSpContainerRecord = new OfficeArtSpContainerRecord();
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtFSPRecord(75, (uint)(1025 + nIndex), 0, 0, 1, 1));
						OfficeArtFOPTRecord pOfficeArtFOPTRecord = new OfficeArtFOPTRecord();
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_PROTECTION_BOOLEAN_PROPERTIES), 0, 33226880);
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_PIB), 1, (int)(nImageDataIndex));
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_BLIP_BOOLEAN_PROPERTIES), 0, 393216);
						if (!ExternalString.Equal(szUrl, ""))
						{
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_STYLE_BOOLEAN_PROPERTIES), 0, 0x00110001);
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_LINE_STYLE_BOOLEAN_PROPERTIES), 0, 0x00180010);
						}
						else
						{
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_STYLE_BOOLEAN_PROPERTIES), 0, 1048576);
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_LINE_STYLE_BOOLEAN_PROPERTIES), 0, 1572880);
						}
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_SHAPE_BOOLEAN_PROPERTIES), 0, 1572880);
						pOfficeArtFOPTRecord.AddStringProperty((ushort)(OfficeArtRecord.OPIDType.OPID_WZ_NAME), "Picture 1");
						if (!ExternalString.Equal(szUrl, ""))
						{
							Blob pBlob = new Blob(true);
							BlobView pBlobView = pBlob.GetBlobView();
							IHlinkStruct pIHlink = new IHlinkStruct();
							pIHlink.m_hyperlink.m_streamVersion = 2;
							pIHlink.m_hyperlink.m_hlstmfHasMoniker = 1;
							pIHlink.m_hyperlink.m_hlstmfIsAbsolute = 1;
							pIHlink.m_hyperlink.m_haxUrl = new InternalString(szUrl);
							pIHlink.BlobWrite(pBlobView);
							pOfficeArtFOPTRecord.AddBlobProperty((ushort)(OfficeArtRecord.OPIDType.OPID_HYPERLINK), 1, pBlob);
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_GROUP_SHAPE_BOOLEAN_PROPERTIES), 0, 0x000A0008);
							{
								pIHlink = null;
							}
							{
								pBlob = null;
							}
						}
						else
						{
							pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_GROUP_SHAPE_BOOLEAN_PROPERTIES), 0, 131072);
						}
						{
							NumberDuck.Secret.OfficeArtFOPTRecord __1214438724 = pOfficeArtFOPTRecord;
							pOfficeArtFOPTRecord = null;
							pOfficeArtSpContainerRecord.AddOfficeArtRecord(__1214438724);
						}
						OfficeArtDimensions pDimensions = ComputeDimensions(pWorksheet, (ushort)(pPicture.GetX()), (ushort)(pPicture.GetSubX()), (ushort)(pPicture.GetWidth()), (ushort)(pPicture.GetY()), (ushort)(pPicture.GetSubY()), (ushort)(pPicture.GetHeight()));
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtClientAnchorSheetRecord(pDimensions.m_nCellX1, pDimensions.m_nSubCellX1, pDimensions.m_nCellY1, pDimensions.m_nSubCellY1, pDimensions.m_nCellX2, pDimensions.m_nSubCellX2, pDimensions.m_nCellY2, pDimensions.m_nSubCellY2));
						{
							pDimensions = null;
						}
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtClientDataRecord());
						pOfficeArtSpContainerRecordVector.PushBack(pOfficeArtSpContainerRecord);
						{
							NumberDuck.Secret.OfficeArtSpContainerRecord __1049470179 = pOfficeArtSpContainerRecord;
							pOfficeArtSpContainerRecord = null;
							pOfficeArtSpgrContainerRecord.AddOfficeArtRecord(__1049470179);
						}
						if (nIndex == 0)
							nFirstOffset = pOfficeArtDgContainerRecord.GetRecursiveSize() + pOfficeArtSpgrContainerRecord.GetRecursiveSize();
						nIndex++;
					}
					for (ushort i = 0; i < pWorksheet.GetNumChart(); i++)
					{
						Chart pChart = pWorksheet.GetChartByIndex(i);
						pOfficeArtSpContainerRecord = new OfficeArtSpContainerRecord();
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtFSPRecord(201, (uint)(1026 + nIndex), 0, 0, 1, 1));
						OfficeArtFOPTRecord pOfficeArtFOPTRecord = new OfficeArtFOPTRecord();
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_PROTECTION_BOOLEAN_PROPERTIES), 0, 31785220);
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_TEXT_BOOLEAN_PROPERTIES), 0, 524296);
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_COLOR), 0, 134217806);
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_FILL_STYLE_BOOLEAN_PROPERTIES), 0, 1048592);
						pOfficeArtFOPTRecord.AddStringProperty((ushort)(OfficeArtRecord.OPIDType.OPID_WZ_NAME), "Chart 1");
						pOfficeArtFOPTRecord.AddProperty((ushort)(OfficeArtRecord.OPIDType.OPID_GROUP_SHAPE_BOOLEAN_PROPERTIES), 0, 131072);
						{
							NumberDuck.Secret.OfficeArtFOPTRecord __1214438724 = pOfficeArtFOPTRecord;
							pOfficeArtFOPTRecord = null;
							pOfficeArtSpContainerRecord.AddOfficeArtRecord(__1214438724);
						}
						OfficeArtDimensions pDimensions = ComputeDimensions(pWorksheet, (ushort)(pChart.GetX()), (ushort)(pChart.GetSubX()), (ushort)(pChart.GetWidth()), (ushort)(pChart.GetY()), (ushort)(pChart.GetSubY()), (ushort)(pChart.GetHeight()));
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtClientAnchorSheetRecord(pDimensions.m_nCellX1, pDimensions.m_nSubCellX1, pDimensions.m_nCellY1, pDimensions.m_nSubCellY1, pDimensions.m_nCellX2, pDimensions.m_nSubCellX2, pDimensions.m_nCellY2, pDimensions.m_nSubCellY2));
						{
							pDimensions = null;
						}
						pOfficeArtSpContainerRecord.AddOfficeArtRecord(new OfficeArtClientDataRecord());
						pOfficeArtSpContainerRecordVector.PushBack(pOfficeArtSpContainerRecord);
						{
							NumberDuck.Secret.OfficeArtSpContainerRecord __1049470179 = pOfficeArtSpContainerRecord;
							pOfficeArtSpContainerRecord = null;
							pOfficeArtSpgrContainerRecord.AddOfficeArtRecord(__1049470179);
						}
						if (nIndex == 0)
							nFirstOffset = pOfficeArtDgContainerRecord.GetRecursiveSize() + pOfficeArtSpgrContainerRecord.GetRecursiveSize();
						nIndex++;
					}
					{
						NumberDuck.Secret.OfficeArtSpgrContainerRecord __2757789752 = pOfficeArtSpgrContainerRecord;
						pOfficeArtSpgrContainerRecord = null;
						pOfficeArtDgContainerRecord.AddOfficeArtRecord(__2757789752);
					}
					nIndex = 0;
					for (ushort i = 0; i < pWorksheet.GetNumPicture(); i++)
					{
						if (nIndex == 0)
							pBiffRecordContainer.AddBiffRecord(new MsoDrawingRecord(pOfficeArtDgContainerRecord, nFirstOffset));
						else
							pBiffRecordContainer.AddBiffRecord(new MsoDrawingRecord(pOfficeArtSpContainerRecordVector.Get(nIndex)));
						pBiffRecordContainer.AddBiffRecord(new ObjRecord(nIndex, FtCmoStruct.ObjType.OBJ_TYPE_PICTURE));
						nIndex++;
					}
					for (ushort i = 0; i < pWorksheet.GetNumChart(); i++)
					{
						Chart pChart = pWorksheet.GetChartByIndex(i);
						if (nIndex == 0)
							pBiffRecordContainer.AddBiffRecord(new MsoDrawingRecord(pOfficeArtDgContainerRecord, nFirstOffset));
						else
							pBiffRecordContainer.AddBiffRecord(new MsoDrawingRecord(pOfficeArtSpContainerRecordVector.Get(nIndex)));
						pBiffRecordContainer.AddBiffRecord(new ObjRecord(nIndex, FtCmoStruct.ObjType.OBJ_TYPE_CHART));
						pBiffRecordContainer.AddBiffRecord(new BofRecord(BofRecord.BofType.BOF_TYPE_CHART));
						pBiffRecordContainer.AddBiffRecord(new ChartFrtInfoRecord());
						pBiffRecordContainer.AddBiffRecord(new UnitsRecord());
						pBiffRecordContainer.AddBiffRecord(new ChartRecord());
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						pBiffRecordContainer.AddBiffRecord(new SclRecord());
						pBiffRecordContainer.AddBiffRecord(new PlotGrowthRecord());
						pBiffRecordContainer.AddBiffRecord(new FrameRecord(false));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						{
							pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetFrameBorderLine(), false));
							pBiffRecordContainer.AddBiffRecord(new AreaFormatRecord(pChart.GetFrameFill()));
						}
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						for (ushort j = 0; j < pChart.GetNumSeries(); j++)
						{
							Series pSeries = pChart.GetSeriesByIndex(j);
							pBiffRecordContainer.AddBiffRecord(new SeriesRecord());
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							{
								Formula pFormula = new Formula(pSeries.GetName(), pWorksheet.m_pImpl);
								if (pFormula.GetNumToken() > 0)
								{
									pBiffRecordContainer.AddBiffRecord(new BraiRecord(0x00, 0x02, false, 0, pFormula, pWorkbookGlobals));
								}
								else
								{
									pBiffRecordContainer.AddBiffRecord(new BraiRecord(0x00, 0x01, false, 0, null, null));
								}
								{
									pFormula = null;
								}
							}
							{
								Formula pFormula = new Formula(pSeries.GetValues(), pWorksheet.m_pImpl);
								pBiffRecordContainer.AddBiffRecord(new BraiRecord(0x01, 0x02, false, 0, pFormula, pWorkbookGlobals));
								{
									pFormula = null;
								}
							}
							{
								Formula pFormula = new Formula(pChart.GetCategories(), pWorksheet.m_pImpl);
								if (pFormula.GetNumToken() > 0)
								{
									pBiffRecordContainer.AddBiffRecord(new BraiRecord(0x02, 0x02, false, 0, pFormula, pWorkbookGlobals));
								}
								else
								{
									pBiffRecordContainer.AddBiffRecord(new BraiRecord(0x02, 0x00, false, 0, null, null));
								}
								{
									pFormula = null;
								}
							}
							pBiffRecordContainer.AddBiffRecord(new BraiRecord(3, 1, false, 0, null, null));
							pBiffRecordContainer.AddBiffRecord(new DataFormatRecord(j));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new Chart3DBarShapeRecord());
							pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pSeries.GetLine(), false));
							pBiffRecordContainer.AddBiffRecord(new AreaFormatRecord(pSeries.GetFill()));
							pBiffRecordContainer.AddBiffRecord(new PieFormatRecord());
							pBiffRecordContainer.AddBiffRecord(new MarkerFormatRecord(pSeries.GetMarker()));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
							pBiffRecordContainer.AddBiffRecord(new SerToCrtRecord(0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						}
						pBiffRecordContainer.AddBiffRecord(new ShtPropsRecord());
						pBiffRecordContainer.AddBiffRecord(new AxesUsedRecord());
						pBiffRecordContainer.AddBiffRecord(new AxisParentRecord());
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						pBiffRecordContainer.AddBiffRecord(new PosRecord(23, -34, 0xffff, 3877, 4038));
						pBiffRecordContainer.AddBiffRecord(new AxisRecord(0x0000));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						pBiffRecordContainer.AddBiffRecord(new CatSerRangeRecord());
						pBiffRecordContainer.AddBiffRecord(new AxcExtRecord());
						pBiffRecordContainer.AddBiffRecord(new AxisLineRecord(0x0000));
						pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetHorizontalAxisLine(), true));
						pBiffRecordContainer.AddBiffRecord(new AxisLineRecord(0x0001));
						pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetHorizontalGridLine(), false));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						pBiffRecordContainer.AddBiffRecord(new AxisRecord(0x0001));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						pBiffRecordContainer.AddBiffRecord(new ValueRangeRecord());
						pBiffRecordContainer.AddBiffRecord(new TickRecord(true));
						pBiffRecordContainer.AddBiffRecord(new FontXRecord());
						pBiffRecordContainer.AddBiffRecord(new AxisLineRecord(0x0000));
						pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetVerticalAxisLine(), true));
						pBiffRecordContainer.AddBiffRecord(new AxisLineRecord(0x0001));
						pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetVerticalGridLine(), false));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						if (!ExternalString.Equal(pChart.GetHorizontalAxisLabel(), ""))
						{
							pBiffRecordContainer.AddBiffRecord(new TextRecord(0, 0, 0x00212121, true, false, false, false, 63, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new PosRecord(0, 0, 0x0000, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new FontXRecord());
							pBiffRecordContainer.AddBiffRecord(new BraiRecord(0, 1, false, 0, null, null));
							pBiffRecordContainer.AddBiffRecord(new SeriesTextRecord(pChart.GetHorizontalAxisLabel()));
							pBiffRecordContainer.AddBiffRecord(new ObjectLinkRecord(0x0003, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						}
						if (!ExternalString.Equal(pChart.GetVerticalAxisLabel(), ""))
						{
							pBiffRecordContainer.AddBiffRecord(new TextRecord(0, 0, 0x00212121, true, false, false, false, 63, 90));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new PosRecord(0, 0, 0x0000, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new FontXRecord());
							pBiffRecordContainer.AddBiffRecord(new BraiRecord(0, 1, false, 0, null, null));
							pBiffRecordContainer.AddBiffRecord(new SeriesTextRecord(pChart.GetVerticalAxisLabel()));
							pBiffRecordContainer.AddBiffRecord(new ObjectLinkRecord(0x0002, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						}
						pBiffRecordContainer.AddBiffRecord(new PlotAreaRecord());
						pBiffRecordContainer.AddBiffRecord(new FrameRecord(true));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetPlotBorderLine(), false));
						pBiffRecordContainer.AddBiffRecord(new AreaFormatRecord(pChart.GetPlotFill()));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						pBiffRecordContainer.AddBiffRecord(new ChartFormatRecord());
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
						switch (pChart.GetType())
						{
							case Chart.Type.TYPE_COLUMN:
							case Chart.Type.TYPE_COLUMN_STACKED:
							case Chart.Type.TYPE_COLUMN_STACKED_100:
							case Chart.Type.TYPE_BAR:
							case Chart.Type.TYPE_BAR_STACKED:
							case Chart.Type.TYPE_BAR_STACKED_100:
							{
								pBiffRecordContainer.AddBiffRecord(new BarRecord(pChart.GetType()));
								break;
							}

							case Chart.Type.TYPE_AREA:
							case Chart.Type.TYPE_AREA_STACKED:
							case Chart.Type.TYPE_AREA_STACKED_100:
							{
								pBiffRecordContainer.AddBiffRecord(new AreaRecord(pChart.GetType()));
								break;
							}

							case Chart.Type.TYPE_LINE:
							case Chart.Type.TYPE_LINE_STACKED:
							case Chart.Type.TYPE_LINE_STACKED_100:
							{
								pBiffRecordContainer.AddBiffRecord(new LineRecord(pChart.GetType()));
								break;
							}

							case Chart.Type.TYPE_SCATTER:
							{
								pBiffRecordContainer.AddBiffRecord(new ScatterRecord(pChart.GetType()));
								break;
							}

							default:
							{
								nbAssert.Assert(false);
								break;
							}

						}
						pBiffRecordContainer.AddBiffRecord(new CrtLinkRecord());
						if (!pChart.GetLegend().GetHidden())
						{
							pBiffRecordContainer.AddBiffRecord(new LegendRecord(pChart.GetLegend()));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new FrameRecord(true));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new LineFormatRecord(pChart.GetLegend().GetBorderLine(), false));
							pBiffRecordContainer.AddBiffRecord(new AreaFormatRecord(pChart.GetLegend().GetFill()));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						}
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						if (!ExternalString.Equal(pChart.GetTitle(), ""))
						{
							pBiffRecordContainer.AddBiffRecord(new TextRecord(0, 0, 0x00212121, true, false, false, true, 63, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_Begin, 0));
							pBiffRecordContainer.AddBiffRecord(new PosRecord(0, 0, 0x0000, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new FontXRecord());
							pBiffRecordContainer.AddBiffRecord(new BraiRecord(0, 1, false, 0, null, null));
							pBiffRecordContainer.AddBiffRecord(new SeriesTextRecord(pChart.GetTitle()));
							pBiffRecordContainer.AddBiffRecord(new ObjectLinkRecord(0x0001, 0, 0));
							pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						}
						pBiffRecordContainer.AddBiffRecord(new CrtLayout12ARecord());
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_End, 0));
						pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_EOF, 0));
						nIndex++;
					}
					{
						pOfficeArtSpContainerRecordVector = null;
					}
					{
						pOfficeArtDgContainerRecord = null;
					}
				}
				if (nWorksheetIndex == 0)
					pBiffRecordContainer.AddBiffRecord(new Window2Record(true, true, pWorksheet.m_pImpl.m_bShowGridlines));
				else
					pBiffRecordContainer.AddBiffRecord(new Window2Record(false, false, pWorksheet.m_pImpl.m_bShowGridlines));
				pBiffRecordContainer.AddBiffRecord(new SelectionRecord());
				int nOffset = 0;
				while (nOffset < pWorksheet.m_pImpl.m_pMergedCellVector.GetSize())
				{
					int nSize = pWorksheet.m_pImpl.m_pMergedCellVector.GetSize() - nOffset;
					if (nSize > 1026)
						nSize = 1026;
					pBiffRecordContainer.AddBiffRecord(new MergeCellsRecord(pWorksheet.m_pImpl.m_pMergedCellVector, nOffset, nSize));
					nOffset += nSize;
				}
				pBiffRecordContainer.AddBiffRecord(new BiffRecord(BiffRecord.Type.TYPE_EOF, 0));
				{
					pRowRecordVector = null;
				}
				{
					pCellRecordVector = null;
				}
			}

			public static OfficeArtDimensions ComputeDimensions(Worksheet pWorksheet, ushort nCellX, ushort nSubCellX, ushort nWidth, ushort nCellY, ushort nSubCellY, ushort nHeight)
			{
				OfficeArtDimensions pDimensions = new OfficeArtDimensions();
				pDimensions.m_nCellX1 = nCellX;
				ushort nColumnWidth = pWorksheet.GetColumnWidth(pDimensions.m_nCellX1);
				if (nSubCellX > nColumnWidth)
					nSubCellX = nColumnWidth;
				pDimensions.m_nSubCellX1 = (ushort)((nSubCellX * 1024 + nColumnWidth / 2) / nColumnWidth);
				pDimensions.m_nCellY1 = nCellY;
				ushort nRowHeight = pWorksheet.GetRowHeight(pDimensions.m_nCellY1);
				if (nSubCellY > nRowHeight)
					nSubCellY = nRowHeight;
				pDimensions.m_nSubCellY1 = (ushort)((nSubCellY * 256 + nRowHeight / 2) / nRowHeight);
				pDimensions.m_nCellX2 = pDimensions.m_nCellX1;
				pDimensions.m_nSubCellX2 = (ushort)(nSubCellX + nWidth);
				while (pDimensions.m_nSubCellX2 > nColumnWidth)
				{
					pDimensions.m_nSubCellX2 = (ushort)(pDimensions.m_nSubCellX2 - nColumnWidth);
					pDimensions.m_nCellX2++;
					nColumnWidth = pWorksheet.GetColumnWidth(pDimensions.m_nCellX2);
				}
				pDimensions.m_nSubCellX2 = (ushort)((pDimensions.m_nSubCellX2 * 1024 + nColumnWidth / 2) / nColumnWidth);
				pDimensions.m_nCellY2 = pDimensions.m_nCellY1;
				pDimensions.m_nSubCellY2 = (ushort)(nSubCellY + nHeight);
				while (pDimensions.m_nSubCellY2 > nRowHeight)
				{
					pDimensions.m_nSubCellY2 = (ushort)(pDimensions.m_nSubCellY2 - nRowHeight);
					pDimensions.m_nCellY2++;
					nRowHeight = pWorksheet.GetRowHeight(pDimensions.m_nCellY2);
				}
				pDimensions.m_nSubCellY2 = (ushort)((pDimensions.m_nSubCellY2 * 256 + nRowHeight / 2) / nRowHeight);
				{
					NumberDuck.Secret.OfficeArtDimensions __2049702641 = pDimensions;
					pDimensions = null;
					return __2049702641;
				}
			}

			public int LoopToEnd(int i, BiffRecordContainer pBiffRecordContainer)
			{
				ushort nDepth = 1;
				while (nDepth > 0)
				{
					i++;
					BiffRecord pBiffRecord = pBiffRecordContainer.m_pBiffRecordVector.Get(i);
					if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_Begin)
						nDepth++;
					else if (pBiffRecord.GetType() == BiffRecord.Type.TYPE_End)
						nDepth--;
				}
				return i;
			}

			~BiffWorksheet()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffUtils
		{
			public static double RkValueDecode(uint nRkValue)
			{
				uint nMultiplierBit = nRkValue & 0x00000001;
				uint nTypeBit = (nRkValue & 0x00000002) >> 1;
				uint nEncodedValue = (nRkValue & 0xFFFFFFFC);
				double fTemp;
				if (nTypeBit == 0)
				{
					ulong nTemp = nEncodedValue;
					nTemp = nTemp << 32;
					fTemp = Utils.ByteConvertUint64ToDouble(nTemp);
				}
				else
				{
					nEncodedValue = nEncodedValue >> 2;
					fTemp = Utils.ByteConvertUint32ToInt32(nEncodedValue);
				}
				if (nMultiplierBit == 1)
					return fTemp / 100.0f;
				else
					return fTemp;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BiffRecordContainer : WorkbookGlobals
		{
			public OwnedVector<BiffRecord> m_pBiffRecordVector;
			public BiffRecordContainer()
			{
				m_pBiffRecordVector = new OwnedVector<BiffRecord>();
			}

			public BiffRecordContainer(BiffRecord pInitialBiffRecord, Stream pStream)
			{
				m_pBiffRecordVector = new OwnedVector<BiffRecord>();
				int nDepth = 1;
				AddBiffRecord(pInitialBiffRecord);
				while (true)
				{
					BiffRecord pBiffRecord = BiffRecord.CreateBiffRecord(pStream);
					BiffRecord.Type eType = pBiffRecord.GetType();
					{
						NumberDuck.Secret.BiffRecord __3036547922 = pBiffRecord;
						pBiffRecord = null;
						AddBiffRecord(__3036547922);
					}
					if (eType == BiffRecord.Type.TYPE_BOF)
					{
						nDepth++;
					}
					if (eType == BiffRecord.Type.TYPE_EOF)
					{
						nDepth--;
						if (nDepth == 0)
						{
							break;
						}
					}
				}
			}

			public void AddBiffRecord(BiffRecord pBiffRecord)
			{
				m_pBiffRecordVector.PushBack(pBiffRecord);
			}

			public uint GetSize()
			{
				uint nSize = 0;
				for (int i = 0; i < m_pBiffRecordVector.GetSize(); i++)
					nSize += m_pBiffRecordVector.Get(i).GetSize();
				return nSize;
			}

			public void Write(Stream pStream)
			{
				Blob pBlob = new Blob(true);
				BlobView pBlobView = pBlob.GetBlobView();
				for (int i = 0; i < m_pBiffRecordVector.GetSize(); i++)
				{
					pBlob.Resize(0, false);
					pBlobView.SetOffset(0);
					m_pBiffRecordVector.Get(i).Write(pStream, pBlobView);
				}
			}

			~BiffRecordContainer()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDirectoryImplementation
		{
			public uint m_nMinimumStandardStreamSize;
			public OwnedVector<Stream> m_pStreamVector;
			public CompoundFile m_pCompoundFile;
			public void AppendStream(Stream pStream)
			{
				nbAssert.Assert(m_pStreamVector.GetSize() == pStream.GetStreamId());
				m_pStreamVector.PushBack(pStream);
			}

			public static int RedBlackTreeComparisonCallback(object pObjectA, object pObjectB)
			{
				Stream pStreamA = (Stream)(pObjectA);
				Stream pStreamB = (Stream)(pObjectB);
				if (pStreamA.GetNameLengthUtf16() > pStreamB.GetNameLengthUtf16())
					return -1;
				else if (pStreamA.GetNameLengthUtf16() < pStreamB.GetNameLengthUtf16())
					return 1;
				ushort nLength = pStreamA.GetNameLengthUtf16();
				for (ushort i = 0; i < nLength; i++)
				{
					ushort a = pStreamA.GetNameUtf16(i);
					ushort b = pStreamB.GetNameUtf16(i);
					if (a > b)
						return -1;
					else if (a < b)
						return 1;
				}
				return 0;
			}

			public void RedBlackTreeWalk(RedBlackNode pNode, Stream pStorage)
			{
				Stream pStream = (Stream)(pNode.GetObject());
				if (pNode.GetParent() == null)
				{
					pStorage.SetLeftChildNodeStreamId(-1);
					pStorage.SetRightChildNodeStreamId(-1);
					pStorage.SetRootNodeStreamId(pStream.GetStreamId());
				}
				pStream.SetRootNodeStreamId(-1);
				pStream.SetNodeColour((byte)(pNode.GetColor()));
				for (int nDirection = 0; nDirection < 2; nDirection++)
				{
					if (pNode.GetChild(nDirection) != null)
					{
						Stream pChild = (Stream)(pNode.GetChild(nDirection).GetObject());
						if (nDirection == 0)
							pStream.SetLeftChildNodeStreamId(pChild.GetStreamId());
						else
							pStream.SetRightChildNodeStreamId(pChild.GetStreamId());
						this.RedBlackTreeWalk(pNode.GetChild(nDirection), pStorage);
					}
					else
					{
						if (nDirection == 0)
							pStream.SetLeftChildNodeStreamId(-1);
						else
							pStream.SetRightChildNodeStreamId(-1);
					}
				}
			}

			~StreamDirectoryImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StreamDirectory : SectorChain
		{
			protected StreamDirectoryImplementation m_pImpl;
			public StreamDirectory(int nSectorSize, uint nMinimumStandardStreamSize, CompoundFile pCompoundFile) : base(nSectorSize)
			{
				m_pImpl = new StreamDirectoryImplementation();
				m_pImpl.m_nMinimumStandardStreamSize = nMinimumStandardStreamSize;
				m_pImpl.m_pStreamVector = new OwnedVector<Stream>();
				m_pImpl.m_pCompoundFile = pCompoundFile;
			}

			public int GetNumStream()
			{
				return m_pImpl.m_pStreamVector.GetSize();
			}

			public Stream GetStreamByIndex(int nStreamDirectoryId)
			{
				nbAssert.Assert(nStreamDirectoryId >= 0);
				nbAssert.Assert((uint)(nStreamDirectoryId) < GetNumStream());
				return m_pImpl.m_pStreamVector.Get(nStreamDirectoryId);
			}

			public Stream GetStreamByName(string sxName)
			{
				for (int i = 0; i < m_pImpl.m_pStreamVector.GetSize(); i++)
				{
					Stream pStream = m_pImpl.m_pStreamVector.Get(i);
					if (ExternalString.Equal(pStream.GetName(), sxName))
						return pStream;
				}
				return null;
			}

			public Stream CreateStream(string sxName, Stream.Type eType)
			{
				if (eType == Stream.Type.TYPE_EMPTY)
					return null;
				for (int i = 0; i < m_pImpl.m_pStreamVector.GetSize(); i++)
				{
					if (m_pImpl.m_pStreamVector.Get(i).GetType() == Stream.Type.TYPE_EMPTY)
					{
						Stream pStream = m_pImpl.m_pStreamVector.Get(i);
						pStream.Allocate(eType, 0);
						pStream.SetName(sxName);
						RedBlackTreeRebuild();
						return pStream;
					}
				}
				m_pImpl.m_pCompoundFile.StreamDirectoryExtend();
				return CreateStream(sxName, eType);
			}

			public override void AppendSector(Sector pSector)
			{
				base.AppendSector(pSector);
				BlobView pBlobView = GetBlobView();
				int nOffset = pBlobView.GetSize() - pSector.GetDataSize();
				while (nOffset < pBlobView.GetSize())
				{
					m_pImpl.AppendStream(new Stream(GetNumStream(), (int)(m_pImpl.m_nMinimumStandardStreamSize), m_pBlob, nOffset, m_pImpl.m_pCompoundFile));
					nOffset += Stream.DATA_SIZE;
				}
			}

			public override void Extend(Sector pSector)
			{
				AppendSector(pSector);
			}

			public void RedBlackTreeRebuild()
			{
				RedBlackTree pRedBlackTree = new RedBlackTree(StreamDirectoryImplementation.RedBlackTreeComparisonCallback);
				for (int i = 1; i < m_pImpl.m_pStreamVector.GetSize(); i++)
				{
					Stream pStream = m_pImpl.m_pStreamVector.Get(i);
					if (pStream.GetType() != Stream.Type.TYPE_EMPTY)
						nbAssert.Assert(pRedBlackTree.AddObject(pStream));
				}
				m_pImpl.RedBlackTreeWalk(pRedBlackTree.GetRootNode(), m_pImpl.m_pStreamVector.Get(0));
				{
					pRedBlackTree = null;
				}
			}

			~StreamDirectory()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorAllocationTable : SectorChain
		{
			protected int m_nNumFreeSectorId;
			public SectorAllocationTable(int nSectorSize) : base(nSectorSize)
			{
				m_nNumFreeSectorId = 0;
			}

			public virtual int GetNumSectorId()
			{
				return GetDataSize() >> 2;
			}

			public virtual int GetNumFreeSectorId()
			{
				return m_nNumFreeSectorId;
			}

			public virtual int GetSectorId(int nIndex)
			{
				nbAssert.Assert(nIndex >= 0);
				nbAssert.Assert(nIndex < GetNumSectorId());
				BlobView pBlobView = GetBlobView();
				pBlobView.SetOffset(nIndex << 2);
				return pBlobView.UnpackInt32();
			}

			public virtual void SetSectorId(int nIndex, int nSectorId)
			{
				BlobView pBlobView = GetBlobView();
				nbAssert.Assert(nIndex >= 0);
				nbAssert.Assert(nIndex < GetNumSectorId());
				int nLastSectorId = pBlobView.UnpackInt32At(nIndex << 2);
				if (nSectorId != nLastSectorId)
				{
					if (nSectorId == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
						m_nNumFreeSectorId++;
					if (nLastSectorId == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
						m_nNumFreeSectorId--;
					pBlobView.SetOffset(nIndex << 2);
					pBlobView.PackInt32(nSectorId);
				}
			}

			public int GetFreeSectorId()
			{
				int nNumSectorId = GetNumSectorId();
				for (int i = nNumSectorId - 1; i >= 0; i--)
					if (GetSectorId(i) == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
						return i;
				return -1;
			}

			public override void AppendSector(Sector pSector)
			{
				base.AppendSector(pSector);
				BlobView pBlobView = pSector.GetBlobView();
				pBlobView.SetOffset(0);
				while (pBlobView.GetOffset() < pBlobView.GetSize())
				{
					if (pBlobView.UnpackInt32() == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
						m_nNumFreeSectorId++;
				}
			}

			public override void Extend(Sector pSector)
			{
				BlobView pBlobView = pSector.GetBlobView();
				pBlobView.SetOffset(0);
				while (pBlobView.GetOffset() < pBlobView.GetSize())
				{
					pBlobView.PackInt32((int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID));
				}
				AppendSector(pSector);
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SectorImplementation
		{
			public int m_nSectorId;
			public BlobView m_pBlobView;
			~SectorImplementation()
			{
			}

		}
		class Sector
		{
			public const int MINIMUM_SECTOR_SIZE = 128;
			public const int MINIMUM_SECTOR_SIZE_SHIFT = 7;
			protected SectorImplementation m_pImpl;
			public enum SectorId
			{
				FREE_SECTOR_SECTOR_ID = -1,
				END_OF_CHAIN_SECTOR_ID = -2,
				SECTOR_ALLOCATION_TABLE_SECTOR_ID = -3,
				MASTER_SECTOR_ALLOCATION_TABLE_SECTOR_ID = -4,
			}

			public Sector(int nSectorId, BlobView pBlobView, int nDataSize)
			{
				nbAssert.Assert(pBlobView != null);
				nbAssert.Assert(((nDataSize & (nDataSize - 1)) == 0));
				m_pImpl = new SectorImplementation();
				m_pImpl.m_nSectorId = nSectorId;
				m_pImpl.m_pBlobView = new BlobView(pBlobView.GetBlob(), pBlobView.GetOffset(), pBlobView.GetOffset() + nDataSize);
				pBlobView.SetOffset(pBlobView.GetOffset() + nDataSize);
			}

			public int GetDataSize()
			{
				return m_pImpl.m_pBlobView.GetSize();
			}

			public int GetSectorId()
			{
				return m_pImpl.m_nSectorId;
			}

			public BlobView GetBlobView()
			{
				return m_pImpl.m_pBlobView;
			}

			~Sector()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MasterSectorAllocationTable : SectorAllocationTable
		{
			public const int INITIAL_SECTOR_ID_ARRAY_SIZE = 109;
			protected CompoundFileHeader m_pHeader;
			public MasterSectorAllocationTable(CompoundFileHeader pHeader) : base(pHeader.GetSectorSize())
			{
				m_pHeader = pHeader;
			}

			public override int GetNumSectorId()
			{
				return INITIAL_SECTOR_ID_ARRAY_SIZE + GetNumSector() * ((m_nSectorSize >> 2) - 1);
			}

			public int GetNumInternalSectorId()
			{
				return INITIAL_SECTOR_ID_ARRAY_SIZE + GetNumSector() * (m_nSectorSize >> 2);
			}

			public override int GetSectorId(int nIndex)
			{
				int nTranslatedIndex = TranslateIndex(nIndex);
				if (nTranslatedIndex == nIndex)
					return m_pHeader.m_pMasterSectorAllocationTable[nTranslatedIndex];
				else
					return base.GetSectorId(nTranslatedIndex);
			}

			public override void SetSectorId(int nIndex, int nSectorId)
			{
				int nTranslatedIndex = TranslateIndex(nIndex);
				if (nTranslatedIndex == nIndex)
				{
					if (nSectorId != m_pHeader.m_pMasterSectorAllocationTable[nTranslatedIndex])
					{
						if (nSectorId == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
							m_nNumFreeSectorId++;
						if (m_pHeader.m_pMasterSectorAllocationTable[nTranslatedIndex] == (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID))
							m_nNumFreeSectorId--;
					}
					m_pHeader.m_pMasterSectorAllocationTable[nTranslatedIndex] = nSectorId;
				}
				else
				{
					base.SetSectorId(nTranslatedIndex, nSectorId);
				}
			}

			public int GetInternalSectorId(int nIndex)
			{
				nbAssert.Assert(nIndex >= 0);
				if (nIndex < INITIAL_SECTOR_ID_ARRAY_SIZE)
					return m_pHeader.m_pMasterSectorAllocationTable[nIndex];
				else
					return base.GetSectorId(nIndex - INITIAL_SECTOR_ID_ARRAY_SIZE);
			}

			public void SetInternalSectorId(int nIndex, int nSectorId)
			{
				nbAssert.Assert(nIndex >= 0);
				if (nIndex < INITIAL_SECTOR_ID_ARRAY_SIZE)
					m_pHeader.m_pMasterSectorAllocationTable[nIndex] = nSectorId;
				else
					base.SetSectorId(nIndex - INITIAL_SECTOR_ID_ARRAY_SIZE, nSectorId);
			}

			public int TranslateIndex(int nIndex)
			{
				nbAssert.Assert(nIndex >= 0);
				if (nIndex >= INITIAL_SECTOR_ID_ARRAY_SIZE)
				{
					int nSectorIndex = nIndex - INITIAL_SECTOR_ID_ARRAY_SIZE;
					nSectorIndex += nSectorIndex / ((m_nSectorSize >> 2) - 1);
					return nSectorIndex;
				}
				return nIndex;
			}

			public int GetSectorIdToAppend()
			{
				if (GetNumSector() == 0)
					return m_pHeader.m_nMasterSectorAllocationTableSectorId;
				BlobView pBlobView = GetBlobView();
				return pBlobView.UnpackInt32At(GetDataSize() - (1 << 2));
			}

			public override void AppendSector(Sector pSector)
			{
				nbAssert.Assert(GetSectorIdToAppend() == pSector.GetSectorId());
				base.AppendSector(pSector);
			}

			public override void Extend(Sector pSector)
			{
				int nNumSectorIdPerSector = pSector.GetDataSize() >> 2;
				BlobView pBlobView = pSector.GetBlobView();
				pBlobView.SetOffset(0);
				for (int i = 0; i < nNumSectorIdPerSector - 1; i++)
					pBlobView.PackInt32((int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID));
				pBlobView.PackInt32((int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID));
				base.AppendSector(pSector);
				int nIndex = GetNumInternalSectorId() - nNumSectorIdPerSector - 1;
				if (nIndex > INITIAL_SECTOR_ID_ARRAY_SIZE)
					SetInternalSectorId(nIndex, pSector.GetSectorId());
				else
					m_pHeader.m_nMasterSectorAllocationTableSectorId = pSector.GetSectorId();
				m_pHeader.m_nMasterSectorAllocationTableSize++;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CompoundFileHeader
		{
			public const int MAGIC_WORD_SIZE = 8;
			public static readonly byte[] MAGIC_WORD = {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1};
			public const uint UNIQUE_IDENTIFIER_SIZE = 16;
			public const ushort REVISION_NUMBER = 62;
			public const ushort VERSION_NUMBER = 0x0003;
			public const ushort BOM_LITTLE_ENDIAN = 0xFFFF;
			public const ushort BOM_BIG_ENDIAN = 0xFFFE;
			public byte[] m_pMagicWord = new byte[MAGIC_WORD_SIZE];
			public byte[] m_pUniqueIdentifier = new byte[UNIQUE_IDENTIFIER_SIZE];
			public ushort m_nRevisonNumber;
			public ushort m_nVersionNumber;
			public ushort m_nByteOrderIdentifier;
			public ushort m_nSectorSize;
			public ushort m_nShortSectorSize;
			public byte[] m_pUnusedA = new byte[10];
			public uint m_nSectorAllocationTableSize;
			public int m_nStreamDirectoryStreamSectorId;
			public byte[] m_pUnusedB = new byte[4];
			public uint m_nMinimumStandardStreamSize;
			public int m_nShortSectorAllocationTableSectorId;
			public uint m_nShortSectorAllocationTableSize;
			public int m_nMasterSectorAllocationTableSectorId;
			public uint m_nMasterSectorAllocationTableSize;
			public int[] m_pMasterSectorAllocationTable = new int[MasterSectorAllocationTable.INITIAL_SECTOR_ID_ARRAY_SIZE];
			public CompoundFileHeader(uint nSectorSize, uint nShortSectorSize)
			{
				uint i;
				for (i = 0; i < MAGIC_WORD_SIZE; i++)
					m_pMagicWord[i] = MAGIC_WORD[i];
				for (i = 0; i < UNIQUE_IDENTIFIER_SIZE; i++)
					m_pUniqueIdentifier[i] = 0;
				m_nRevisonNumber = REVISION_NUMBER;
				m_nVersionNumber = VERSION_NUMBER;
				m_nByteOrderIdentifier = BOM_BIG_ENDIAN;
				uint nTemp = nSectorSize;
				m_nSectorSize = 0;
				while (nTemp != 1)
				{
					nTemp = nTemp >> 1;
					m_nSectorSize++;
				}
				nTemp = nShortSectorSize;
				m_nShortSectorSize = 0;
				while (nTemp != 1)
				{
					nTemp = nTemp >> 1;
					m_nShortSectorSize++;
				}
				for (i = 0; i < 10; i++)
					m_pUnusedA[i] = 0;
				m_nSectorAllocationTableSize = 0;
				m_nStreamDirectoryStreamSectorId = 0;
				for (i = 0; i < 4; i++)
					m_pUnusedB[i] = 0;
				m_nMinimumStandardStreamSize = 0;
				m_nShortSectorAllocationTableSectorId = 0;
				m_nShortSectorAllocationTableSize = 0;
				m_nMasterSectorAllocationTableSectorId = 0;
				m_nMasterSectorAllocationTableSize = 0;
				for (i = 0; i < MasterSectorAllocationTable.INITIAL_SECTOR_ID_ARRAY_SIZE; i++)
					m_pMasterSectorAllocationTable[i] = (int)(Sector.SectorId.FREE_SECTOR_SECTOR_ID);
			}

			public int GetSectorSize()
			{
				return 1 << m_nSectorSize;
			}

			public bool Unpack(Blob pBlob)
			{
				uint i;
				BlobView pBlobView = pBlob.GetBlobView();
				for (i = 0; i < MAGIC_WORD_SIZE; i++)
				{
					m_pMagicWord[i] = pBlobView.UnpackUint8();
					if (m_pMagicWord[i] != MAGIC_WORD[i])
						return false;
				}
				for (i = 0; i < UNIQUE_IDENTIFIER_SIZE; i++)
					m_pUniqueIdentifier[i] = pBlobView.UnpackUint8();
				m_nRevisonNumber = pBlobView.UnpackUint16();
				m_nVersionNumber = pBlobView.UnpackUint16();
				m_nByteOrderIdentifier = pBlobView.UnpackUint16();
				m_nSectorSize = pBlobView.UnpackUint16();
				m_nShortSectorSize = pBlobView.UnpackUint16();
				for (i = 0; i < 10; i++)
					m_pUnusedA[i] = pBlobView.UnpackUint8();
				m_nSectorAllocationTableSize = pBlobView.UnpackUint32();
				m_nStreamDirectoryStreamSectorId = pBlobView.UnpackInt32();
				for (i = 0; i < 4; i++)
					m_pUnusedB[i] = pBlobView.UnpackUint8();
				m_nMinimumStandardStreamSize = pBlobView.UnpackUint32();
				m_nShortSectorAllocationTableSectorId = pBlobView.UnpackInt32();
				m_nShortSectorAllocationTableSize = pBlobView.UnpackUint32();
				m_nMasterSectorAllocationTableSectorId = pBlobView.UnpackInt32();
				m_nMasterSectorAllocationTableSize = pBlobView.UnpackUint32();
				for (i = 0; i < MasterSectorAllocationTable.INITIAL_SECTOR_ID_ARRAY_SIZE; i++)
					m_pMasterSectorAllocationTable[i] = pBlobView.UnpackInt32();
				return true;
			}

			public void Pack(Blob pBlob)
			{
				uint i;
				BlobView pBlobView = pBlob.GetBlobView();
				for (i = 0; i < MAGIC_WORD_SIZE; i++)
					pBlobView.PackUint8(m_pMagicWord[i]);
				for (i = 0; i < UNIQUE_IDENTIFIER_SIZE; i++)
					pBlobView.PackUint8(m_pUniqueIdentifier[i]);
				pBlobView.PackUint16(m_nRevisonNumber);
				pBlobView.PackUint16(m_nVersionNumber);
				pBlobView.PackUint16(m_nByteOrderIdentifier);
				pBlobView.PackUint16(m_nSectorSize);
				pBlobView.PackUint16(m_nShortSectorSize);
				for (i = 0; i < 10; i++)
					pBlobView.PackUint8(m_pUnusedA[i]);
				pBlobView.PackUint32(m_nSectorAllocationTableSize);
				pBlobView.PackInt32(m_nStreamDirectoryStreamSectorId);
				for (i = 0; i < 4; i++)
					pBlobView.PackUint8(m_pUnusedB[i]);
				pBlobView.PackUint32(m_nMinimumStandardStreamSize);
				pBlobView.PackUint32((uint)(m_nShortSectorAllocationTableSectorId));
				pBlobView.PackUint32(m_nShortSectorAllocationTableSize);
				pBlobView.PackInt32(m_nMasterSectorAllocationTableSectorId);
				pBlobView.PackUint32(m_nMasterSectorAllocationTableSize);
				for (i = 0; i < MasterSectorAllocationTable.INITIAL_SECTOR_ID_ARRAY_SIZE; i++)
					pBlobView.PackUint32((uint)(m_pMasterSectorAllocationTable[i]));
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CompoundFile
		{
			public const int HEADER_SIZE = 512;
			protected CompoundFileHeader m_pHeader;
			protected Blob m_pBlob;
			protected MasterSectorAllocationTable m_pMasterSectorAllocationTable;
			protected SectorAllocationTable m_pSectorAllocationTable;
			protected SectorAllocationTable m_pShortSectorAllocationTable;
			protected StreamDirectory m_pStreamDirectory;
			protected OwnedVector<Sector> m_pSectorVector;
			protected OwnedVector<Sector> m_pShortSectorVector;
			public CompoundFile(uint nSectorSize = 512, uint nShortSectorSize = 64, uint nMinimumStandardStreamSize = 4096)
			{
				nbAssert.Assert(nSectorSize >= Sector.MINIMUM_SECTOR_SIZE && ((nSectorSize & (nSectorSize - 1)) == 0));
				nbAssert.Assert(nShortSectorSize < nSectorSize && ((nShortSectorSize & (nShortSectorSize - 1)) == 0));
				m_pSectorVector = new OwnedVector<Sector>();
				m_pShortSectorVector = new OwnedVector<Sector>();
				m_pBlob = new Blob(false);
				m_pBlob.Resize(HEADER_SIZE, false);
				m_pHeader = new CompoundFileHeader(nSectorSize, nShortSectorSize);
				m_pMasterSectorAllocationTable = new MasterSectorAllocationTable(m_pHeader);
				m_pHeader.m_nMasterSectorAllocationTableSectorId = (int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID);
				m_pSectorAllocationTable = new SectorAllocationTable(1 << m_pHeader.m_nSectorSize);
				SectorAllocationTableExtend(false);
				m_pHeader.m_nMinimumStandardStreamSize = nMinimumStandardStreamSize;
				m_pStreamDirectory = new StreamDirectory((int)(nSectorSize), m_pHeader.m_nMinimumStandardStreamSize, this);
				StreamDirectoryExtend();
				Stream pStream = m_pStreamDirectory.GetStreamByIndex(0);
				pStream.SetName("Root Entry");
				m_pShortSectorAllocationTable = new SectorAllocationTable(1 << m_pHeader.m_nSectorSize);
				SectorAllocationTableExtend(true);
			}

			~CompoundFile()
			{
			}

			public bool Load(string sxFileName)
			{
				{
					m_pMasterSectorAllocationTable = null;
				}
				m_pMasterSectorAllocationTable = null;
				{
					m_pSectorAllocationTable = null;
				}
				m_pSectorAllocationTable = null;
				{
					m_pShortSectorAllocationTable = null;
				}
				m_pShortSectorAllocationTable = null;
				{
					m_pStreamDirectory = null;
				}
				m_pStreamDirectory = null;
				m_pSectorVector.Clear();
				m_pShortSectorVector.Clear();
				{
					m_pBlob = null;
				}
				m_pBlob = new Blob(false);
				if (m_pBlob.Load(sxFileName))
				{
					if (m_pHeader.Unpack(m_pBlob))
					{
						int nSectorId;
						int nSectorSize = (1 << m_pHeader.m_nSectorSize);
						int nNumSector = (m_pBlob.GetSize() - HEADER_SIZE) / nSectorSize;
						BlobView pBlobView = m_pBlob.GetBlobView();
						nbAssert.Assert(pBlobView.GetOffset() == HEADER_SIZE);
						nbAssert.Assert(pBlobView.GetOffset() + nNumSector * nSectorSize == m_pBlob.GetSize());
						for (int i = 0; i < nNumSector; i++)
							m_pSectorVector.PushBack(new Sector(i, pBlobView, nSectorSize));
						m_pMasterSectorAllocationTable = new MasterSectorAllocationTable(m_pHeader);
						nSectorId = m_pMasterSectorAllocationTable.GetSectorIdToAppend();
						while (nSectorId >= 0)
						{
							m_pMasterSectorAllocationTable.AppendSector(m_pSectorVector.Get(nSectorId));
							nSectorId = m_pMasterSectorAllocationTable.GetSectorIdToAppend();
						}
						nbAssert.Assert(m_pMasterSectorAllocationTable.GetNumSector() == (int)(m_pHeader.m_nMasterSectorAllocationTableSize));
						m_pSectorAllocationTable = new SectorAllocationTable(1 << m_pHeader.m_nSectorSize);
						for (int i = 0; i < m_pHeader.m_nSectorAllocationTableSize; i++)
							m_pSectorAllocationTable.AppendSector(GetSector(m_pMasterSectorAllocationTable.GetSectorId(i), false));
						nbAssert.Assert(m_pSectorAllocationTable.GetNumSector() == (int)(m_pHeader.m_nSectorAllocationTableSize));
						m_pShortSectorAllocationTable = new SectorAllocationTable(1 << m_pHeader.m_nSectorSize);
						FillSectorChain(m_pShortSectorAllocationTable, m_pHeader.m_nShortSectorAllocationTableSectorId, false);
						nbAssert.Assert(m_pShortSectorAllocationTable.GetNumSector() == (int)(m_pHeader.m_nShortSectorAllocationTableSize));
						m_pStreamDirectory = new StreamDirectory((int)(nSectorSize), m_pHeader.m_nMinimumStandardStreamSize, this);
						FillSectorChain(m_pStreamDirectory, m_pHeader.m_nStreamDirectoryStreamSectorId, false);
						Stream pRootStream = m_pStreamDirectory.GetStreamByIndex(0);
						nbAssert.Assert(pRootStream.GetType() == Stream.Type.TYPE_ROOT_STORAGE);
						pRootStream.FillSectorChain();
						int nShortSectorSize = (1 << m_pHeader.m_nShortSectorSize);
						int nNumShortSector = pRootStream.GetSize() / nShortSectorSize;
						pRootStream.GetSectorChain().SetOffset(0);
						for (int i = 0; i < nNumShortSector; i++)
						{
							Sector pSector = new Sector(i, pRootStream.GetSectorChain().GetBlobView(), nShortSectorSize);
							{
								NumberDuck.Secret.Sector __3878760436 = pSector;
								pSector = null;
								m_pShortSectorVector.PushBack(__3878760436);
							}
						}
						for (int i = 1; i < m_pStreamDirectory.GetNumStream(); i++)
						{
							Stream pStream = m_pStreamDirectory.GetStreamByIndex(i);
							if (pStream.GetType() != Stream.Type.TYPE_EMPTY)
							{
								pStream.FillSectorChain();
							}
						}
						return true;
					}
				}
				return false;
			}

			public bool Save(string sxFileName)
			{
				for (int i = 0; i < m_pStreamDirectory.GetNumStream(); i++)
					m_pStreamDirectory.GetStreamByIndex(i).WriteToSectors();
				m_pStreamDirectory.GetStreamByIndex(0).WriteToSectors();
				m_pMasterSectorAllocationTable.WriteToSectors();
				m_pSectorAllocationTable.WriteToSectors();
				m_pShortSectorAllocationTable.WriteToSectors();
				m_pStreamDirectory.WriteToSectors();
				m_pBlob.GetBlobView().SetOffset(0);
				m_pHeader.Pack(m_pBlob);
				bool bResult = m_pBlob.Save(sxFileName);
				return bResult;
			}

			public CompoundFileHeader GetHeader()
			{
				return m_pHeader;
			}

			public StreamDirectory GetStreamDirectory()
			{
				return m_pStreamDirectory;
			}

			public Stream CreateStream(string sxName, Stream.Type eType)
			{
				return m_pStreamDirectory.CreateStream(sxName, eType);
			}

			public int GetNumStream()
			{
				return m_pStreamDirectory.GetNumStream();
			}

			public Stream GetStreamByIndex(int nIndex)
			{
				return m_pStreamDirectory.GetStreamByIndex(nIndex);
			}

			public Stream GetStreamByName(string sxName)
			{
				return m_pStreamDirectory.GetStreamByName(sxName);
			}

			public int GetSectorSize(bool bShortSector)
			{
				if (bShortSector)
					return 1 << m_pHeader.m_nShortSectorSize;
				else
					return 1 << m_pHeader.m_nSectorSize;
			}

			public int GetSectorId(int nSectorId, bool bShortSector)
			{
				nbAssert.Assert(nSectorId >= 0);
				if (bShortSector)
				{
					return m_pShortSectorAllocationTable.GetSectorId(nSectorId);
				}
				else
				{
					return m_pSectorAllocationTable.GetSectorId(nSectorId);
				}
			}

			public Sector GetSector(int nSectorId, bool bShortSector)
			{
				nbAssert.Assert(nSectorId >= 0);
				if (bShortSector)
				{
					nbAssert.Assert(nSectorId < m_pShortSectorVector.GetSize());
					return m_pShortSectorVector.Get(nSectorId);
				}
				else
				{
					nbAssert.Assert(nSectorId < m_pSectorVector.GetSize());
					return m_pSectorVector.Get(nSectorId);
				}
			}

			public void FillSectorChain(SectorChain pSectorChain, int nInitialSectorId, bool bShortSector)
			{
				nbAssert.Assert(pSectorChain != null);
				int nSectorId = nInitialSectorId;
				while (nSectorId != (int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID))
				{
					pSectorChain.AppendSector(GetSector(nSectorId, bShortSector));
					nSectorId = GetSectorId(nSectorId, bShortSector);
				}
			}

			public int GetFreeSectorId(bool bShortSector)
			{
				if (bShortSector)
					return m_pShortSectorAllocationTable.GetFreeSectorId();
				return m_pSectorAllocationTable.GetFreeSectorId();
			}

			public SectorAllocationTable GetSectorAllocationTable(bool bShortSector)
			{
				if (bShortSector)
					return m_pShortSectorAllocationTable;
				else
					return m_pSectorAllocationTable;
			}

			public MasterSectorAllocationTable GetMasterSectorAllocationTable()
			{
				return m_pMasterSectorAllocationTable;
			}

			public void MasterSectorAllocationTableExtend()
			{
				int nFreeSectorId = GetFreeSectorId(false);
				if (nFreeSectorId < 0)
				{
					int nSectorSize = GetSectorSize(false);
					BlobView pBlobView = m_pBlob.GetBlobView();
					m_pBlob.Resize(HEADER_SIZE + (m_pSectorVector.GetSize() + 1) * nSectorSize, false);
					pBlobView.SetOffset(HEADER_SIZE + m_pSectorVector.GetSize() * nSectorSize);
					nFreeSectorId = m_pSectorVector.GetSize();
					m_pSectorVector.PushBack(new Sector(nFreeSectorId, pBlobView, GetSectorSize(false)));
				}
				m_pMasterSectorAllocationTable.Extend(m_pSectorVector.Get(nFreeSectorId));
				if (m_pSectorAllocationTable.GetNumFreeSectorId() == 0)
					SectorAllocationTableExtend(false);
				m_pSectorAllocationTable.SetSectorId(nFreeSectorId, (int)(Sector.SectorId.MASTER_SECTOR_ALLOCATION_TABLE_SECTOR_ID));
			}

			public void SectorAllocationTableExtend(bool bShortSector)
			{
				int nSectorSize = GetSectorSize(bShortSector);
				if (bShortSector)
				{
					Stream pRootStream = m_pStreamDirectory.GetStreamByIndex(0);
					SectorChainExtend(m_pShortSectorAllocationTable);
					m_pHeader.m_nShortSectorAllocationTableSectorId = m_pShortSectorAllocationTable.GetSectorByIndex(0).GetSectorId();
					m_pHeader.m_nShortSectorAllocationTableSize = (uint)(m_pShortSectorAllocationTable.GetNumSector());
					pRootStream.Resize(m_pShortSectorAllocationTable.GetNumSectorId() * nSectorSize);
					BlobView pBlobView = pRootStream.GetSectorChain().GetBlobView();
					pBlobView.SetOffset(m_pShortSectorVector.GetSize() * nSectorSize);
					while (m_pShortSectorVector.GetSize() != m_pShortSectorAllocationTable.GetNumSectorId())
					{
						Sector pSector = new Sector(m_pShortSectorVector.GetSize(), pBlobView, nSectorSize);
						{
							NumberDuck.Secret.Sector __3878760436 = pSector;
							pSector = null;
							m_pShortSectorVector.PushBack(__3878760436);
						}
					}
				}
				else
				{
					BlobView pBlobView = m_pBlob.GetBlobView();
					if (m_pMasterSectorAllocationTable.GetNumSectorId() == m_pSectorAllocationTable.GetNumSector())
						MasterSectorAllocationTableExtend();
					m_pBlob.Resize(HEADER_SIZE + (m_pSectorVector.GetSize() + 1) * nSectorSize, false);
					pBlobView.SetOffset(HEADER_SIZE + m_pSectorVector.GetSize() * nSectorSize);
					int nSectorId = m_pSectorVector.GetSize();
					Sector pSector = new Sector(nSectorId, pBlobView, nSectorSize);
					Sector pTemp = pSector;
					{
						NumberDuck.Secret.Sector __3878760436 = pSector;
						pSector = null;
						m_pSectorVector.PushBack(__3878760436);
					}
					m_pMasterSectorAllocationTable.SetSectorId(m_pSectorAllocationTable.GetNumSector(), nSectorId);
					m_pSectorAllocationTable.Extend(pTemp);
					if (m_pSectorVector.GetSize() > 1)
						m_pSectorAllocationTable.SetSectorId(m_pSectorVector.Get(m_pSectorVector.GetSize() - 1).GetSectorId(), nSectorId);
					m_pSectorAllocationTable.SetSectorId(nSectorId, (int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID));
					m_pHeader.m_nSectorAllocationTableSize++;
					m_pBlob.Resize(HEADER_SIZE + m_pSectorAllocationTable.GetNumSectorId() * nSectorSize, false);
					while (m_pSectorVector.GetSize() != m_pSectorAllocationTable.GetNumSectorId())
					{
						m_pSectorVector.PushBack(new Sector(m_pSectorVector.GetSize(), pBlobView, nSectorSize));
					}
				}
			}

			public void SectorChainExtend(SectorChain pSectorChain)
			{
				bool bShortSector = pSectorChain.GetSectorSize() == GetSectorSize(true);
				if (GetSectorAllocationTable(bShortSector).GetNumFreeSectorId() == 0)
					SectorAllocationTableExtend(bShortSector);
				int nFreeSectorId = GetFreeSectorId(bShortSector);
				Sector pSector = GetSector(nFreeSectorId, bShortSector);
				pSectorChain.Extend(pSector);
				SectorAllocationTable pSectorAllocationTable = GetSectorAllocationTable(bShortSector);
				if (pSectorChain.GetNumSector() > 1)
					pSectorAllocationTable.SetSectorId(pSectorChain.GetSectorByIndex(pSectorChain.GetNumSector() - 2).GetSectorId(), pSector.GetSectorId());
				pSectorAllocationTable.SetSectorId(pSector.GetSectorId(), (int)(Sector.SectorId.END_OF_CHAIN_SECTOR_ID));
			}

			public void StreamDirectoryExtend()
			{
				SectorChainExtend(m_pStreamDirectory);
				m_pHeader.m_nStreamDirectoryStreamSectorId = m_pStreamDirectory.GetSectorByIndex(0).GetSectorId();
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Token
		{
			public enum Type
			{
				TYPE_FUNC_COUNT = 0x0000,
				TYPE_FUNC_IF = 0x0001,
				TYPE_FUNC_ISNA = 0x0002,
				TYPE_FUNC_ISERROR = 0x0003,
				TYPE_FUNC_SUM = 0x0004,
				TYPE_FUNC_AVERAGE = 0x0005,
				TYPE_FUNC_MIN = 0x0006,
				TYPE_FUNC_MAX = 0x0007,
				TYPE_FUNC_ROW = 0x0008,
				TYPE_FUNC_COLUMN = 0x0009,
				TYPE_FUNC_NA = 0x000A,
				TYPE_FUNC_NPV = 0x000B,
				TYPE_FUNC_STDEV = 0x000C,
				TYPE_FUNC_DOLLAR = 0x000D,
				TYPE_FUNC_FIXED = 0x000E,
				TYPE_FUNC_SIN = 0x000F,
				TYPE_FUNC_COS = 0x0010,
				TYPE_FUNC_TAN = 0x0011,
				TYPE_FUNC_ATAN = 0x0012,
				TYPE_FUNC_PI = 0x0013,
				TYPE_FUNC_SQRT = 0x0014,
				TYPE_FUNC_EXP = 0x0015,
				TYPE_FUNC_LN = 0x0016,
				TYPE_FUNC_LOG10 = 0x0017,
				TYPE_FUNC_ABS = 0x0018,
				TYPE_FUNC_INT = 0x0019,
				TYPE_FUNC_SIGN = 0x001A,
				TYPE_FUNC_ROUND = 0x001B,
				TYPE_FUNC_LOOKUP = 0x001C,
				TYPE_FUNC_INDEX = 0x001D,
				TYPE_FUNC_REPT = 0x001E,
				TYPE_FUNC_MID = 0x001F,
				TYPE_FUNC_LEN = 0x0020,
				TYPE_FUNC_VALUE = 0x0021,
				TYPE_FUNC_TRUE = 0x0022,
				TYPE_FUNC_FALSE = 0x0023,
				TYPE_FUNC_AND = 0x0024,
				TYPE_FUNC_OR = 0x0025,
				TYPE_FUNC_NOT = 0x0026,
				TYPE_FUNC_MOD = 0x0027,
				TYPE_FUNC_DCOUNT = 0x0028,
				TYPE_FUNC_DSUM = 0x0029,
				TYPE_FUNC_DAVERAGE = 0x002A,
				TYPE_FUNC_DMIN = 0x002B,
				TYPE_FUNC_DMAX = 0x002C,
				TYPE_FUNC_DSTDEV = 0x002D,
				TYPE_FUNC_VAR = 0x002E,
				TYPE_FUNC_DVAR = 0x002F,
				TYPE_FUNC_TEXT = 0x0030,
				TYPE_FUNC_LINEST = 0x0031,
				TYPE_FUNC_TREND = 0x0032,
				TYPE_FUNC_LOGEST = 0x0033,
				TYPE_FUNC_GROWTH = 0x0034,
				TYPE_FUNC_GOTO = 0x0035,
				TYPE_FUNC_HALT = 0x0036,
				TYPE_FUNC_RETURN = 0x0037,
				TYPE_FUNC_PV = 0x0038,
				TYPE_FUNC_FV = 0x0039,
				TYPE_FUNC_NPER = 0x003A,
				TYPE_FUNC_PMT = 0x003B,
				TYPE_FUNC_RATE = 0x003C,
				TYPE_FUNC_MIRR = 0x003D,
				TYPE_FUNC_IRR = 0x003E,
				TYPE_FUNC_RAND = 0x003F,
				TYPE_FUNC_MATCH = 0x0040,
				TYPE_FUNC_DATE = 0x0041,
				TYPE_FUNC_TIME = 0x0042,
				TYPE_FUNC_DAY = 0x0043,
				TYPE_FUNC_MONTH = 0x0044,
				TYPE_FUNC_YEAR = 0x0045,
				TYPE_FUNC_WEEKDAY = 0x0046,
				TYPE_FUNC_HOUR = 0x0047,
				TYPE_FUNC_MINUTE = 0x0048,
				TYPE_FUNC_SECOND = 0x0049,
				TYPE_FUNC_NOW = 0x004A,
				TYPE_FUNC_AREAS = 0x004B,
				TYPE_FUNC_ROWS = 0x004C,
				TYPE_FUNC_COLUMNS = 0x004D,
				TYPE_FUNC_OFFSET = 0x004E,
				TYPE_FUNC_ABSREF = 0x004F,
				TYPE_FUNC_RELREF = 0x0050,
				TYPE_FUNC_ARGUMENT = 0x0051,
				TYPE_FUNC_SEARCH = 0x0052,
				TYPE_FUNC_TRANSPOSE = 0x0053,
				TYPE_FUNC_ERROR = 0x0054,
				TYPE_FUNC_STEP = 0x0055,
				TYPE_FUNC_TYPE = 0x0056,
				TYPE_FUNC_ECHO = 0x0057,
				TYPE_FUNC_SET_NAME = 0x0058,
				TYPE_FUNC_CALLER = 0x0059,
				TYPE_FUNC_DEREF = 0x005A,
				TYPE_FUNC_WINDOWS = 0x005B,
				TYPE_FUNC_SERIES = 0x005C,
				TYPE_FUNC_DOCUMENTS = 0x005D,
				TYPE_FUNC_ACTIVE_CELL = 0x005E,
				TYPE_FUNC_SELECTION = 0x005F,
				TYPE_FUNC_RESULT = 0x0060,
				TYPE_FUNC_ATAN2 = 0x0061,
				TYPE_FUNC_ASIN = 0x0062,
				TYPE_FUNC_ACOS = 0x0063,
				TYPE_FUNC_CHOOSE = 0x0064,
				TYPE_FUNC_HLOOKUP = 0x0065,
				TYPE_FUNC_VLOOKUP = 0x0066,
				TYPE_FUNC_LINKS = 0x0067,
				TYPE_FUNC_INPUT = 0x0068,
				TYPE_FUNC_ISREF = 0x0069,
				TYPE_FUNC_GET_FORMULA = 0x006A,
				TYPE_FUNC_GET_NAME = 0x006B,
				TYPE_FUNC_SET_VALUE = 0x006C,
				TYPE_FUNC_LOG = 0x006D,
				TYPE_FUNC_EXEC = 0x006E,
				TYPE_FUNC_CHAR = 0x006F,
				TYPE_FUNC_LOWER = 0x0070,
				TYPE_FUNC_UPPER = 0x0071,
				TYPE_FUNC_PROPER = 0x0072,
				TYPE_FUNC_LEFT = 0x0073,
				TYPE_FUNC_RIGHT = 0x0074,
				TYPE_FUNC_EXACT = 0x0075,
				TYPE_FUNC_TRIM = 0x0076,
				TYPE_FUNC_REPLACE = 0x0077,
				TYPE_FUNC_SUBSTITUTE = 0x0078,
				TYPE_FUNC_CODE = 0x0079,
				TYPE_FUNC_NAMES = 0x007A,
				TYPE_FUNC_DIRECTORY = 0x007B,
				TYPE_FUNC_FIND = 0x007C,
				TYPE_FUNC_CELL = 0x007D,
				TYPE_FUNC_ISERR = 0x007E,
				TYPE_FUNC_ISTEXT = 0x007F,
				TYPE_FUNC_ISNUMBER = 0x0080,
				TYPE_FUNC_ISBLANK = 0x0081,
				TYPE_FUNC_T = 0x0082,
				TYPE_FUNC_N = 0x0083,
				TYPE_FUNC_FOPEN = 0x0084,
				TYPE_FUNC_FCLOSE = 0x0085,
				TYPE_FUNC_FSIZE = 0x0086,
				TYPE_FUNC_FREADLN = 0x0087,
				TYPE_FUNC_FREAD = 0x0088,
				TYPE_FUNC_FWRITELN = 0x0089,
				TYPE_FUNC_FWRITE = 0x008A,
				TYPE_FUNC_FPOS = 0x008B,
				TYPE_FUNC_DATEVALUE = 0x008C,
				TYPE_FUNC_TIMEVALUE = 0x008D,
				TYPE_FUNC_SLN = 0x008E,
				TYPE_FUNC_SYD = 0x008F,
				TYPE_FUNC_DDB = 0x0090,
				TYPE_FUNC_GET_DEF = 0x0091,
				TYPE_FUNC_REFTEXT = 0x0092,
				TYPE_FUNC_TEXTREF = 0x0093,
				TYPE_FUNC_INDIRECT = 0x0094,
				TYPE_FUNC_REGISTER = 0x0095,
				TYPE_FUNC_CALL = 0x0096,
				TYPE_FUNC_ADD_BAR = 0x0097,
				TYPE_FUNC_ADD_MENU = 0x0098,
				TYPE_FUNC_ADD_COMMAND = 0x0099,
				TYPE_FUNC_ENABLE_COMMAND = 0x009A,
				TYPE_FUNC_CHECK_COMMAND = 0x009B,
				TYPE_FUNC_RENAME_COMMAND = 0x009C,
				TYPE_FUNC_SHOW_BAR = 0x009D,
				TYPE_FUNC_DELETE_MENU = 0x009E,
				TYPE_FUNC_DELETE_COMMAND = 0x009F,
				TYPE_FUNC_GET_CHART_ITEM = 0x00A0,
				TYPE_FUNC_DIALOG_BOX = 0x00A1,
				TYPE_FUNC_CLEAN = 0x00A2,
				TYPE_FUNC_MDETERM = 0x00A3,
				TYPE_FUNC_MINVERSE = 0x00A4,
				TYPE_FUNC_MMULT = 0x00A5,
				TYPE_FUNC_FILES = 0x00A6,
				TYPE_FUNC_IPMT = 0x00A7,
				TYPE_FUNC_PPMT = 0x00A8,
				TYPE_FUNC_COUNTA = 0x00A9,
				TYPE_FUNC_CANCEL_KEY = 0x00AA,
				TYPE_FUNC_FOR = 0x00AB,
				TYPE_FUNC_WHILE = 0x00AC,
				TYPE_FUNC_BREAK = 0x00AD,
				TYPE_FUNC_NEXT = 0x00AE,
				TYPE_FUNC_INITIATE = 0x00AF,
				TYPE_FUNC_REQUEST = 0x00B0,
				TYPE_FUNC_POKE = 0x00B1,
				TYPE_FUNC_EXECUTE = 0x00B2,
				TYPE_FUNC_TERMINATE = 0x00B3,
				TYPE_FUNC_RESTART = 0x00B4,
				TYPE_FUNC_HELP = 0x00B5,
				TYPE_FUNC_GET_BAR = 0x00B6,
				TYPE_FUNC_PRODUCT = 0x00B7,
				TYPE_FUNC_FACT = 0x00B8,
				TYPE_FUNC_GET_CELL = 0x00B9,
				TYPE_FUNC_GET_WORKSPACE = 0x00BA,
				TYPE_FUNC_GET_WINDOW = 0x00BB,
				TYPE_FUNC_GET_DOCUMENT = 0x00BC,
				TYPE_FUNC_DPRODUCT = 0x00BD,
				TYPE_FUNC_ISNONTEXT = 0x00BE,
				TYPE_FUNC_GET_NOTE = 0x00BF,
				TYPE_FUNC_NOTE = 0x00C0,
				TYPE_FUNC_STDEVP = 0x00C1,
				TYPE_FUNC_VARP = 0x00C2,
				TYPE_FUNC_DSTDEVP = 0x00C3,
				TYPE_FUNC_DVARP = 0x00C4,
				TYPE_FUNC_TRUNC = 0x00C5,
				TYPE_FUNC_ISLOGICAL = 0x00C6,
				TYPE_FUNC_DCOUNTA = 0x00C7,
				TYPE_FUNC_DELETE_BAR = 0x00C8,
				TYPE_FUNC_UNREGISTER = 0x00C9,
				TYPE_FUNC_USDOLLAR = 0x00CC,
				TYPE_FUNC_FINDB = 0x00CD,
				TYPE_FUNC_SEARCHB = 0x00CE,
				TYPE_FUNC_REPLACEB = 0x00CF,
				TYPE_FUNC_LEFTB = 0x00D0,
				TYPE_FUNC_RIGHTB = 0x00D1,
				TYPE_FUNC_MIDB = 0x00D2,
				TYPE_FUNC_LENB = 0x00D3,
				TYPE_FUNC_ROUNDUP = 0x00D4,
				TYPE_FUNC_ROUNDDOWN = 0x00D5,
				TYPE_FUNC_ASC = 0x00D6,
				TYPE_FUNC_DBCS = 0x00D7,
				TYPE_FUNC_RANK = 0x00D8,
				TYPE_FUNC_ADDRESS = 0x00DB,
				TYPE_FUNC_DAYS360 = 0x00DC,
				TYPE_FUNC_TODAY = 0x00DD,
				TYPE_FUNC_VDB = 0x00DE,
				TYPE_FUNC_ELSE = 0x00DF,
				TYPE_FUNC_ELSE_IF = 0x00E0,
				TYPE_FUNC_END_IF = 0x00E1,
				TYPE_FUNC_FOR_CELL = 0x00E2,
				TYPE_FUNC_MEDIAN = 0x00E3,
				TYPE_FUNC_SUMPRODUCT = 0x00E4,
				TYPE_FUNC_SINH = 0x00E5,
				TYPE_FUNC_COSH = 0x00E6,
				TYPE_FUNC_TANH = 0x00E7,
				TYPE_FUNC_ASINH = 0x00E8,
				TYPE_FUNC_ACOSH = 0x00E9,
				TYPE_FUNC_ATANH = 0x00EA,
				TYPE_FUNC_DGET = 0x00EB,
				TYPE_FUNC_CREATE_OBJECT = 0x00EC,
				TYPE_FUNC_VOLATILE = 0x00ED,
				TYPE_FUNC_LAST_ERROR = 0x00EE,
				TYPE_FUNC_CUSTOM_UNDO = 0x00EF,
				TYPE_FUNC_CUSTOM_REPEAT = 0x00F0,
				TYPE_FUNC_FORMULA_CONVERT = 0x00F1,
				TYPE_FUNC_GET_LINK_INFO = 0x00F2,
				TYPE_FUNC_TEXT_BOX = 0x00F3,
				TYPE_FUNC_INFO = 0x00F4,
				TYPE_FUNC_GROUP = 0x00F5,
				TYPE_FUNC_GET_OBJECT = 0x00F6,
				TYPE_FUNC_DB = 0x00F7,
				TYPE_FUNC_PAUSE = 0x00F8,
				TYPE_FUNC_RESUME = 0x00FB,
				TYPE_FUNC_FREQUENCY = 0x00FC,
				TYPE_FUNC_ADD_TOOLBAR = 0x00FD,
				TYPE_FUNC_DELETE_TOOLBAR = 0x00FE,
				TYPE_FUNC_RESET_TOOLBAR = 0x0100,
				TYPE_FUNC_EVALUATE = 0x0101,
				TYPE_FUNC_GET_TOOLBAR = 0x0102,
				TYPE_FUNC_GET_TOOL = 0x0103,
				TYPE_FUNC_SPELLING_CHECK = 0x0104,
				TYPE_FUNC_ERROR_TYPE = 0x0105,
				TYPE_FUNC_APP_TITLE = 0x0106,
				TYPE_FUNC_WINDOW_TITLE = 0x0107,
				TYPE_FUNC_SAVE_TOOLBAR = 0x0108,
				TYPE_FUNC_ENABLE_TOOL = 0x0109,
				TYPE_FUNC_PRESS_TOOL = 0x010A,
				TYPE_FUNC_REGISTER_ID = 0x010B,
				TYPE_FUNC_GET_WORKBOOK = 0x010C,
				TYPE_FUNC_AVEDEV = 0x010D,
				TYPE_FUNC_BETADIST = 0x010E,
				TYPE_FUNC_GAMMALN = 0x010F,
				TYPE_FUNC_BETAINV = 0x0110,
				TYPE_FUNC_BINOMDIST = 0x0111,
				TYPE_FUNC_CHIDIST = 0x0112,
				TYPE_FUNC_CHIINV = 0x0113,
				TYPE_FUNC_COMBIN = 0x0114,
				TYPE_FUNC_CONFIDENCE = 0x0115,
				TYPE_FUNC_CRITBINOM = 0x0116,
				TYPE_FUNC_EVEN = 0x0117,
				TYPE_FUNC_EXPONDIST = 0x0118,
				TYPE_FUNC_FDIST = 0x0119,
				TYPE_FUNC_FINV = 0x011A,
				TYPE_FUNC_FISHER = 0x011B,
				TYPE_FUNC_FISHERINV = 0x011C,
				TYPE_FUNC_FLOOR = 0x011D,
				TYPE_FUNC_GAMMADIST = 0x011E,
				TYPE_FUNC_GAMMAINV = 0x011F,
				TYPE_FUNC_CEILING = 0x0120,
				TYPE_FUNC_HYPGEOMDIST = 0x0121,
				TYPE_FUNC_LOGNORMDIST = 0x0122,
				TYPE_FUNC_LOGINV = 0x0123,
				TYPE_FUNC_NEGBINOMDIST = 0x0124,
				TYPE_FUNC_NORMDIST = 0x0125,
				TYPE_FUNC_NORMSDIST = 0x0126,
				TYPE_FUNC_NORMINV = 0x0127,
				TYPE_FUNC_NORMSINV = 0x0128,
				TYPE_FUNC_STANDARDIZE = 0x0129,
				TYPE_FUNC_ODD = 0x012A,
				TYPE_FUNC_PERMUT = 0x012B,
				TYPE_FUNC_POISSON = 0x012C,
				TYPE_FUNC_TDIST = 0x012D,
				TYPE_FUNC_WEIBULL = 0x012E,
				TYPE_FUNC_SUMXMY2 = 0x012F,
				TYPE_FUNC_SUMX2MY2 = 0x0130,
				TYPE_FUNC_SUMX2PY2 = 0x0131,
				TYPE_FUNC_CHITEST = 0x0132,
				TYPE_FUNC_CORREL = 0x0133,
				TYPE_FUNC_COVAR = 0x0134,
				TYPE_FUNC_FORECAST = 0x0135,
				TYPE_FUNC_FTEST = 0x0136,
				TYPE_FUNC_INTERCEPT = 0x0137,
				TYPE_FUNC_PEARSON = 0x0138,
				TYPE_FUNC_RSQ = 0x0139,
				TYPE_FUNC_STEYX = 0x013A,
				TYPE_FUNC_SLOPE = 0x013B,
				TYPE_FUNC_TTEST = 0x013C,
				TYPE_FUNC_PROB = 0x013D,
				TYPE_FUNC_DEVSQ = 0x013E,
				TYPE_FUNC_GEOMEAN = 0x013F,
				TYPE_FUNC_HARMEAN = 0x0140,
				TYPE_FUNC_SUMSQ = 0x0141,
				TYPE_FUNC_KURT = 0x0142,
				TYPE_FUNC_SKEW = 0x0143,
				TYPE_FUNC_ZTEST = 0x0144,
				TYPE_FUNC_LARGE = 0x0145,
				TYPE_FUNC_SMALL = 0x0146,
				TYPE_FUNC_QUARTILE = 0x0147,
				TYPE_FUNC_PERCENTILE = 0x0148,
				TYPE_FUNC_PERCENTRANK = 0x0149,
				TYPE_FUNC_MODE = 0x014A,
				TYPE_FUNC_TRIMMEAN = 0x014B,
				TYPE_FUNC_TINV = 0x014C,
				TYPE_FUNC_MOVIE_COMMAND = 0x014E,
				TYPE_FUNC_GET_MOVIE = 0x014F,
				TYPE_FUNC_CONCATENATE = 0x0150,
				TYPE_FUNC_POWER = 0x0151,
				TYPE_FUNC_PIVOT_ADD_DATA = 0x0152,
				TYPE_FUNC_GET_PIVOT_TABLE = 0x0153,
				TYPE_FUNC_GET_PIVOT_FIELD = 0x0154,
				TYPE_FUNC_GET_PIVOT_ITEM = 0x0155,
				TYPE_FUNC_RADIANS = 0x0156,
				TYPE_FUNC_DEGREES = 0x0157,
				TYPE_FUNC_SUBTOTAL = 0x0158,
				TYPE_FUNC_SUMIF = 0x0159,
				TYPE_FUNC_COUNTIF = 0x015A,
				TYPE_FUNC_COUNTBLANK = 0x015B,
				TYPE_FUNC_SCENARIO_GET = 0x015C,
				TYPE_FUNC_OPTIONS_LISTS_GET = 0x015D,
				TYPE_FUNC_ISPMT = 0x015E,
				TYPE_FUNC_DATEDIF = 0x015F,
				TYPE_FUNC_DATESTRING = 0x0160,
				TYPE_FUNC_NUMBERSTRING = 0x0161,
				TYPE_FUNC_ROMAN = 0x0162,
				TYPE_FUNC_OPEN_DIALOG = 0x0163,
				TYPE_FUNC_SAVE_DIALOG = 0x0164,
				TYPE_FUNC_VIEW_GET = 0x0165,
				TYPE_FUNC_GETPIVOTDATA = 0x0166,
				TYPE_FUNC_HYPERLINK = 0x0167,
				TYPE_FUNC_PHONETIC = 0x0168,
				TYPE_FUNC_AVERAGEA = 0x0169,
				TYPE_FUNC_MAXA = 0x016A,
				TYPE_FUNC_MINA = 0x016B,
				TYPE_FUNC_STDEVPA = 0x016C,
				TYPE_FUNC_VARPA = 0x016D,
				TYPE_FUNC_STDEVA = 0x016E,
				TYPE_FUNC_VARA = 0x016F,
				TYPE_FUNC_BAHTTEXT = 0x0170,
				TYPE_FUNC_THAIDAYOFWEEK = 0x0171,
				TYPE_FUNC_THAIDIGIT = 0x0172,
				TYPE_FUNC_THAIMONTHOFYEAR = 0x0173,
				TYPE_FUNC_THAINUMSOUND = 0x0174,
				TYPE_FUNC_THAINUMSTRING = 0x0175,
				TYPE_FUNC_THAISTRINGLENGTH = 0x0176,
				TYPE_FUNC_ISTHAIDIGIT = 0x0177,
				TYPE_FUNC_ROUNDBAHTDOWN = 0x0178,
				TYPE_FUNC_ROUNDBAHTUP = 0x0179,
				TYPE_FUNC_THAIYEAR = 0x017A,
				TYPE_FUNC_RTD = 0x017B,
				TYPE_BOOL,
				TYPE_INT,
				TYPE_NUMBER,
				TYPE_SPACE,
				TYPE_MISS_ARG,
				TYPE_STRING,
				TYPE_COORDINATE,
				TYPE_COORDINATE_3D,
				TYPE_AREA,
				TYPE_AREA_3D,
				TYPE_OPERATOR,
				TYPE_PAREN,
			}

			public enum SubType
			{
				SUB_TYPE_FUNCTION,
				SUB_TYPE_OPERATOR,
				SUB_TYPE_VARIABLE,
			}

			protected Type m_eType;
			protected SubType m_eSubType;
			protected byte m_nParameterCount;
			public Token(Type eType, SubType eSubType, byte nParameterCount)
			{
				m_eType = eType;
				m_eSubType = eSubType;
				m_nParameterCount = nParameterCount;
			}

			~Token()
			{
			}

			public new Token.Type GetType()
			{
				return m_eType;
			}

			public Token.SubType GetSubType()
			{
				return m_eSubType;
			}

			public byte GetParameterCount()
			{
				return m_nParameterCount;
			}

			public virtual void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendString(GetTypeName(m_eType));
			}

			public virtual ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				if (m_eType == Type.TYPE_PAREN)
					return new PtgParenRecord();
				return new PtgFuncVarRecord((ushort)(m_eType), m_nParameterCount);
			}

			public virtual bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				switch (m_eType)
				{
					case Type.TYPE_PAREN:
					{
						return true;
					}

					case Type.TYPE_FUNC_CONCATENATE:
					{
						InternalString sTemp = new InternalString("");
						if (ppValueVector.GetSize() >= m_nParameterCount)
						{
							for (ushort i = 0; i < m_nParameterCount; i++)
							{
								Value pValue = ppValueVector.Get(ppValueVector.GetSize() - 1);
								if (pValue.GetType() == Value.Type.TYPE_STRING)
								{
									sTemp.PrependString(pValue.GetString());
								}
								else
								{
									{
										sTemp = null;
									}
									{
										return false;
									}
								}
								ppValueVector.PopBack();
								{
									pValue = null;
								}
							}
							ppValueVector.PushBack(Secret.ValueImplementation.CreateStringValue(sTemp.GetExternalString()));
							{
								sTemp = null;
							}
							{
								return true;
							}
						}
						{
							sTemp = null;
						}
						{
							return false;
						}
					}

					case Type.TYPE_FUNC_HOUR:
					{
						if (ppValueVector.GetSize() >= 1)
						{
							Value pValue = ppValueVector.Get(ppValueVector.GetSize() - 1);
							if (pValue.GetType() == Value.Type.TYPE_FLOAT)
							{
								Value pOwnedValue = ppValueVector.PopBack();
								uint nInteger = (uint)(pValue.GetFloat());
								double fInteger = (double)(nInteger);
								double fFraction = pValue.GetFloat() - fInteger;
								ppValueVector.PushBack(Secret.ValueImplementation.CreateFloatValue((double)((int)(fFraction * 24))));
								{
									return true;
								}
							}
						}
						return false;
					}

					case Type.TYPE_FUNC_IF:
					{
						if (ppValueVector.GetSize() >= 3)
						{
							Value pA = ppValueVector.Get(ppValueVector.GetSize() - 3);
							Value pB = ppValueVector.Get(ppValueVector.GetSize() - 2);
							Value pC = ppValueVector.Get(ppValueVector.GetSize() - 1);
							if (pA.GetType() == Value.Type.TYPE_BOOLEAN)
							{
								Value pOwnedC = ppValueVector.PopBack();
								Value pOwnedB = ppValueVector.PopBack();
								Value pOwnedA = ppValueVector.PopBack();
								if (pA.GetBoolean())
								{
									{
										NumberDuck.Value __1092584474 = pOwnedB;
										pOwnedB = null;
										ppValueVector.PushBack(__1092584474);
									}
								}
								else
								{
									{
										NumberDuck.Value __3995042842 = pOwnedC;
										pOwnedC = null;
										ppValueVector.PushBack(__3995042842);
									}
								}
								{
									return true;
								}
							}
						}
						break;
					}

					case Type.TYPE_FUNC_PI:
					{
						ppValueVector.PushBack(Secret.ValueImplementation.CreateFloatValue(3.14159265358979));
						return true;
					}

				}
				return false;
			}

			public virtual void InsertColumn(ushort nWorksheet, ushort nColumn)
			{
			}

			public virtual void DeleteColumn(ushort nWorksheet, ushort nColumn)
			{
			}

			public virtual void InsertRow(ushort nWorksheet, ushort nRow)
			{
			}

			public virtual void DeleteRow(ushort nWorksheet, ushort nRow)
			{
			}

			public static string GetTypeName(Type eType)
			{
				switch (eType)
				{
					case Type.TYPE_FUNC_COUNT:
					{
						return "COUNT";
					}

					case Type.TYPE_FUNC_IF:
					{
						return "IF";
					}

					case Type.TYPE_FUNC_ISNA:
					{
						return "ISNA";
					}

					case Type.TYPE_FUNC_ISERROR:
					{
						return "ISERROR";
					}

					case Type.TYPE_FUNC_SUM:
					{
						return "SUM";
					}

					case Type.TYPE_FUNC_AVERAGE:
					{
						return "AVERAGE";
					}

					case Type.TYPE_FUNC_MIN:
					{
						return "MIN";
					}

					case Type.TYPE_FUNC_MAX:
					{
						return "MAX";
					}

					case Type.TYPE_FUNC_ROW:
					{
						return "ROW";
					}

					case Type.TYPE_FUNC_COLUMN:
					{
						return "COLUMN";
					}

					case Type.TYPE_FUNC_NA:
					{
						return "NA";
					}

					case Type.TYPE_FUNC_NPV:
					{
						return "NPV";
					}

					case Type.TYPE_FUNC_STDEV:
					{
						return "STDEV";
					}

					case Type.TYPE_FUNC_DOLLAR:
					{
						return "DOLLAR";
					}

					case Type.TYPE_FUNC_FIXED:
					{
						return "FIXED";
					}

					case Type.TYPE_FUNC_SIN:
					{
						return "SIN";
					}

					case Type.TYPE_FUNC_COS:
					{
						return "COS";
					}

					case Type.TYPE_FUNC_TAN:
					{
						return "TAN";
					}

					case Type.TYPE_FUNC_ATAN:
					{
						return "ATAN";
					}

					case Type.TYPE_FUNC_PI:
					{
						return "PI";
					}

					case Type.TYPE_FUNC_SQRT:
					{
						return "SQRT";
					}

					case Type.TYPE_FUNC_EXP:
					{
						return "EXP";
					}

					case Type.TYPE_FUNC_LN:
					{
						return "LN";
					}

					case Type.TYPE_FUNC_LOG10:
					{
						return "LOG10";
					}

					case Type.TYPE_FUNC_ABS:
					{
						return "ABS";
					}

					case Type.TYPE_FUNC_INT:
					{
						return "INT";
					}

					case Type.TYPE_FUNC_SIGN:
					{
						return "SIGN";
					}

					case Type.TYPE_FUNC_ROUND:
					{
						return "ROUND";
					}

					case Type.TYPE_FUNC_LOOKUP:
					{
						return "LOOKUP";
					}

					case Type.TYPE_FUNC_INDEX:
					{
						return "INDEX";
					}

					case Type.TYPE_FUNC_REPT:
					{
						return "REPT";
					}

					case Type.TYPE_FUNC_MID:
					{
						return "MID";
					}

					case Type.TYPE_FUNC_LEN:
					{
						return "LEN";
					}

					case Type.TYPE_FUNC_VALUE:
					{
						return "VALUE";
					}

					case Type.TYPE_FUNC_TRUE:
					{
						return "TRUE";
					}

					case Type.TYPE_FUNC_FALSE:
					{
						return "FALSE";
					}

					case Type.TYPE_FUNC_AND:
					{
						return "AND";
					}

					case Type.TYPE_FUNC_OR:
					{
						return "OR";
					}

					case Type.TYPE_FUNC_NOT:
					{
						return "NOT";
					}

					case Type.TYPE_FUNC_MOD:
					{
						return "MOD";
					}

					case Type.TYPE_FUNC_DCOUNT:
					{
						return "DCOUNT";
					}

					case Type.TYPE_FUNC_DSUM:
					{
						return "DSUM";
					}

					case Type.TYPE_FUNC_DAVERAGE:
					{
						return "DAVERAGE";
					}

					case Type.TYPE_FUNC_DMIN:
					{
						return "DMIN";
					}

					case Type.TYPE_FUNC_DMAX:
					{
						return "DMAX";
					}

					case Type.TYPE_FUNC_DSTDEV:
					{
						return "DSTDEV";
					}

					case Type.TYPE_FUNC_VAR:
					{
						return "VAR";
					}

					case Type.TYPE_FUNC_DVAR:
					{
						return "DVAR";
					}

					case Type.TYPE_FUNC_TEXT:
					{
						return "TEXT";
					}

					case Type.TYPE_FUNC_LINEST:
					{
						return "LINEST";
					}

					case Type.TYPE_FUNC_TREND:
					{
						return "TREND";
					}

					case Type.TYPE_FUNC_LOGEST:
					{
						return "LOGEST";
					}

					case Type.TYPE_FUNC_GROWTH:
					{
						return "GROWTH";
					}

					case Type.TYPE_FUNC_GOTO:
					{
						return "GOTO";
					}

					case Type.TYPE_FUNC_HALT:
					{
						return "HALT";
					}

					case Type.TYPE_FUNC_RETURN:
					{
						return "RETURN";
					}

					case Type.TYPE_FUNC_PV:
					{
						return "PV";
					}

					case Type.TYPE_FUNC_FV:
					{
						return "FV";
					}

					case Type.TYPE_FUNC_NPER:
					{
						return "NPER";
					}

					case Type.TYPE_FUNC_PMT:
					{
						return "PMT";
					}

					case Type.TYPE_FUNC_RATE:
					{
						return "RATE";
					}

					case Type.TYPE_FUNC_MIRR:
					{
						return "MIRR";
					}

					case Type.TYPE_FUNC_IRR:
					{
						return "IRR";
					}

					case Type.TYPE_FUNC_RAND:
					{
						return "RAND";
					}

					case Type.TYPE_FUNC_MATCH:
					{
						return "MATCH";
					}

					case Type.TYPE_FUNC_DATE:
					{
						return "DATE";
					}

					case Type.TYPE_FUNC_TIME:
					{
						return "TIME";
					}

					case Type.TYPE_FUNC_DAY:
					{
						return "DAY";
					}

					case Type.TYPE_FUNC_MONTH:
					{
						return "MONTH";
					}

					case Type.TYPE_FUNC_YEAR:
					{
						return "YEAR";
					}

					case Type.TYPE_FUNC_WEEKDAY:
					{
						return "WEEKDAY";
					}

					case Type.TYPE_FUNC_HOUR:
					{
						return "HOUR";
					}

					case Type.TYPE_FUNC_MINUTE:
					{
						return "MINUTE";
					}

					case Type.TYPE_FUNC_SECOND:
					{
						return "SECOND";
					}

					case Type.TYPE_FUNC_NOW:
					{
						return "NOW";
					}

					case Type.TYPE_FUNC_AREAS:
					{
						return "AREAS";
					}

					case Type.TYPE_FUNC_ROWS:
					{
						return "ROWS";
					}

					case Type.TYPE_FUNC_COLUMNS:
					{
						return "COLUMNS";
					}

					case Type.TYPE_FUNC_OFFSET:
					{
						return "OFFSET";
					}

					case Type.TYPE_FUNC_ABSREF:
					{
						return "ABSREF";
					}

					case Type.TYPE_FUNC_RELREF:
					{
						return "RELREF";
					}

					case Type.TYPE_FUNC_ARGUMENT:
					{
						return "ARGUMENT";
					}

					case Type.TYPE_FUNC_SEARCH:
					{
						return "SEARCH";
					}

					case Type.TYPE_FUNC_TRANSPOSE:
					{
						return "TRANSPOSE";
					}

					case Type.TYPE_FUNC_ERROR:
					{
						return "ERROR";
					}

					case Type.TYPE_FUNC_STEP:
					{
						return "STEP";
					}

					case Type.TYPE_FUNC_TYPE:
					{
						return "TYPE";
					}

					case Type.TYPE_FUNC_ECHO:
					{
						return "ECHO";
					}

					case Type.TYPE_FUNC_SET_NAME:
					{
						return "SET.NAME";
					}

					case Type.TYPE_FUNC_CALLER:
					{
						return "CALLER";
					}

					case Type.TYPE_FUNC_DEREF:
					{
						return "DEREF";
					}

					case Type.TYPE_FUNC_WINDOWS:
					{
						return "WINDOWS";
					}

					case Type.TYPE_FUNC_SERIES:
					{
						return "SERIES";
					}

					case Type.TYPE_FUNC_DOCUMENTS:
					{
						return "DOCUMENTS";
					}

					case Type.TYPE_FUNC_ACTIVE_CELL:
					{
						return "ACTIVE.CELL";
					}

					case Type.TYPE_FUNC_SELECTION:
					{
						return "SELECTION";
					}

					case Type.TYPE_FUNC_RESULT:
					{
						return "RESULT";
					}

					case Type.TYPE_FUNC_ATAN2:
					{
						return "ATAN2";
					}

					case Type.TYPE_FUNC_ASIN:
					{
						return "ASIN";
					}

					case Type.TYPE_FUNC_ACOS:
					{
						return "ACOS";
					}

					case Type.TYPE_FUNC_CHOOSE:
					{
						return "CHOOSE";
					}

					case Type.TYPE_FUNC_HLOOKUP:
					{
						return "HLOOKUP";
					}

					case Type.TYPE_FUNC_VLOOKUP:
					{
						return "VLOOKUP";
					}

					case Type.TYPE_FUNC_LINKS:
					{
						return "LINKS";
					}

					case Type.TYPE_FUNC_INPUT:
					{
						return "INPUT";
					}

					case Type.TYPE_FUNC_ISREF:
					{
						return "ISREF";
					}

					case Type.TYPE_FUNC_GET_FORMULA:
					{
						return "GET.FORMULA";
					}

					case Type.TYPE_FUNC_GET_NAME:
					{
						return "GET.NAME";
					}

					case Type.TYPE_FUNC_SET_VALUE:
					{
						return "SET.VALUE";
					}

					case Type.TYPE_FUNC_LOG:
					{
						return "LOG";
					}

					case Type.TYPE_FUNC_EXEC:
					{
						return "EXEC";
					}

					case Type.TYPE_FUNC_CHAR:
					{
						return "CHAR";
					}

					case Type.TYPE_FUNC_LOWER:
					{
						return "LOWER";
					}

					case Type.TYPE_FUNC_UPPER:
					{
						return "UPPER";
					}

					case Type.TYPE_FUNC_PROPER:
					{
						return "PROPER";
					}

					case Type.TYPE_FUNC_LEFT:
					{
						return "LEFT";
					}

					case Type.TYPE_FUNC_RIGHT:
					{
						return "RIGHT";
					}

					case Type.TYPE_FUNC_EXACT:
					{
						return "EXACT";
					}

					case Type.TYPE_FUNC_TRIM:
					{
						return "TRIM";
					}

					case Type.TYPE_FUNC_REPLACE:
					{
						return "REPLACE";
					}

					case Type.TYPE_FUNC_SUBSTITUTE:
					{
						return "SUBSTITUTE";
					}

					case Type.TYPE_FUNC_CODE:
					{
						return "CODE";
					}

					case Type.TYPE_FUNC_NAMES:
					{
						return "NAMES";
					}

					case Type.TYPE_FUNC_DIRECTORY:
					{
						return "DIRECTORY";
					}

					case Type.TYPE_FUNC_FIND:
					{
						return "FIND";
					}

					case Type.TYPE_FUNC_CELL:
					{
						return "CELL";
					}

					case Type.TYPE_FUNC_ISERR:
					{
						return "ISERR";
					}

					case Type.TYPE_FUNC_ISTEXT:
					{
						return "ISTEXT";
					}

					case Type.TYPE_FUNC_ISNUMBER:
					{
						return "ISNUMBER";
					}

					case Type.TYPE_FUNC_ISBLANK:
					{
						return "ISBLANK";
					}

					case Type.TYPE_FUNC_T:
					{
						return "T";
					}

					case Type.TYPE_FUNC_N:
					{
						return "N";
					}

					case Type.TYPE_FUNC_FOPEN:
					{
						return "FOPEN";
					}

					case Type.TYPE_FUNC_FCLOSE:
					{
						return "FCLOSE";
					}

					case Type.TYPE_FUNC_FSIZE:
					{
						return "FSIZE";
					}

					case Type.TYPE_FUNC_FREADLN:
					{
						return "FREADLN";
					}

					case Type.TYPE_FUNC_FREAD:
					{
						return "FREAD";
					}

					case Type.TYPE_FUNC_FWRITELN:
					{
						return "FWRITELN";
					}

					case Type.TYPE_FUNC_FWRITE:
					{
						return "FWRITE";
					}

					case Type.TYPE_FUNC_FPOS:
					{
						return "FPOS";
					}

					case Type.TYPE_FUNC_DATEVALUE:
					{
						return "DATEVALUE";
					}

					case Type.TYPE_FUNC_TIMEVALUE:
					{
						return "TIMEVALUE";
					}

					case Type.TYPE_FUNC_SLN:
					{
						return "SLN";
					}

					case Type.TYPE_FUNC_SYD:
					{
						return "SYD";
					}

					case Type.TYPE_FUNC_DDB:
					{
						return "DDB";
					}

					case Type.TYPE_FUNC_GET_DEF:
					{
						return "GET.DEF";
					}

					case Type.TYPE_FUNC_REFTEXT:
					{
						return "REFTEXT";
					}

					case Type.TYPE_FUNC_TEXTREF:
					{
						return "TEXTREF";
					}

					case Type.TYPE_FUNC_INDIRECT:
					{
						return "INDIRECT";
					}

					case Type.TYPE_FUNC_REGISTER:
					{
						return "REGISTER";
					}

					case Type.TYPE_FUNC_CALL:
					{
						return "CALL";
					}

					case Type.TYPE_FUNC_ADD_BAR:
					{
						return "ADD.BAR";
					}

					case Type.TYPE_FUNC_ADD_MENU:
					{
						return "ADD.MENU";
					}

					case Type.TYPE_FUNC_ADD_COMMAND:
					{
						return "ADD.COMMAND";
					}

					case Type.TYPE_FUNC_ENABLE_COMMAND:
					{
						return "ENABLE.COMMAND";
					}

					case Type.TYPE_FUNC_CHECK_COMMAND:
					{
						return "CHECK.COMMAND";
					}

					case Type.TYPE_FUNC_RENAME_COMMAND:
					{
						return "RENAME.COMMAND";
					}

					case Type.TYPE_FUNC_SHOW_BAR:
					{
						return "SHOW.BAR";
					}

					case Type.TYPE_FUNC_DELETE_MENU:
					{
						return "DELETE.MENU";
					}

					case Type.TYPE_FUNC_DELETE_COMMAND:
					{
						return "DELETE.COMMAND";
					}

					case Type.TYPE_FUNC_GET_CHART_ITEM:
					{
						return "GET.CHART.ITEM";
					}

					case Type.TYPE_FUNC_DIALOG_BOX:
					{
						return "DIALOG.BOX";
					}

					case Type.TYPE_FUNC_CLEAN:
					{
						return "CLEAN";
					}

					case Type.TYPE_FUNC_MDETERM:
					{
						return "MDETERM";
					}

					case Type.TYPE_FUNC_MINVERSE:
					{
						return "MINVERSE";
					}

					case Type.TYPE_FUNC_MMULT:
					{
						return "MMULT";
					}

					case Type.TYPE_FUNC_FILES:
					{
						return "FILES";
					}

					case Type.TYPE_FUNC_IPMT:
					{
						return "IPMT";
					}

					case Type.TYPE_FUNC_PPMT:
					{
						return "PPMT";
					}

					case Type.TYPE_FUNC_COUNTA:
					{
						return "COUNTA";
					}

					case Type.TYPE_FUNC_CANCEL_KEY:
					{
						return "CANCEL.KEY";
					}

					case Type.TYPE_FUNC_FOR:
					{
						return "FOR";
					}

					case Type.TYPE_FUNC_WHILE:
					{
						return "WHILE";
					}

					case Type.TYPE_FUNC_BREAK:
					{
						return "BREAK";
					}

					case Type.TYPE_FUNC_NEXT:
					{
						return "NEXT";
					}

					case Type.TYPE_FUNC_INITIATE:
					{
						return "INITIATE";
					}

					case Type.TYPE_FUNC_REQUEST:
					{
						return "REQUEST";
					}

					case Type.TYPE_FUNC_POKE:
					{
						return "POKE";
					}

					case Type.TYPE_FUNC_EXECUTE:
					{
						return "EXECUTE";
					}

					case Type.TYPE_FUNC_TERMINATE:
					{
						return "TERMINATE";
					}

					case Type.TYPE_FUNC_RESTART:
					{
						return "RESTART";
					}

					case Type.TYPE_FUNC_HELP:
					{
						return "HELP";
					}

					case Type.TYPE_FUNC_GET_BAR:
					{
						return "GET.BAR";
					}

					case Type.TYPE_FUNC_PRODUCT:
					{
						return "PRODUCT";
					}

					case Type.TYPE_FUNC_FACT:
					{
						return "FACT";
					}

					case Type.TYPE_FUNC_GET_CELL:
					{
						return "GET.CELL";
					}

					case Type.TYPE_FUNC_GET_WORKSPACE:
					{
						return "GET.WORKSPACE";
					}

					case Type.TYPE_FUNC_GET_WINDOW:
					{
						return "GET.WINDOW";
					}

					case Type.TYPE_FUNC_GET_DOCUMENT:
					{
						return "GET.DOCUMENT";
					}

					case Type.TYPE_FUNC_DPRODUCT:
					{
						return "DPRODUCT";
					}

					case Type.TYPE_FUNC_ISNONTEXT:
					{
						return "ISNONTEXT";
					}

					case Type.TYPE_FUNC_GET_NOTE:
					{
						return "GET.NOTE";
					}

					case Type.TYPE_FUNC_NOTE:
					{
						return "NOTE";
					}

					case Type.TYPE_FUNC_STDEVP:
					{
						return "STDEVP";
					}

					case Type.TYPE_FUNC_VARP:
					{
						return "VARP";
					}

					case Type.TYPE_FUNC_DSTDEVP:
					{
						return "DSTDEVP";
					}

					case Type.TYPE_FUNC_DVARP:
					{
						return "DVARP";
					}

					case Type.TYPE_FUNC_TRUNC:
					{
						return "TRUNC";
					}

					case Type.TYPE_FUNC_ISLOGICAL:
					{
						return "ISLOGICAL";
					}

					case Type.TYPE_FUNC_DCOUNTA:
					{
						return "DCOUNTA";
					}

					case Type.TYPE_FUNC_DELETE_BAR:
					{
						return "DELETE.BAR";
					}

					case Type.TYPE_FUNC_UNREGISTER:
					{
						return "UNREGISTER";
					}

					case Type.TYPE_FUNC_USDOLLAR:
					{
						return "USDOLLAR";
					}

					case Type.TYPE_FUNC_FINDB:
					{
						return "FINDB";
					}

					case Type.TYPE_FUNC_SEARCHB:
					{
						return "SEARCHB";
					}

					case Type.TYPE_FUNC_REPLACEB:
					{
						return "REPLACEB";
					}

					case Type.TYPE_FUNC_LEFTB:
					{
						return "LEFTB";
					}

					case Type.TYPE_FUNC_RIGHTB:
					{
						return "RIGHTB";
					}

					case Type.TYPE_FUNC_MIDB:
					{
						return "MIDB";
					}

					case Type.TYPE_FUNC_LENB:
					{
						return "LENB";
					}

					case Type.TYPE_FUNC_ROUNDUP:
					{
						return "ROUNDUP";
					}

					case Type.TYPE_FUNC_ROUNDDOWN:
					{
						return "ROUNDDOWN";
					}

					case Type.TYPE_FUNC_ASC:
					{
						return "ASC";
					}

					case Type.TYPE_FUNC_DBCS:
					{
						return "DBCS";
					}

					case Type.TYPE_FUNC_RANK:
					{
						return "RANK";
					}

					case Type.TYPE_FUNC_ADDRESS:
					{
						return "ADDRESS";
					}

					case Type.TYPE_FUNC_DAYS360:
					{
						return "DAYS360";
					}

					case Type.TYPE_FUNC_TODAY:
					{
						return "TODAY";
					}

					case Type.TYPE_FUNC_VDB:
					{
						return "VDB";
					}

					case Type.TYPE_FUNC_ELSE:
					{
						return "ELSE";
					}

					case Type.TYPE_FUNC_ELSE_IF:
					{
						return "ELSE.IF";
					}

					case Type.TYPE_FUNC_END_IF:
					{
						return "END.IF";
					}

					case Type.TYPE_FUNC_FOR_CELL:
					{
						return "FOR.CELL";
					}

					case Type.TYPE_FUNC_MEDIAN:
					{
						return "MEDIAN";
					}

					case Type.TYPE_FUNC_SUMPRODUCT:
					{
						return "SUMPRODUCT";
					}

					case Type.TYPE_FUNC_SINH:
					{
						return "SINH";
					}

					case Type.TYPE_FUNC_COSH:
					{
						return "COSH";
					}

					case Type.TYPE_FUNC_TANH:
					{
						return "TANH";
					}

					case Type.TYPE_FUNC_ASINH:
					{
						return "ASINH";
					}

					case Type.TYPE_FUNC_ACOSH:
					{
						return "ACOSH";
					}

					case Type.TYPE_FUNC_ATANH:
					{
						return "ATANH";
					}

					case Type.TYPE_FUNC_DGET:
					{
						return "DGET";
					}

					case Type.TYPE_FUNC_CREATE_OBJECT:
					{
						return "CREATE.OBJECT";
					}

					case Type.TYPE_FUNC_VOLATILE:
					{
						return "VOLATILE";
					}

					case Type.TYPE_FUNC_LAST_ERROR:
					{
						return "LAST.ERROR";
					}

					case Type.TYPE_FUNC_CUSTOM_UNDO:
					{
						return "CUSTOM.UNDO";
					}

					case Type.TYPE_FUNC_CUSTOM_REPEAT:
					{
						return "CUSTOM.REPEAT";
					}

					case Type.TYPE_FUNC_FORMULA_CONVERT:
					{
						return "FORMULA.CONVERT";
					}

					case Type.TYPE_FUNC_GET_LINK_INFO:
					{
						return "GET.LINK.INFO";
					}

					case Type.TYPE_FUNC_TEXT_BOX:
					{
						return "TEXT.BOX";
					}

					case Type.TYPE_FUNC_INFO:
					{
						return "INFO";
					}

					case Type.TYPE_FUNC_GROUP:
					{
						return "GROUP";
					}

					case Type.TYPE_FUNC_GET_OBJECT:
					{
						return "GET.OBJECT";
					}

					case Type.TYPE_FUNC_DB:
					{
						return "DB";
					}

					case Type.TYPE_FUNC_PAUSE:
					{
						return "PAUSE";
					}

					case Type.TYPE_FUNC_RESUME:
					{
						return "RESUME";
					}

					case Type.TYPE_FUNC_FREQUENCY:
					{
						return "FREQUENCY";
					}

					case Type.TYPE_FUNC_ADD_TOOLBAR:
					{
						return "ADD.TOOLBAR";
					}

					case Type.TYPE_FUNC_DELETE_TOOLBAR:
					{
						return "DELETE.TOOLBAR";
					}

					case Type.TYPE_FUNC_RESET_TOOLBAR:
					{
						return "RESET.TOOLBAR";
					}

					case Type.TYPE_FUNC_EVALUATE:
					{
						return "EVALUATE";
					}

					case Type.TYPE_FUNC_GET_TOOLBAR:
					{
						return "GET.TOOLBAR";
					}

					case Type.TYPE_FUNC_GET_TOOL:
					{
						return "GET.TOOL";
					}

					case Type.TYPE_FUNC_SPELLING_CHECK:
					{
						return "SPELLING.CHECK";
					}

					case Type.TYPE_FUNC_ERROR_TYPE:
					{
						return "ERROR.TYPE";
					}

					case Type.TYPE_FUNC_APP_TITLE:
					{
						return "APP.TITLE";
					}

					case Type.TYPE_FUNC_WINDOW_TITLE:
					{
						return "WINDOW.TITLE";
					}

					case Type.TYPE_FUNC_SAVE_TOOLBAR:
					{
						return "SAVE.TOOLBAR";
					}

					case Type.TYPE_FUNC_ENABLE_TOOL:
					{
						return "ENABLE.TOOL";
					}

					case Type.TYPE_FUNC_PRESS_TOOL:
					{
						return "PRESS.TOOL";
					}

					case Type.TYPE_FUNC_REGISTER_ID:
					{
						return "REGISTER.ID";
					}

					case Type.TYPE_FUNC_GET_WORKBOOK:
					{
						return "GET.WORKBOOK";
					}

					case Type.TYPE_FUNC_AVEDEV:
					{
						return "AVEDEV";
					}

					case Type.TYPE_FUNC_BETADIST:
					{
						return "BETADIST";
					}

					case Type.TYPE_FUNC_GAMMALN:
					{
						return "GAMMALN";
					}

					case Type.TYPE_FUNC_BETAINV:
					{
						return "BETAINV";
					}

					case Type.TYPE_FUNC_BINOMDIST:
					{
						return "BINOMDIST";
					}

					case Type.TYPE_FUNC_CHIDIST:
					{
						return "CHIDIST";
					}

					case Type.TYPE_FUNC_CHIINV:
					{
						return "CHIINV";
					}

					case Type.TYPE_FUNC_COMBIN:
					{
						return "COMBIN";
					}

					case Type.TYPE_FUNC_CONFIDENCE:
					{
						return "CONFIDENCE";
					}

					case Type.TYPE_FUNC_CRITBINOM:
					{
						return "CRITBINOM";
					}

					case Type.TYPE_FUNC_EVEN:
					{
						return "EVEN";
					}

					case Type.TYPE_FUNC_EXPONDIST:
					{
						return "EXPONDIST";
					}

					case Type.TYPE_FUNC_FDIST:
					{
						return "FDIST";
					}

					case Type.TYPE_FUNC_FINV:
					{
						return "FINV";
					}

					case Type.TYPE_FUNC_FISHER:
					{
						return "FISHER";
					}

					case Type.TYPE_FUNC_FISHERINV:
					{
						return "FISHERINV";
					}

					case Type.TYPE_FUNC_FLOOR:
					{
						return "FLOOR";
					}

					case Type.TYPE_FUNC_GAMMADIST:
					{
						return "GAMMADIST";
					}

					case Type.TYPE_FUNC_GAMMAINV:
					{
						return "GAMMAINV";
					}

					case Type.TYPE_FUNC_CEILING:
					{
						return "CEILING";
					}

					case Type.TYPE_FUNC_HYPGEOMDIST:
					{
						return "HYPGEOMDIST";
					}

					case Type.TYPE_FUNC_LOGNORMDIST:
					{
						return "LOGNORMDIST";
					}

					case Type.TYPE_FUNC_LOGINV:
					{
						return "LOGINV";
					}

					case Type.TYPE_FUNC_NEGBINOMDIST:
					{
						return "NEGBINOMDIST";
					}

					case Type.TYPE_FUNC_NORMDIST:
					{
						return "NORMDIST";
					}

					case Type.TYPE_FUNC_NORMSDIST:
					{
						return "NORMSDIST";
					}

					case Type.TYPE_FUNC_NORMINV:
					{
						return "NORMINV";
					}

					case Type.TYPE_FUNC_NORMSINV:
					{
						return "NORMSINV";
					}

					case Type.TYPE_FUNC_STANDARDIZE:
					{
						return "STANDARDIZE";
					}

					case Type.TYPE_FUNC_ODD:
					{
						return "ODD";
					}

					case Type.TYPE_FUNC_PERMUT:
					{
						return "PERMUT";
					}

					case Type.TYPE_FUNC_POISSON:
					{
						return "POISSON";
					}

					case Type.TYPE_FUNC_TDIST:
					{
						return "TDIST";
					}

					case Type.TYPE_FUNC_WEIBULL:
					{
						return "WEIBULL";
					}

					case Type.TYPE_FUNC_SUMXMY2:
					{
						return "SUMXMY2";
					}

					case Type.TYPE_FUNC_SUMX2MY2:
					{
						return "SUMX2MY2";
					}

					case Type.TYPE_FUNC_SUMX2PY2:
					{
						return "SUMX2PY2";
					}

					case Type.TYPE_FUNC_CHITEST:
					{
						return "CHITEST";
					}

					case Type.TYPE_FUNC_CORREL:
					{
						return "CORREL";
					}

					case Type.TYPE_FUNC_COVAR:
					{
						return "COVAR";
					}

					case Type.TYPE_FUNC_FORECAST:
					{
						return "FORECAST";
					}

					case Type.TYPE_FUNC_FTEST:
					{
						return "FTEST";
					}

					case Type.TYPE_FUNC_INTERCEPT:
					{
						return "INTERCEPT";
					}

					case Type.TYPE_FUNC_PEARSON:
					{
						return "PEARSON";
					}

					case Type.TYPE_FUNC_RSQ:
					{
						return "RSQ";
					}

					case Type.TYPE_FUNC_STEYX:
					{
						return "STEYX";
					}

					case Type.TYPE_FUNC_SLOPE:
					{
						return "SLOPE";
					}

					case Type.TYPE_FUNC_TTEST:
					{
						return "TTEST";
					}

					case Type.TYPE_FUNC_PROB:
					{
						return "PROB";
					}

					case Type.TYPE_FUNC_DEVSQ:
					{
						return "DEVSQ";
					}

					case Type.TYPE_FUNC_GEOMEAN:
					{
						return "GEOMEAN";
					}

					case Type.TYPE_FUNC_HARMEAN:
					{
						return "HARMEAN";
					}

					case Type.TYPE_FUNC_SUMSQ:
					{
						return "SUMSQ";
					}

					case Type.TYPE_FUNC_KURT:
					{
						return "KURT";
					}

					case Type.TYPE_FUNC_SKEW:
					{
						return "SKEW";
					}

					case Type.TYPE_FUNC_ZTEST:
					{
						return "ZTEST";
					}

					case Type.TYPE_FUNC_LARGE:
					{
						return "LARGE";
					}

					case Type.TYPE_FUNC_SMALL:
					{
						return "SMALL";
					}

					case Type.TYPE_FUNC_QUARTILE:
					{
						return "QUARTILE";
					}

					case Type.TYPE_FUNC_PERCENTILE:
					{
						return "PERCENTILE";
					}

					case Type.TYPE_FUNC_PERCENTRANK:
					{
						return "PERCENTRANK";
					}

					case Type.TYPE_FUNC_MODE:
					{
						return "MODE";
					}

					case Type.TYPE_FUNC_TRIMMEAN:
					{
						return "TRIMMEAN";
					}

					case Type.TYPE_FUNC_TINV:
					{
						return "TINV";
					}

					case Type.TYPE_FUNC_MOVIE_COMMAND:
					{
						return "MOVIE.COMMAND";
					}

					case Type.TYPE_FUNC_GET_MOVIE:
					{
						return "GET.MOVIE";
					}

					case Type.TYPE_FUNC_CONCATENATE:
					{
						return "CONCATENATE";
					}

					case Type.TYPE_FUNC_POWER:
					{
						return "POWER";
					}

					case Type.TYPE_FUNC_PIVOT_ADD_DATA:
					{
						return "PIVOT.ADD.DATA";
					}

					case Type.TYPE_FUNC_GET_PIVOT_TABLE:
					{
						return "GET.PIVOT.TABLE";
					}

					case Type.TYPE_FUNC_GET_PIVOT_FIELD:
					{
						return "GET.PIVOT.FIELD";
					}

					case Type.TYPE_FUNC_GET_PIVOT_ITEM:
					{
						return "GET.PIVOT.ITEM";
					}

					case Type.TYPE_FUNC_RADIANS:
					{
						return "RADIANS";
					}

					case Type.TYPE_FUNC_DEGREES:
					{
						return "DEGREES";
					}

					case Type.TYPE_FUNC_SUBTOTAL:
					{
						return "SUBTOTAL";
					}

					case Type.TYPE_FUNC_SUMIF:
					{
						return "SUMIF";
					}

					case Type.TYPE_FUNC_COUNTIF:
					{
						return "COUNTIF";
					}

					case Type.TYPE_FUNC_COUNTBLANK:
					{
						return "COUNTBLANK";
					}

					case Type.TYPE_FUNC_SCENARIO_GET:
					{
						return "SCENARIO.GET";
					}

					case Type.TYPE_FUNC_OPTIONS_LISTS_GET:
					{
						return "OPTIONS.LISTS.GET";
					}

					case Type.TYPE_FUNC_ISPMT:
					{
						return "ISPMT";
					}

					case Type.TYPE_FUNC_DATEDIF:
					{
						return "DATEDIF";
					}

					case Type.TYPE_FUNC_DATESTRING:
					{
						return "DATESTRING";
					}

					case Type.TYPE_FUNC_NUMBERSTRING:
					{
						return "NUMBERSTRING";
					}

					case Type.TYPE_FUNC_ROMAN:
					{
						return "ROMAN";
					}

					case Type.TYPE_FUNC_OPEN_DIALOG:
					{
						return "OPEN.DIALOG";
					}

					case Type.TYPE_FUNC_SAVE_DIALOG:
					{
						return "SAVE.DIALOG";
					}

					case Type.TYPE_FUNC_VIEW_GET:
					{
						return "VIEW.GET";
					}

					case Type.TYPE_FUNC_GETPIVOTDATA:
					{
						return "GETPIVOTDATA";
					}

					case Type.TYPE_FUNC_HYPERLINK:
					{
						return "HYPERLINK";
					}

					case Type.TYPE_FUNC_PHONETIC:
					{
						return "PHONETIC";
					}

					case Type.TYPE_FUNC_AVERAGEA:
					{
						return "AVERAGEA";
					}

					case Type.TYPE_FUNC_MAXA:
					{
						return "MAXA";
					}

					case Type.TYPE_FUNC_MINA:
					{
						return "MINA";
					}

					case Type.TYPE_FUNC_STDEVPA:
					{
						return "STDEVPA";
					}

					case Type.TYPE_FUNC_VARPA:
					{
						return "VARPA";
					}

					case Type.TYPE_FUNC_STDEVA:
					{
						return "STDEVA";
					}

					case Type.TYPE_FUNC_VARA:
					{
						return "VARA";
					}

					case Type.TYPE_FUNC_BAHTTEXT:
					{
						return "BAHTTEXT";
					}

					case Type.TYPE_FUNC_THAIDAYOFWEEK:
					{
						return "THAIDAYOFWEEK";
					}

					case Type.TYPE_FUNC_THAIDIGIT:
					{
						return "THAIDIGIT";
					}

					case Type.TYPE_FUNC_THAIMONTHOFYEAR:
					{
						return "THAIMONTHOFYEAR";
					}

					case Type.TYPE_FUNC_THAINUMSOUND:
					{
						return "THAINUMSOUND";
					}

					case Type.TYPE_FUNC_THAINUMSTRING:
					{
						return "THAINUMSTRING";
					}

					case Type.TYPE_FUNC_THAISTRINGLENGTH:
					{
						return "THAISTRINGLENGTH";
					}

					case Type.TYPE_FUNC_ISTHAIDIGIT:
					{
						return "ISTHAIDIGIT";
					}

					case Type.TYPE_FUNC_ROUNDBAHTDOWN:
					{
						return "ROUNDBAHTDOWN";
					}

					case Type.TYPE_FUNC_ROUNDBAHTUP:
					{
						return "ROUNDBAHTUP";
					}

					case Type.TYPE_FUNC_THAIYEAR:
					{
						return "THAIYEAR";
					}

					case Type.TYPE_FUNC_RTD:
					{
						return "RTD";
					}

					case Type.TYPE_PAREN:
					{
						return "";
					}

				}
				return "???";
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SumToken : Token
		{
			public SumToken(byte nParameterCount) : base(Type.TYPE_FUNC_SUM, SubType.SUB_TYPE_FUNCTION, nParameterCount)
			{
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendString("SUM");
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgFuncVarRecord((ushort)(Token.Type.TYPE_FUNC_SUM), m_nParameterCount);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				double fValue = 0.0f;
				if (ppValueVector.GetSize() >= m_nParameterCount)
				{
					for (ushort i = 0; i < m_nParameterCount; i++)
					{
						Value pValue = ppValueVector.Get(ppValueVector.GetSize() - 1);
						switch (pValue.GetType())
						{
							case Value.Type.TYPE_FLOAT:
							{
								fValue += pValue.GetFloat();
								break;
							}

							case Value.Type.TYPE_AREA:
							{
								Area pArea = pValue.m_pImpl.m_pArea;
								for (ushort nX = pArea.m_pTopLeft.m_nX; nX <= pArea.m_pBottomRight.m_nX; nX++)
								{
									for (ushort nY = pArea.m_pTopLeft.m_nY; nY <= pArea.m_pBottomRight.m_nY; nY++)
									{
										Cell pCell = pWorksheetImplementation.m_pWorksheet.GetCell(nX, nY);
										if (pCell.GetType() == Value.Type.TYPE_FLOAT)
											fValue += pCell.GetFloat();
									}
								}
								break;
							}

							case Value.Type.TYPE_AREA_3D:
							{
								Area3d pArea3d = pValue.m_pImpl.m_pArea3d;
								Area pArea = pArea3d.m_pArea;
								for (ushort nSheet = pArea3d.m_pWorksheetRange.m_nFirst; nSheet <= pArea3d.m_pWorksheetRange.m_nLast; nSheet++)
								{
									Worksheet pWorksheet = pWorksheetImplementation.GetWorkbook().GetWorksheetByIndex(nSheet);
									if (pWorksheet == null)
										return false;
									for (ushort nX = pArea.m_pTopLeft.m_nX; nX <= pArea.m_pBottomRight.m_nX; nX++)
									{
										for (ushort nY = pArea.m_pTopLeft.m_nY; nY <= pArea.m_pBottomRight.m_nY; nY++)
										{
											Cell pCell = pWorksheet.GetCell(nX, nY);
											if (pCell.GetType() == Value.Type.TYPE_FLOAT)
												fValue += pCell.GetFloat();
											else if (pCell.GetType() == Value.Type.TYPE_FORMULA)
											{
												Value pResult = pCell.GetValue().m_pImpl.m_pFormula.Evaluate(pWorksheet.m_pImpl, (ushort)(nDepth + 1));
												if (pResult.GetType() == Value.Type.TYPE_FLOAT)
													fValue += pResult.GetFloat();
											}
										}
									}
								}
								break;
							}

						}
						ppValueVector.PopBack();
						{
							pValue = null;
						}
					}
				}
				ppValueVector.PushBack(ValueImplementation.CreateFloatValue(fValue));
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StringToken : Token
		{
			protected InternalString m_sString;
			public StringToken(string sxString) : base(Type.TYPE_STRING, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_sString = new InternalString(sxString);
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				int i;
				sOut.AppendChar('"');
				for (i = 0; i < m_sString.GetLength(); i++)
				{
					char nChar = m_sString.GetChar(i);
					if (nChar == '"')
						sOut.AppendChar('"');
					sOut.AppendChar(nChar);
				}
				sOut.AppendChar('"');
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgStrRecord(m_sString.GetExternalString());
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateStringValue(m_sString.GetExternalString()));
				return true;
			}

			~StringToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SpaceToken : Token
		{
			public enum SpaceType
			{
				TYPE_SPACE_BEFORE_BASE_EXPRESSION = 0x00,
				TYPE_RETURN_BEFORE_BASE_EXPRESSION = 0x01,
				TYPE_SPACE_BEFORE_OPEN = 0x02,
				TYPE_RETURN_BEFORE_OPEN = 0x03,
				TYPE_SPACE_BEFORE_CLOSE = 0x04,
				TYPE_RETURN_BEFORE_CLOSE = 0x05,
				TYPE_SPACE_BEFORE_EXPRESSION = 0x06,
			}

			protected SpaceType m_eSpaceType;
			protected byte m_nCount;
			public SpaceToken(SpaceType eSpaceType, byte nCount) : base(Type.TYPE_SPACE, SubType.SUB_TYPE_VARIABLE, 0)
			{
				nbAssert.Assert(eSpaceType >= SpaceType.TYPE_SPACE_BEFORE_BASE_EXPRESSION && eSpaceType <= SpaceType.TYPE_SPACE_BEFORE_EXPRESSION);
				nbAssert.Assert(nCount > 0);
				m_eSpaceType = eSpaceType;
				m_nCount = nCount;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				char cChar = ' ';
				if (m_eSpaceType == SpaceType.TYPE_RETURN_BEFORE_BASE_EXPRESSION || m_eSpaceType == SpaceType.TYPE_RETURN_BEFORE_OPEN || m_eSpaceType == SpaceType.TYPE_RETURN_BEFORE_CLOSE)
					cChar = '\n';
				for (byte i = 0; i < m_nCount; i++)
					sOut.AppendChar(cChar);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgAttrSpaceRecord((byte)(m_eSpaceType), m_nCount);
			}

			public SpaceToken.SpaceType GetSpaceType()
			{
				return m_eSpaceType;
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OperatorToken : Token
		{
			protected InternalString m_sOperator;
			public OperatorToken(string szOperator) : base(Type.TYPE_OPERATOR, SubType.SUB_TYPE_OPERATOR, 2)
			{
				m_sOperator = new InternalString(szOperator);
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendString(m_sOperator.GetExternalString());
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				if (m_sOperator.IsEqual("+"))
					return new PtgOperatorRecord(0x03);
				if (m_sOperator.IsEqual("-"))
					return new PtgOperatorRecord(0x04);
				if (m_sOperator.IsEqual("*"))
					return new PtgOperatorRecord(0x05);
				if (m_sOperator.IsEqual("/"))
					return new PtgOperatorRecord(0x06);
				if (m_sOperator.IsEqual("^"))
					return new PtgOperatorRecord(0x07);
				if (m_sOperator.IsEqual("&"))
					return new PtgOperatorRecord(0x08);
				if (m_sOperator.IsEqual("<"))
					return new PtgOperatorRecord(0x09);
				if (m_sOperator.IsEqual("<="))
					return new PtgOperatorRecord(0x0A);
				if (m_sOperator.IsEqual("="))
					return new PtgOperatorRecord(0x0B);
				if (m_sOperator.IsEqual(">="))
					return new PtgOperatorRecord(0x0C);
				if (m_sOperator.IsEqual(">"))
					return new PtgOperatorRecord(0x0D);
				if (m_sOperator.IsEqual("<>"))
					return new PtgOperatorRecord(0x0E);
				nbAssert.Assert(false);
				return null;
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				if (ppValueVector.GetSize() >= 2)
				{
					Value a = ppValueVector.Get(ppValueVector.GetSize() - 2);
					Value b = ppValueVector.Get(ppValueVector.GetSize() - 1);
					if (a.GetType() == Value.Type.TYPE_FLOAT && b.GetType() == Value.Type.TYPE_FLOAT)
					{
						Value pOwnedB = ppValueVector.PopBack();
						Value pOwnedA = ppValueVector.PopBack();
						bool bReturnable = true;
						if (m_sOperator.IsEqual("+"))
							ppValueVector.PushBack(ValueImplementation.CreateFloatValue(a.GetFloat() + b.GetFloat()));
						else if (m_sOperator.IsEqual("*"))
							ppValueVector.PushBack(ValueImplementation.CreateFloatValue(a.GetFloat() * b.GetFloat()));
						else if (m_sOperator.IsEqual("-"))
							ppValueVector.PushBack(ValueImplementation.CreateFloatValue(a.GetFloat() - b.GetFloat()));
						else if (m_sOperator.IsEqual("/") && b.GetFloat() != 0.0f)
							ppValueVector.PushBack(ValueImplementation.CreateFloatValue(a.GetFloat() / b.GetFloat()));
						else if (m_sOperator.IsEqual("^"))
							ppValueVector.PushBack(ValueImplementation.CreateFloatValue(Utils.Pow(a.GetFloat(), b.GetFloat())));
						else if (m_sOperator.IsEqual(">"))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() > b.GetFloat()));
						else if (m_sOperator.IsEqual(">="))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() >= b.GetFloat()));
						else if (m_sOperator.IsEqual("<"))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() < b.GetFloat()));
						else if (m_sOperator.IsEqual("<="))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() <= b.GetFloat()));
						else if (m_sOperator.IsEqual("="))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() == b.GetFloat()));
						else if (m_sOperator.IsEqual("<>"))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(a.GetFloat() != b.GetFloat()));
						else
							bReturnable = false;
						{
							return bReturnable;
						}
					}
					if (a.GetType() == Value.Type.TYPE_STRING && b.GetType() == Value.Type.TYPE_STRING)
					{
						Value pOwnedB = ppValueVector.PopBack();
						Value pOwnedA = ppValueVector.PopBack();
						bool bReturnable = true;
						if (m_sOperator.IsEqual("="))
							ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(ExternalString.Equal(a.GetString(), b.GetString())));
						else if (m_sOperator.IsEqual("&"))
						{
							InternalString sTemp = new InternalString(a.GetString());
							sTemp.AppendString(b.GetString());
							ppValueVector.PushBack(ValueImplementation.CreateStringValue(sTemp.GetExternalString()));
							{
								sTemp = null;
							}
						}
						else
							bReturnable = false;
						{
							return bReturnable;
						}
					}
				}
				return false;
			}

			~OperatorToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class NumToken : Token
		{
			protected double m_fNumber;
			public NumToken(double fNumber) : base(Type.TYPE_NUMBER, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_fNumber = fNumber;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendDouble(m_fNumber);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgNumRecord(m_fNumber);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateFloatValue(m_fNumber));
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MissArgToken : Token
		{
			public MissArgToken() : base(Type.TYPE_MISS_ARG, SubType.SUB_TYPE_VARIABLE, 0)
			{
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgMissArgRecord();
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class IntToken : Token
		{
			protected ushort m_nInt;
			public IntToken(ushort nInt) : base(Type.TYPE_INT, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_nInt = nInt;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendUint32(m_nInt);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgIntRecord(m_nInt);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateFloatValue((double)(m_nInt)));
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ParseFunctionData
		{
			public int m_nCount;
		}
		class ParseSpaceData
		{
			public int m_nIndex;
			public char m_cChar;
			public int m_nCount;
			public ParseSpaceData()
			{
				m_nIndex = 0;
				m_cChar = '\0';
				m_nCount = 0;
			}

		}
		class Formula
		{
			protected OwnedVector<Token> m_pTokenVector;
			protected Value m_pValue;
			protected InternalString m_sTemp;
			public Formula(string szFormula, WorksheetImplementation pWorksheetImplementation)
			{
				m_sTemp = new InternalString(szFormula);
				m_pTokenVector = new OwnedVector<Token>();
				m_pValue = null;
				if (m_sTemp.GetLength() > 0 && m_sTemp.GetChar(0) == '=')
				{
					m_sTemp.CropFront(1);
					Parse(m_sTemp, pWorksheetImplementation);
				}
			}

			public Formula(Vector<ParsedExpressionRecord> pParsedExpressionRecordVector, WorkbookGlobals pWorkbookGlobals)
			{
				int i;
				m_sTemp = new InternalString("");
				m_pTokenVector = new OwnedVector<Token>();
				m_pValue = null;
				for (i = 0; i < pParsedExpressionRecordVector.GetSize(); i++)
				{
					ParsedExpressionRecord pParsedExpressionRecord = pParsedExpressionRecordVector.Get(i);
					Token pToken = pParsedExpressionRecord.GetToken(pWorkbookGlobals);
					if (pToken != null)
					{
						NumberDuck.Secret.Token __2538616708 = pToken;
						pToken = null;
						m_pTokenVector.PushBack(__2538616708);
					}
					if (pParsedExpressionRecord.GetType() == ParsedExpressionRecord.Type.TYPE_UNKNOWN)
					{
						break;
					}
				}
			}

			public Formula(OwnedVector<ParsedExpressionRecord> pParsedExpressionRecordVector, WorkbookGlobals pWorkbookGlobals)
			{
				int i;
				m_sTemp = new InternalString("");
				m_pTokenVector = new OwnedVector<Token>();
				m_pValue = null;
				for (i = 0; i < pParsedExpressionRecordVector.GetSize(); i++)
				{
					ParsedExpressionRecord pParsedExpressionRecord = pParsedExpressionRecordVector.Get(i);
					Token pToken = pParsedExpressionRecord.GetToken(pWorkbookGlobals);
					if (pToken != null)
					{
						NumberDuck.Secret.Token __2538616708 = pToken;
						pToken = null;
						m_pTokenVector.PushBack(__2538616708);
					}
					if (pParsedExpressionRecord.GetType() == ParsedExpressionRecord.Type.TYPE_UNKNOWN)
					{
						break;
					}
				}
			}

			public ushort GetNumToken()
			{
				return (ushort)(m_pTokenVector.GetSize());
			}

			public Token GetTokenByIndex(ushort nIndex)
			{
				nbAssert.Assert(nIndex < m_pTokenVector.GetSize());
				return m_pTokenVector.Get(nIndex);
			}

			public Value Evaluate(WorksheetImplementation pWorksheetImplementation, ushort nDepth)
			{
				bool bError = false;
				OwnedVector<Value> ppValueVector = new OwnedVector<Value>();
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
				{
					if (!m_pTokenVector.Get(i).Evaluate(pWorksheetImplementation, ppValueVector, (ushort)(nDepth + 1)))
					{
						bError = true;
						break;
					}
				}
				if (m_pValue != null)
					{
						m_pValue = null;
					}
				if (bError || ppValueVector.GetSize() != 1)
				{
					ppValueVector.Clear();
					m_pValue = ValueImplementation.CreateErrorValue();
				}
				else
				{
					m_pValue = ppValueVector.PopBack();
				}
				{
					return m_pValue;
				}
			}

			public string ToString(WorksheetImplementation pWorksheetImplementation)
			{
				InternalString sPreSpace = new InternalString("");
				InternalString sPostSpace = new InternalString("");
				OwnedVector<InternalString> sParameterVector = new OwnedVector<InternalString>();
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
				{
					Token pToken = m_pTokenVector.Get(i);
					SpaceToken pSpaceToken = null;
					if (pToken.GetType() == Token.Type.TYPE_SPACE)
						pSpaceToken = (SpaceToken)(pToken);
					if (pToken.GetSubType() == Token.SubType.SUB_TYPE_OPERATOR)
					{
						if (sParameterVector.GetSize() < 2 || pToken.GetParameterCount() != 2)
						{
							m_sTemp.Set("=");
							{
								return m_sTemp.GetExternalString();
							}
						}
						InternalString sTempB = sParameterVector.PopBack();
						InternalString sTempA = sParameterVector.PopBack();
						sTempA.AppendString(sPreSpace.GetExternalString());
						pToken.ToString(pWorksheetImplementation, sTempA);
						sTempA.AppendString(sTempB.GetExternalString());
						{
							NumberDuck.Secret.InternalString __2442592686 = sTempA;
							sTempA = null;
							sParameterVector.PushBack(__2442592686);
						}
						sPreSpace.Set("");
						sPostSpace.Set("");
					}
					else if (pToken.GetSubType() == Token.SubType.SUB_TYPE_FUNCTION)
					{
						InternalString sParameters = new InternalString("");
						for (ushort j = 0; j < pToken.GetParameterCount(); j++)
						{
							if (sParameterVector.GetSize() == 0)
							{
								m_sTemp.Set("=");
								{
									return m_sTemp.GetExternalString();
								}
							}
							InternalString sPopped = sParameterVector.PopBack();
							if (j > 0)
								sPopped.AppendString(",");
							sPopped.AppendString(sParameters.GetExternalString());
							{
								sParameters = null;
							}
							{
								NumberDuck.Secret.InternalString __2953685605 = sPopped;
								sPopped = null;
								sParameters = __2953685605;
							}
						}
						InternalString sTemp = new InternalString("");
						sTemp.AppendString(sPreSpace.GetExternalString());
						pToken.ToString(pWorksheetImplementation, sTemp);
						sTemp.AppendString("(");
						sTemp.AppendString(sParameters.GetExternalString());
						sTemp.AppendString(sPostSpace.GetExternalString());
						sTemp.AppendString(")");
						{
							NumberDuck.Secret.InternalString __1006353954 = sTemp;
							sTemp = null;
							sParameterVector.PushBack(__1006353954);
						}
						sPreSpace.Set("");
						sPostSpace.Set("");
					}
					else
					{
						if (pSpaceToken != null)
						{
							if (pSpaceToken.GetSpaceType() == SpaceToken.SpaceType.TYPE_SPACE_BEFORE_BASE_EXPRESSION || pSpaceToken.GetSpaceType() == SpaceToken.SpaceType.TYPE_RETURN_BEFORE_BASE_EXPRESSION)
								pToken.ToString(pWorksheetImplementation, sPreSpace);
							else
								pToken.ToString(pWorksheetImplementation, sPostSpace);
						}
						else
						{
							InternalString sTemp = new InternalString("");
							sTemp.AppendString(sPreSpace.GetExternalString());
							pToken.ToString(pWorksheetImplementation, sTemp);
							sTemp.AppendString(sPostSpace.GetExternalString());
							{
								NumberDuck.Secret.InternalString __1006353954 = sTemp;
								sTemp = null;
								sParameterVector.PushBack(__1006353954);
							}
							sPreSpace.Set("");
							sPostSpace.Set("");
						}
					}
				}
				m_sTemp.Set("=");
				if (sParameterVector.GetSize() > 0)
					m_sTemp.AppendString(sParameterVector.Get(0).GetExternalString());
				{
					return m_sTemp.GetExternalString();
				}
			}

			public void ToRgce(RgceStruct pRgce, WorkbookGlobals pWorkbookGlobals)
			{
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
				{
					Token pToken = m_pTokenVector.Get(i);
					ParsedExpressionRecord pTemp = pToken.ToParsedExpression(pWorkbookGlobals);
					nbAssert.Assert(pTemp != null);
					{
						NumberDuck.Secret.ParsedExpressionRecord __432555651 = pTemp;
						pTemp = null;
						pRgce.m_pParsedExpressionRecordVector.PushBack(__432555651);
					}
				}
			}

			public void InsertColumn(ushort nWorksheet, ushort nColumn)
			{
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
					m_pTokenVector.Get(i).InsertColumn(nWorksheet, nColumn);
			}

			public void DeleteColumn(ushort nWorksheet, ushort nColumn)
			{
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
					m_pTokenVector.Get(i).DeleteColumn(nWorksheet, nColumn);
			}

			public void InsertRow(ushort nWorksheet, ushort nRow)
			{
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
					m_pTokenVector.Get(i).InsertRow(nWorksheet, nRow);
			}

			public void DeleteRow(ushort nWorksheet, ushort nRow)
			{
				for (int i = 0; i < m_pTokenVector.GetSize(); i++)
					m_pTokenVector.Get(i).DeleteRow(nWorksheet, nRow);
			}

			public bool ValidateForChart(WorksheetImplementation pWorksheetImplementation)
			{
				if (GetNumToken() == 1)
				{
					Token pToken = GetTokenByIndex(0);
					if (pToken.GetType() == Token.Type.TYPE_AREA_3D)
					{
						Area3d pArea3d = ((Area3dToken)(pToken)).GetArea3d().CreateClone();
						pArea3d.m_pArea.m_pTopLeft.m_bXRelative = false;
						pArea3d.m_pArea.m_pTopLeft.m_bYRelative = false;
						pArea3d.m_pArea.m_pBottomRight.m_bXRelative = false;
						pArea3d.m_pArea.m_pBottomRight.m_bYRelative = false;
						m_pTokenVector.Erase(0);
						{
							NumberDuck.Secret.Area3d __2738670685 = pArea3d;
							pArea3d = null;
							m_pTokenVector.Insert(0, new Area3dToken(__2738670685));
						}
					}
					else if (pToken.GetType() == Token.Type.TYPE_AREA)
					{
						ushort nWorksheet = 0;
						Workbook pWorkbook = pWorksheetImplementation.GetWorkbook();
						for (ushort i = 0; i < pWorkbook.GetNumWorksheet(); i++)
						{
							if (pWorkbook.GetWorksheetByIndex(i).m_pImpl == pWorksheetImplementation)
							{
								nWorksheet = i;
								break;
							}
						}
						Area pArea = ((AreaToken)(pToken)).GetArea().CreateClone();
						pArea.m_pTopLeft.m_bXRelative = false;
						pArea.m_pTopLeft.m_bYRelative = false;
						pArea.m_pBottomRight.m_bXRelative = false;
						pArea.m_pBottomRight.m_bYRelative = false;
						Area3d pArea3d;
						{
							NumberDuck.Secret.Area __4245081970 = pArea;
							pArea = null;
							pArea3d = new Area3d(nWorksheet, nWorksheet, __4245081970);
						}
						m_pTokenVector.Erase(0);
						{
							NumberDuck.Secret.Area3d __2738670685 = pArea3d;
							pArea3d = null;
							m_pTokenVector.Insert(0, new Area3dToken(__2738670685));
						}
					}
					return true;
				}
				return false;
			}

			public string ToChartString(WorksheetImplementation pWorksheetImplementation)
			{
				m_sTemp.Set("=");
				if (GetNumToken() == 1)
				{
					Token pToken = GetTokenByIndex(0);
					if (pToken.GetType() == Token.Type.TYPE_AREA_3D)
					{
						Area3d pArea3d = ((Area3dToken)(pToken)).GetArea3d().CreateClone();
						pArea3d.m_pArea.m_pTopLeft.m_bXRelative = true;
						pArea3d.m_pArea.m_pTopLeft.m_bYRelative = true;
						pArea3d.m_pArea.m_pBottomRight.m_bXRelative = true;
						pArea3d.m_pArea.m_pBottomRight.m_bYRelative = true;
						pWorksheetImplementation.Area3dToAddress(pArea3d, m_sTemp);
						{
							pArea3d = null;
						}
					}
				}
				return m_sTemp.GetExternalString();
			}

			public bool ValidateForChartName(WorksheetImplementation pWorksheetImplementation)
			{
				if (GetNumToken() == 1)
				{
					Token pToken = GetTokenByIndex(0);
					if (pToken.GetType() == Token.Type.TYPE_COORDINATE_3D)
					{
						Coordinate3d pCoordinate3d = ((Coordinate3dToken)(pToken)).GetCoordinate3d().CreateClone();
						pCoordinate3d.m_pCoordinate.m_bXRelative = false;
						pCoordinate3d.m_pCoordinate.m_bYRelative = false;
						m_pTokenVector.Erase(0);
						{
							NumberDuck.Secret.Coordinate3d __1094936853 = pCoordinate3d;
							pCoordinate3d = null;
							m_pTokenVector.Insert(0, new Coordinate3dToken(__1094936853));
						}
						{
							return true;
						}
					}
					else if (pToken.GetType() == Token.Type.TYPE_COORDINATE)
					{
						ushort nWorksheet = 0;
						Workbook pWorkbook = pWorksheetImplementation.GetWorkbook();
						for (ushort i = 0; i < pWorkbook.GetNumWorksheet(); i++)
						{
							if (pWorkbook.GetWorksheetByIndex(i).m_pImpl == pWorksheetImplementation)
							{
								nWorksheet = i;
								break;
							}
						}
						Coordinate pCoordinate = ((CoordinateToken)(pToken)).GetCoordinate().CreateClone();
						pCoordinate.m_bXRelative = false;
						pCoordinate.m_bYRelative = false;
						Coordinate3d pCoordinate3d;
						{
							NumberDuck.Secret.Coordinate __3642692973 = pCoordinate;
							pCoordinate = null;
							pCoordinate3d = new Coordinate3d(nWorksheet, nWorksheet, __3642692973);
						}
						m_pTokenVector.Erase(0);
						{
							NumberDuck.Secret.Coordinate3d __1094936853 = pCoordinate3d;
							pCoordinate3d = null;
							m_pTokenVector.Insert(0, new Coordinate3dToken(__1094936853));
						}
						{
							return true;
						}
					}
				}
				return false;
			}

			public string ToChartNameString(WorksheetImplementation pWorksheetImplementation)
			{
				m_sTemp.Set("=");
				if (GetNumToken() == 1)
				{
					Token pToken = GetTokenByIndex(0);
					if (pToken.GetType() == Token.Type.TYPE_COORDINATE_3D)
					{
						Coordinate3d pCoordinate3d = ((Coordinate3dToken)(pToken)).GetCoordinate3d().CreateClone();
						pCoordinate3d.m_pCoordinate.m_bXRelative = true;
						pCoordinate3d.m_pCoordinate.m_bYRelative = true;
						pCoordinate3d.ToString(pWorksheetImplementation, m_sTemp);
						{
							pCoordinate3d = null;
						}
					}
				}
				return m_sTemp.GetExternalString();
			}

			protected byte Parse(InternalString sFormula, WorksheetImplementation pWorksheetImplementation)
			{
				if (sFormula.GetLength() == 0)
					return 0;
				{
					ParseSpaceData pParseSpaceData = new ParseSpaceData();
					while (ParseSpace(sFormula, pParseSpaceData))
					{
						SpaceToken.SpaceType eSpaceType = SpaceToken.SpaceType.TYPE_SPACE_BEFORE_BASE_EXPRESSION;
						if (pParseSpaceData.m_cChar == '\n')
							eSpaceType = SpaceToken.SpaceType.TYPE_RETURN_BEFORE_BASE_EXPRESSION;
						m_pTokenVector.PushBack(new SpaceToken(eSpaceType, (byte)(pParseSpaceData.m_nCount)));
					}
					int nIndex = pParseSpaceData.m_nIndex;
					if (nIndex > 0)
					{
						sFormula.CropFront(nIndex);
						{
							return Parse(sFormula, pWorksheetImplementation);
						}
					}
				}
				if (ParseBool(sFormula, pWorksheetImplementation))
					return 1;
				if (ParseInt(sFormula))
					return 1;
				if (ParseFloat(sFormula))
					return 1;
				ParseFunctionData pParseFunctionData = new ParseFunctionData();
				if (ParseFunction(sFormula, "SUM", pParseFunctionData, pWorksheetImplementation))
				{
					m_pTokenVector.PushBack(new SumToken((byte)(pParseFunctionData.m_nCount)));
					{
						return 1;
					}
				}
				for (int i = 0; i <= (int)(Token.Type.TYPE_FUNC_RTD); i++)
				{
					Token.Type e = (Token.Type)(i);
					if (ParseFunction(sFormula, Token.GetTypeName(e), pParseFunctionData, pWorksheetImplementation))
					{
						m_pTokenVector.PushBack(new Token(e, Token.SubType.SUB_TYPE_FUNCTION, (byte)(pParseFunctionData.m_nCount)));
						{
							return 1;
						}
					}
				}
				if (ParseFunction(sFormula, "", pParseFunctionData, pWorksheetImplementation))
				{
					m_pTokenVector.PushBack(new Token(Token.Type.TYPE_PAREN, Token.SubType.SUB_TYPE_FUNCTION, (byte)(pParseFunctionData.m_nCount)));
					{
						return 1;
					}
				}
				Coordinate pCoordinate = WorksheetImplementation.AddressToCoordinate(sFormula.GetExternalString());
				if (pCoordinate != null)
				{
					{
						NumberDuck.Secret.Coordinate __3642692973 = pCoordinate;
						pCoordinate = null;
						m_pTokenVector.PushBack(new CoordinateToken(__3642692973));
					}
					{
						return 1;
					}
				}
				Area pArea = WorksheetImplementation.AddressToArea(sFormula.GetExternalString());
				if (pArea != null)
				{
					{
						NumberDuck.Secret.Area __4245081970 = pArea;
						pArea = null;
						m_pTokenVector.PushBack(new AreaToken(__4245081970));
					}
					{
						return 1;
					}
				}
				Coordinate3d pCoordinate3d = pWorksheetImplementation.ParseCoordinate3d(sFormula);
				if (pCoordinate3d != null)
				{
					{
						NumberDuck.Secret.Coordinate3d __1094936853 = pCoordinate3d;
						pCoordinate3d = null;
						m_pTokenVector.PushBack(new Coordinate3dToken(__1094936853));
					}
					{
						return 1;
					}
				}
				Area3d pArea3d = pWorksheetImplementation.ParseArea3d(sFormula);
				if (pArea3d != null)
				{
					{
						NumberDuck.Secret.Area3d __2738670685 = pArea3d;
						pArea3d = null;
						m_pTokenVector.PushBack(new Area3dToken(__2738670685));
					}
					{
						return 1;
					}
				}
				if (ParseString(sFormula))
				{
					InternalString sTemp = new InternalString("");
					for (int i = 1; i < sFormula.GetLength() - 1; i++)
					{
						char nChar = sFormula.GetChar(i);
						sTemp.AppendChar(nChar);
						if (nChar == '"')
							i++;
					}
					m_pTokenVector.PushBack(new StringToken(sTemp.GetExternalString()));
					{
						return 1;
					}
				}
				InternalString sToken = new InternalString("");
				InternalString sOperator = new InternalString("");
				InternalString sTrailingSpaces = new InternalString("");
				InternalString sNextOperator = new InternalString("");
				InternalString sLastOperator = new InternalString("");
				InternalString sLastTrailingSpaces = new InternalString("");
				int nCount = 0;
				ushort nQuoteDepth = 0;
				ushort nParenDepth = 0;
				for (int i = 0; i < sFormula.GetLength(); i++)
				{
					char nChar = sFormula.GetChar(i);
					if (nChar == '"')
					{
						if (nQuoteDepth == 1)
						{
							if (i < sFormula.GetLength() - 1 && sFormula.GetChar(i + 1) == '"')
								nQuoteDepth++;
							else
								nQuoteDepth--;
						}
						else if (nQuoteDepth == 2)
						{
							nQuoteDepth--;
						}
						else
							nQuoteDepth++;
					}
					if (nQuoteDepth == 0)
					{
						if (nChar == '(')
							nParenDepth++;
						if (nChar == ')')
							if (nParenDepth == 0)
							{
								{
									return 0;
								}
							}
							else
								nParenDepth--;
					}
					sNextOperator.Set("");
					if (nChar == '<' && i + 1 < sFormula.GetLength() && sFormula.GetChar(i + 1) == '>')
						sNextOperator.Set("<>");
					else if (nChar == '<' && i + 1 < sFormula.GetLength() && sFormula.GetChar(i + 1) == '=')
						sNextOperator.Set("<=");
					else if (nChar == '>' && i + 1 < sFormula.GetLength() && sFormula.GetChar(i + 1) == '=')
						sNextOperator.Set(">=");
					else if (nChar == '+' || nChar == '-' || nChar == '*' || nChar == '/' || nChar == '<' || nChar == '>' || nChar == '^' || nChar == '&' || nChar == '=')
						sNextOperator.AppendChar(nChar);
					if (nQuoteDepth == 0 && nParenDepth == 0 && (nChar == ',' || sNextOperator.GetLength() > 0))
					{
						sLastOperator.Set(sOperator.GetExternalString());
						sLastTrailingSpaces.Set(sTrailingSpaces.GetExternalString());
						int nIndex = sToken.GetLength();
						while (nIndex > 0 && (sToken.GetChar(nIndex - 1) == ' ' || sToken.GetChar(nIndex - 1) == '\n'))
						{
							sTrailingSpaces.PrependChar(sToken.GetChar(nIndex - 1));
							nIndex--;
						}
						sToken.SubStr(0, nIndex);
						if (sToken.GetLength() > 0)
						{
							sOperator.Set(sNextOperator.GetExternalString());
							nCount = nCount + Parse(sToken, pWorksheetImplementation);
							sToken.Set("");
						}
						if (sLastOperator.GetLength() > 0 && nCount >= 2)
						{
							InsertOperator(sLastOperator, sLastTrailingSpaces);
							nCount--;
						}
					}
					else
					{
						sToken.AppendChar(nChar);
					}
				}
				if (!ExternalString.Equal(sToken.GetExternalString(), sFormula.GetExternalString()))
					nCount = nCount + Parse(sToken, pWorksheetImplementation);
				if (sOperator.GetLength() > 0 && nCount >= 2)
				{
					InsertOperator(sOperator, sTrailingSpaces);
					nCount--;
				}
				{
					return (byte)(nCount);
				}
			}

			protected bool ParseFunction(InternalString sFormula, string szFunction, ParseFunctionData pData, WorksheetImplementation pWorksheetImplementation)
			{
				InternalString sTemp = new InternalString(szFunction);
				sTemp.AppendChar('(');
				int nFunctionLength = sTemp.GetLength();
				bool bStartsWith = sFormula.StartsWith(sTemp.GetExternalString());
				bool bEndsWith = sFormula.EndsWith(")");
				{
					sTemp = null;
				}
				if (bStartsWith && bEndsWith)
				{
					ushort nQuoteDepth = 0;
					ushort nParenDepth = 0;
					for (int i = nFunctionLength; i < sFormula.GetLength() - 1; i++)
					{
						if (sFormula.GetChar(i) == '"')
						{
							if (nQuoteDepth == 1)
							{
								if (i < sFormula.GetLength() - 1 && sFormula.GetChar(i + 1) == '"')
									i++;
								else
									nQuoteDepth--;
							}
							else
								nQuoteDepth++;
						}
						if (nQuoteDepth == 0)
						{
							if (sFormula.GetChar(i) == '(')
								nParenDepth++;
							if (sFormula.GetChar(i) == ')')
								if (nParenDepth == 0)
								{
									return false;
								}
								else
									nParenDepth--;
						}
					}
					if (nQuoteDepth == 0 && nParenDepth == 0)
					{
						ParseSpaceData pParseSpaceData = new ParseSpaceData();
						pParseSpaceData.m_nIndex = (int)(sFormula.GetLength() - 1);
						while (true)
						{
							if (pParseSpaceData.m_nIndex <= 0)
								break;
							char nChar = sFormula.GetChar(pParseSpaceData.m_nIndex - 1);
							if (nChar != ' ' && nChar != '\n')
								break;
							pParseSpaceData.m_nIndex--;
						}
						int nSpaceStart = pParseSpaceData.m_nIndex;
						InternalString sCount = new InternalString(sFormula.GetExternalString());
						sCount.SubStr(nFunctionLength, nSpaceStart - nFunctionLength);
						pData.m_nCount = Parse(sCount, pWorksheetImplementation);
						{
							sCount = null;
						}
						while (ParseSpace(sFormula, pParseSpaceData))
						{
							SpaceToken.SpaceType eSpaceType = SpaceToken.SpaceType.TYPE_SPACE_BEFORE_CLOSE;
							if (pParseSpaceData.m_cChar == '\n')
								eSpaceType = SpaceToken.SpaceType.TYPE_RETURN_BEFORE_CLOSE;
							m_pTokenVector.PushBack(new SpaceToken(eSpaceType, (byte)(pParseSpaceData.m_nCount)));
						}
						{
							return true;
						}
					}
				}
				{
					return false;
				}
			}

			protected bool ParseSpace(InternalString sFormula, ParseSpaceData pData)
			{
				pData.m_nCount = 0;
				while (pData.m_nIndex < sFormula.GetLength() && sFormula.GetChar(pData.m_nIndex) == ' ')
				{
					pData.m_cChar = ' ';
					pData.m_nIndex++;
					pData.m_nCount++;
				}
				if (pData.m_nCount > 0)
					return true;
				while (pData.m_nIndex < sFormula.GetLength() && sFormula.GetChar(pData.m_nIndex) == '\n')
				{
					pData.m_cChar = '\n';
					pData.m_nIndex++;
					pData.m_nCount++;
				}
				if (pData.m_nCount > 0)
					return true;
				return false;
			}

			protected bool ParseString(InternalString sFormula)
			{
				int nLength = sFormula.GetLength();
				if (nLength >= 2)
					if (sFormula.GetChar(0) == '"' && sFormula.GetChar(nLength - 1) == '"')
					{
						for (int i = 1; i < nLength - 1; i++)
						{
							if (sFormula.GetChar(i) == '"')
								if (i < nLength - 2 && sFormula.GetChar(i + 1) == '"')
									i++;
								else
									return false;
						}
						return true;
					}
				return false;
			}

			protected bool ParseBool(InternalString sFormula, WorksheetImplementation pWorksheetImplementation)
			{
				ParseFunctionData pParseFunctionData = new ParseFunctionData();
				Token pToken = null;
				if (sFormula.IsEqual("FALSE"))
					pToken = new BoolToken(false, false);
				else if (sFormula.IsEqual("TRUE"))
					pToken = new BoolToken(true, false);
				else if (ParseFunction(sFormula, "FALSE", pParseFunctionData, pWorksheetImplementation))
					pToken = new BoolToken(false, true);
				else if (ParseFunction(sFormula, "TRUE", pParseFunctionData, pWorksheetImplementation))
					pToken = new BoolToken(true, true);
				if (pToken != null)
				{
					{
						NumberDuck.Secret.Token __2538616708 = pToken;
						pToken = null;
						m_pTokenVector.PushBack(__2538616708);
					}
					{
						return true;
					}
				}
				{
					return false;
				}
			}

			protected bool ParseInt(InternalString sFormula)
			{
				ushort nInt = 0;
				ushort nMultiplier = 1;
				for (int i = 0; i < sFormula.GetLength(); i++)
				{
					int nIndex = sFormula.GetLength() - 1 - i;
					char nChar = sFormula.GetChar(nIndex);
					if (nChar < '0' || nChar > '9')
						return false;
					nInt = (ushort)(nInt + (nChar - '0') * nMultiplier);
					nMultiplier = (ushort)(nMultiplier * 10);
				}
				m_pTokenVector.PushBack(new IntToken(nInt));
				return true;
			}

			protected bool ParseFloat(InternalString sFormula)
			{
				bool bNumber = false;
				bool bDecimalPoint = false;
				for (int i = 0; i < sFormula.GetLength(); i++)
				{
					char nChar = sFormula.GetChar(i);
					if (nChar == '-' && i == 0)
						continue;
					if (nChar >= '0' && nChar <= '9')
					{
						bNumber = true;
						continue;
					}
					if (nChar == '.' && bNumber && !bDecimalPoint)
					{
						bDecimalPoint = true;
						continue;
					}
					return false;
				}
				double fTemp = sFormula.ParseDouble();
				m_pTokenVector.PushBack(new NumToken(fTemp));
				return true;
			}

			protected void InsertOperator(InternalString sOperator, InternalString sTrailingSpaces)
			{
				ParseSpaceData pParseSpaceData = new ParseSpaceData();
				while (ParseSpace(sTrailingSpaces, pParseSpaceData))
				{
					SpaceToken.SpaceType eSpaceType = SpaceToken.SpaceType.TYPE_SPACE_BEFORE_BASE_EXPRESSION;
					if (pParseSpaceData.m_cChar == '\n')
						eSpaceType = SpaceToken.SpaceType.TYPE_RETURN_BEFORE_BASE_EXPRESSION;
					m_pTokenVector.PushBack(new SpaceToken(eSpaceType, (byte)(pParseSpaceData.m_nCount)));
				}
				m_pTokenVector.PushBack(new OperatorToken(sOperator.GetExternalString()));
			}

			~Formula()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CoordinateToken : Token
		{
			protected Coordinate m_pCoordinate;
			public CoordinateToken(Coordinate pCoordinate) : base(Type.TYPE_COORDINATE, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_pCoordinate = pCoordinate;
			}

			public Coordinate GetCoordinate()
			{
				return m_pCoordinate;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				WorksheetImplementation.CoordinateToAddress(m_pCoordinate, sOut);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgRefRecord(m_pCoordinate);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				Cell pCell = pWorksheetImplementation.m_pWorksheet.GetCell(m_pCoordinate.m_nX, m_pCoordinate.m_nY);
				Value pValue = pCell.GetValue();
				ppValueVector.PushBack(ValueImplementation.CopyValue(pValue));
				return true;
			}

			~CoordinateToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Coordinate3dToken : Token
		{
			protected Coordinate3d m_pCoordinate3d;
			public Coordinate3dToken(Coordinate3d pCoordinate3d) : base(Type.TYPE_COORDINATE_3D, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_pCoordinate3d = pCoordinate3d;
			}

			public Coordinate3d GetCoordinate3d()
			{
				return m_pCoordinate3d;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				m_pCoordinate3d.ToString(pWorksheetImplementation, sOut);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgRef3dRecord(m_pCoordinate3d, pWorkbookGlobals);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				if (m_pCoordinate3d.m_nWorksheetFirst == m_pCoordinate3d.m_nWorksheetLast)
				{
					Cell pCell = pWorksheetImplementation.GetWorkbook().GetWorksheetByIndex(m_pCoordinate3d.m_nWorksheetFirst).GetCell(m_pCoordinate3d.m_pCoordinate.m_nX, m_pCoordinate3d.m_pCoordinate.m_nY);
					Value pValue = pCell.GetValue();
					ppValueVector.PushBack(ValueImplementation.CopyValue(pValue));
				}
				else
				{
					Area pArea;
					pArea = new Area(m_pCoordinate3d.m_pCoordinate.CreateClone(), m_pCoordinate3d.m_pCoordinate.CreateClone());
					Area3d pArea3d;
					{
						NumberDuck.Secret.Area __4245081970 = pArea;
						pArea = null;
						pArea3d = new Area3d(m_pCoordinate3d.m_nWorksheetFirst, m_pCoordinate3d.m_nWorksheetLast, __4245081970);
					}
					{
						NumberDuck.Secret.Area3d __2738670685 = pArea3d;
						pArea3d = null;
						ppValueVector.PushBack(ValueImplementation.CreateArea3dValue(__2738670685));
					}
				}
				return true;
			}

			public override void InsertColumn(ushort nWorksheet, ushort nColumn)
			{
				if (nWorksheet == m_pCoordinate3d.m_nWorksheetLast && nWorksheet == m_pCoordinate3d.m_nWorksheetFirst)
				{
					if (m_pCoordinate3d.m_pCoordinate.m_nX >= nColumn && m_pCoordinate3d.m_pCoordinate.m_nX < Worksheet.MAX_COLUMN)
						m_pCoordinate3d.m_pCoordinate.m_nX++;
				}
			}

			public override void DeleteColumn(ushort nWorksheet, ushort nColumn)
			{
				if (nWorksheet == m_pCoordinate3d.m_nWorksheetLast && nWorksheet == m_pCoordinate3d.m_nWorksheetFirst)
				{
					if (m_pCoordinate3d.m_pCoordinate.m_nX > nColumn)
						m_pCoordinate3d.m_pCoordinate.m_nX--;
				}
			}

			public override void InsertRow(ushort nWorksheet, ushort nRow)
			{
				if (nWorksheet == m_pCoordinate3d.m_nWorksheetLast && nWorksheet == m_pCoordinate3d.m_nWorksheetFirst)
				{
					if (m_pCoordinate3d.m_pCoordinate.m_nY >= nRow && m_pCoordinate3d.m_pCoordinate.m_nY < Worksheet.MAX_ROW)
						m_pCoordinate3d.m_pCoordinate.m_nY++;
				}
			}

			public override void DeleteRow(ushort nWorksheet, ushort nRow)
			{
				if (nWorksheet == m_pCoordinate3d.m_nWorksheetLast && nWorksheet == m_pCoordinate3d.m_nWorksheetFirst)
				{
					if (m_pCoordinate3d.m_pCoordinate.m_nY > nRow)
						m_pCoordinate3d.m_pCoordinate.m_nY--;
				}
			}

			~Coordinate3dToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class BoolToken : Token
		{
			protected bool m_bBool;
			public BoolToken(bool bBool, bool bFunction) : base(Type.TYPE_BOOL, bFunction ? SubType.SUB_TYPE_FUNCTION : SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_bBool = bBool;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				sOut.AppendString(m_bBool ? "TRUE" : "FALSE");
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				if (m_eSubType == SubType.SUB_TYPE_FUNCTION)
					if (m_bBool)
						return new PtgFuncVarRecord((ushort)(Token.Type.TYPE_FUNC_TRUE), m_nParameterCount);
					else
						return new PtgFuncVarRecord((ushort)(Token.Type.TYPE_FUNC_FALSE), m_nParameterCount);
				return new PtgBoolRecord(m_bBool);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateBooleanValue(m_bBool));
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class AreaToken : Token
		{
			protected Area m_pArea;
			public AreaToken(Area pArea) : base(Type.TYPE_AREA, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_pArea = pArea;
			}

			public Area GetArea()
			{
				return m_pArea;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				WorksheetImplementation.AreaToAddress(m_pArea, sOut);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgAreaRecord(m_pArea);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateAreaValue(m_pArea.CreateClone()));
				return true;
			}

			~AreaToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class Area3dToken : Token
		{
			protected Area3d m_pArea3d;
			public Area3dToken(Area3d pArea3d) : base(Type.TYPE_AREA_3D, SubType.SUB_TYPE_VARIABLE, 0)
			{
				m_pArea3d = pArea3d;
			}

			public Area3d GetArea3d()
			{
				return m_pArea3d;
			}

			public override void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				pWorksheetImplementation.Area3dToAddress(m_pArea3d, sOut);
			}

			public override ParsedExpressionRecord ToParsedExpression(WorkbookGlobals pWorkbookGlobals)
			{
				return new PtgArea3dRecord(m_pArea3d, pWorkbookGlobals);
			}

			public override bool Evaluate(WorksheetImplementation pWorksheetImplementation, OwnedVector<Value> ppValueVector, ushort nDepth)
			{
				ppValueVector.PushBack(ValueImplementation.CreateArea3dValue(m_pArea3d.CreateClone()));
				return true;
			}

			public override void InsertColumn(ushort nWorksheet, ushort nColumn)
			{
				if (nWorksheet == m_pArea3d.m_pWorksheetRange.m_nLast && nWorksheet == m_pArea3d.m_pWorksheetRange.m_nFirst)
				{
					if (m_pArea3d.m_pArea.m_pTopLeft.m_nX >= nColumn && m_pArea3d.m_pArea.m_pTopLeft.m_nX < Worksheet.MAX_COLUMN)
						m_pArea3d.m_pArea.m_pTopLeft.m_nX++;
					if (m_pArea3d.m_pArea.m_pBottomRight.m_nX >= nColumn && m_pArea3d.m_pArea.m_pTopLeft.m_nX < Worksheet.MAX_COLUMN)
						m_pArea3d.m_pArea.m_pBottomRight.m_nX++;
				}
			}

			public override void DeleteColumn(ushort nWorksheet, ushort nColumn)
			{
				if (nWorksheet == m_pArea3d.m_pWorksheetRange.m_nLast && nWorksheet == m_pArea3d.m_pWorksheetRange.m_nFirst)
				{
					if (m_pArea3d.m_pArea.m_pTopLeft.m_nX > nColumn)
						m_pArea3d.m_pArea.m_pTopLeft.m_nX--;
					if (m_pArea3d.m_pArea.m_pBottomRight.m_nX > nColumn)
						m_pArea3d.m_pArea.m_pBottomRight.m_nX--;
				}
			}

			public override void InsertRow(ushort nWorksheet, ushort nRow)
			{
				if (nWorksheet == m_pArea3d.m_pWorksheetRange.m_nLast && nWorksheet == m_pArea3d.m_pWorksheetRange.m_nFirst)
				{
					if (m_pArea3d.m_pArea.m_pTopLeft.m_nY >= nRow && m_pArea3d.m_pArea.m_pTopLeft.m_nY < Worksheet.MAX_ROW)
						m_pArea3d.m_pArea.m_pTopLeft.m_nY++;
					if (m_pArea3d.m_pArea.m_pBottomRight.m_nY >= nRow && m_pArea3d.m_pArea.m_pTopLeft.m_nY < Worksheet.MAX_ROW)
						m_pArea3d.m_pArea.m_pBottomRight.m_nY++;
				}
			}

			public override void DeleteRow(ushort nWorksheet, ushort nRow)
			{
				if (nWorksheet == m_pArea3d.m_pWorksheetRange.m_nLast && nWorksheet == m_pArea3d.m_pWorksheetRange.m_nFirst)
				{
					if (m_pArea3d.m_pArea.m_pTopLeft.m_nY > nRow)
						m_pArea3d.m_pArea.m_pTopLeft.m_nY--;
					if (m_pArea3d.m_pArea.m_pBottomRight.m_nY > nRow)
						m_pArea3d.m_pArea.m_pBottomRight.m_nY--;
				}
			}

			~Area3dToken()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class RedBlackNode
		{
			public enum Color
			{
				COLOR_RED = 0,
				COLOR_BLACK,
			}

			public virtual RedBlackNode.Color GetColor()
			{
				return Color.COLOR_RED;
			}

			public virtual RedBlackNode GetParent()
			{
				return null;
			}

			public virtual RedBlackNode GetLeftChild()
			{
				return null;
			}

			public virtual RedBlackNode GetRightChild()
			{
				return null;
			}

			public virtual RedBlackNode GetChild(int nDirection)
			{
				return null;
			}

			public virtual object GetObject()
			{
				return null;
			}

		}
		class RedBlackNodeImplementation : RedBlackNode
		{
			public Color m_eColor;
			public RedBlackNodeImplementation m_pParent;
			public RedBlackNodeImplementation[] m_pChild = new RedBlackNodeImplementation[2];
			public object m_pObject;
			public RedBlackNodeImplementation(object pObject)
			{
				m_eColor = RedBlackNode.Color.COLOR_BLACK;
				m_pObject = pObject;
				m_pParent = null;
				m_pChild[0] = null;
				m_pChild[1] = null;
			}

			public RedBlackNodeImplementation GetUncle()
			{
				RedBlackNodeImplementation pGrandparent = GetGrandparent();
				if (pGrandparent != null)
				{
					if (pGrandparent.m_pChild[0] == m_pParent)
						return pGrandparent.m_pChild[1];
					else
						return pGrandparent.m_pChild[0];
				}
				return null;
			}

			public RedBlackNodeImplementation GetGrandparent()
			{
				if (m_pParent != null)
					return m_pParent.m_pParent;
				return null;
			}

			public override RedBlackNode.Color GetColor()
			{
				return m_eColor;
			}

			public override RedBlackNode GetParent()
			{
				return m_pParent;
			}

			public override RedBlackNode GetLeftChild()
			{
				return m_pChild[0];
			}

			public override RedBlackNode GetRightChild()
			{
				return m_pChild[1];
			}

			public override RedBlackNode GetChild(int nDirection)
			{
				return m_pChild[nDirection];
			}

			public override object GetObject()
			{
				return m_pObject;
			}

			~RedBlackNodeImplementation()
			{
			}

		}
		class RedBlackTree
		{
			public delegate int ComparisonCallback(object pObjectA, object pObjectB);
			protected RedBlackTreeImplementation m_pImpl;
			public RedBlackTree(ComparisonCallback pComparisonCallback)
			{
				nbAssert.Assert(pComparisonCallback != null);
				m_pImpl = new RedBlackTreeImplementation();
				m_pImpl.m_pComparisonCallback = pComparisonCallback;
				m_pImpl.m_pRootNode = null;
			}

			~RedBlackTree()
			{
				{
					NumberDuck.Secret.RedBlackNodeImplementation __608107594 = m_pImpl.m_pRootNode;
					m_pImpl.m_pRootNode = null;
					m_pImpl.RecursiveDelete(__608107594);
				}
				{
					m_pImpl = null;
				}
			}

			public bool AddObject(object pObject)
			{
				if (pObject == null)
					return false;
				RedBlackNodeImplementation pNewNode = new RedBlackNodeImplementation(pObject);
				if (m_pImpl.m_pRootNode == null)
				{
					{
						NumberDuck.Secret.RedBlackNodeImplementation __615123591 = pNewNode;
						pNewNode = null;
						m_pImpl.m_pRootNode = __615123591;
					}
					{
						return true;
					}
				}
				pNewNode.m_eColor = RedBlackNode.Color.COLOR_RED;
				RedBlackNodeImplementation pNode = m_pImpl.m_pRootNode;
				while (true)
				{
					int nComparison = m_pImpl.m_pComparisonCallback(pNode.GetObject(), pObject);
					nbAssert.Assert(nComparison >= -1);
					nbAssert.Assert(nComparison <= 1);
					if (nComparison == 0)
					{
						return false;
					}
					int nDirection = 0;
					if (nComparison > 0)
						nDirection = nComparison;
					if (pNode.m_pChild[nDirection] == null)
					{
						RedBlackNodeImplementation pTempNode = pNewNode;
						{
							NumberDuck.Secret.RedBlackNodeImplementation __615123591 = pNewNode;
							pNewNode = null;
							pNode.m_pChild[nDirection] = __615123591;
						}
						pTempNode.m_pParent = pNode;
						while (true)
						{
							if (pTempNode.m_pParent == null)
							{
								pTempNode.m_eColor = RedBlackNode.Color.COLOR_BLACK;
								{
									return true;
								}
							}
							if (pTempNode.m_pParent.m_eColor == RedBlackNode.Color.COLOR_BLACK)
							{
								{
									return true;
								}
							}
							RedBlackNodeImplementation pUncle = pTempNode.GetUncle();
							if (pUncle != null && pUncle.m_eColor == RedBlackNode.Color.COLOR_RED)
							{
								pTempNode.m_pParent.m_eColor = RedBlackNode.Color.COLOR_BLACK;
								pUncle.m_eColor = RedBlackNode.Color.COLOR_BLACK;
								pTempNode = pTempNode.GetGrandparent();
								pTempNode.m_eColor = RedBlackNode.Color.COLOR_RED;
								continue;
							}
							RedBlackNodeImplementation pGrandparent = pTempNode.GetGrandparent();
							if ((pTempNode == pTempNode.m_pParent.m_pChild[1]) && (pTempNode.m_pParent == pGrandparent.m_pChild[0]))
							{
								m_pImpl.Rotate(pTempNode.m_pParent, 0);
								pTempNode = pTempNode.m_pChild[0];
							}
							else if ((pTempNode == pTempNode.m_pParent.m_pChild[0]) && (pTempNode.m_pParent == pGrandparent.m_pChild[1]))
							{
								m_pImpl.Rotate(pTempNode.m_pParent, 1);
								pTempNode = pTempNode.m_pChild[1];
							}
							pUncle = pTempNode.GetUncle();
							pGrandparent = pTempNode.GetGrandparent();
							pTempNode.m_pParent.m_eColor = RedBlackNode.Color.COLOR_BLACK;
							pGrandparent.m_eColor = RedBlackNode.Color.COLOR_RED;
							if (pTempNode == pTempNode.m_pParent.m_pChild[0])
								m_pImpl.Rotate(pGrandparent, 1);
							else
								m_pImpl.Rotate(pGrandparent, 0);
							{
								return true;
							}
						}
					}
					pNode = pNode.m_pChild[nDirection];
				}
			}

			public bool DeleteObject(object pObject)
			{
				nbAssert.Assert(false);
				return false;
			}

			public RedBlackNode GetRootNode()
			{
				return m_pImpl.m_pRootNode;
			}

			public RedBlackNode GetNode(object pObject)
			{
				if (pObject == null)
					return null;
				RedBlackNodeImplementation pNode = m_pImpl.m_pRootNode;
				while (pNode != null)
				{
					int nComparison = m_pImpl.m_pComparisonCallback(pNode.GetObject(), pObject);
					nbAssert.Assert(nComparison >= -1);
					nbAssert.Assert(nComparison <= 1);
					if (nComparison == 0)
						return pNode;
					int nDirection = 0;
					if (nComparison > 0)
						nDirection = nComparison;
					pNode = pNode.m_pChild[nDirection];
				}
				return null;
			}

		}
		class RedBlackTreeImplementation
		{
			public RedBlackNodeImplementation m_pRootNode;
			public RedBlackTree.ComparisonCallback m_pComparisonCallback;
			public void RecursiveDelete(RedBlackNodeImplementation pNode)
			{
				RedBlackNodeImplementation pTempNode = pNode;
				if (pTempNode != null)
				{
					{
						NumberDuck.Secret.RedBlackNodeImplementation __3820686225 = pTempNode.m_pChild[0];
						pTempNode.m_pChild[0] = null;
						RecursiveDelete(__3820686225);
					}
					pTempNode.m_pChild[0] = null;
					{
						NumberDuck.Secret.RedBlackNodeImplementation __3065717485 = pTempNode.m_pChild[1];
						pTempNode.m_pChild[1] = null;
						RecursiveDelete(__3065717485);
					}
					pTempNode.m_pChild[1] = null;
					{
						pTempNode = null;
					}
				}
			}

			public void Rotate(RedBlackNodeImplementation pNode, int nDirection)
			{
				int nOpposite = 1;
				if (nDirection > 0)
					nOpposite = 0;
				nbAssert.Assert(pNode.m_pChild[nOpposite] != null);
				RedBlackNodeImplementation pChild;
				{
					NumberDuck.Secret.RedBlackNodeImplementation __474790283 = pNode.m_pChild[nOpposite];
					pNode.m_pChild[nOpposite] = null;
					pChild = __474790283;
				}
				{
					NumberDuck.Secret.RedBlackNodeImplementation __2982843613 = pChild.m_pChild[nDirection];
					pChild.m_pChild[nDirection] = null;
					pNode.m_pChild[nOpposite] = __2982843613;
				}
				if (pNode.m_pChild[nOpposite] != null)
					pNode.m_pChild[nOpposite].m_pParent = pNode;
				RedBlackNodeImplementation pParent = pNode.m_pParent;
				pChild.m_pParent = pParent;
				pNode.m_pParent = pChild;
				if (pParent == null)
				{
					nbAssert.Assert(pNode == m_pRootNode);
					{
						NumberDuck.Secret.RedBlackNodeImplementation __1798688362 = m_pRootNode;
						m_pRootNode = null;
						pChild.m_pChild[nDirection] = __1798688362;
					}
					{
						NumberDuck.Secret.RedBlackNodeImplementation __4076228335 = pChild;
						pChild = null;
						m_pRootNode = __4076228335;
					}
				}
				else
				{
					if (pParent.m_pChild[0] == pNode)
					{
						{
							NumberDuck.Secret.RedBlackNodeImplementation __3633264354 = pParent.m_pChild[0];
							pParent.m_pChild[0] = null;
							pChild.m_pChild[nDirection] = __3633264354;
						}
						{
							NumberDuck.Secret.RedBlackNodeImplementation __4076228335 = pChild;
							pChild = null;
							pParent.m_pChild[0] = __4076228335;
						}
					}
					else
					{
						{
							NumberDuck.Secret.RedBlackNodeImplementation __2374967198 = pParent.m_pChild[1];
							pParent.m_pChild[1] = null;
							pChild.m_pChild[nDirection] = __2374967198;
						}
						{
							NumberDuck.Secret.RedBlackNodeImplementation __4076228335 = pChild;
							pChild = null;
							pParent.m_pChild[1] = __4076228335;
						}
					}
				}
			}

			~RedBlackTreeImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class TableElement<T> where T : class
		{
			public int m_nColumn;
			public int m_nRow;
			public T m_xObject = null;
			~TableElement()
			{
				if (m_xObject != null)
					{
						m_xObject = null;
					}
			}

		}
		class Table<T> where T : class
		{
			protected OwnedVector<TableElement<T>> m_pElementVector;
			public Table()
			{
				m_pElementVector = new OwnedVector<TableElement<T>>();
			}

			public void Set(int nColumn, int nRow, T xObject)
			{
				TableElement<T> pElement = GetOrCreate(nColumn, nRow);
				pElement.m_xObject = xObject;
			}

			public int GetIndex(int nColumn, int nRow)
			{
				int i = 0;
				while (true)
				{
					if (i >= m_pElementVector.GetSize())
						return i;
					TableElement<T> pElement = m_pElementVector.Get(i);
					if (pElement.m_nRow > nRow)
						return i;
					if (pElement.m_nRow == nRow)
						if (pElement.m_nColumn >= nColumn)
							return i;
					i++;
				}
			}

			public TableElement<T> Get(int nColumn, int nRow)
			{
				int nIndex = GetIndex(nColumn, nRow);
				TableElement<T> pElement = GetByIndex(nIndex);
				if (pElement == null || pElement.m_nColumn != nColumn || pElement.m_nRow != nRow)
					return null;
				return pElement;
			}

			public TableElement<T> GetByIndex(int nIndex)
			{
				if (nIndex >= m_pElementVector.GetSize())
					return null;
				return m_pElementVector.Get(nIndex);
			}

			public TableElement<T> GetOrCreate(int nColumn, int nRow)
			{
				int nIndex = GetIndex(nColumn, nRow);
				TableElement<T> pElement = GetByIndex(nIndex);
				if (pElement == null || pElement.m_nColumn != nColumn || pElement.m_nRow != nRow)
				{
					TableElement<T> pOwnedElement = new TableElement<T>();
					pOwnedElement.m_nColumn = nColumn;
					pOwnedElement.m_nRow = nRow;
					pOwnedElement.m_xObject = null;
					pElement = pOwnedElement;
					{
						NumberDuck.Secret.TableElement<T> __828927520 = pOwnedElement;
						pOwnedElement = null;
						m_pElementVector.Insert(nIndex, __828927520);
					}
				}
				return pElement;
			}

			public void Erase(int nIndex)
			{
				m_pElementVector.Erase(nIndex);
			}

			public T PopBack()
			{
				TableElement<T> pElement = m_pElementVector.PopBack();
				T xObject;
				{
					T __3920382863 = pElement.m_xObject;
					pElement.m_xObject = null;
					xObject = __3920382863;
				}
				{
					pElement = null;
				}
				{
					T __1841677085 = xObject;
					xObject = null;
					{
						return __1841677085;
					}
				}
			}

			public int GetSize()
			{
				return m_pElementVector.GetSize();
			}

			~Table()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PngImageInfo
		{
			public int m_nWidth;
			public int m_nHeight;
		}
		class PngLoader
		{
			protected PngImageInfo m_pImageInfo = null;
			public PngImageInfo Load(Blob pBlob)
			{
				if (m_pImageInfo != null)
					{
						m_pImageInfo = null;
					}
				BlobView pBlobView = pBlob.GetBlobView();
				pBlobView.SetOffset(0);
				if (pBlob.GetSize() > 24)
				{
					if (pBlobView.UnpackUint8() == 0x89 && pBlobView.UnpackUint8() == 'P' && pBlobView.UnpackUint8() == 'N' && pBlobView.UnpackUint8() == 'G' && pBlobView.UnpackUint8() == 0x0D && pBlobView.UnpackUint8() == 0x0A && pBlobView.UnpackUint8() == 0x1A && pBlobView.UnpackUint8() == 0xA)
					{
						byte n0;
						byte n1;
						byte n2;
						byte n3;
						int nChunkSize;
						n0 = pBlobView.UnpackUint8();
						n1 = pBlobView.UnpackUint8();
						n2 = pBlobView.UnpackUint8();
						n3 = pBlobView.UnpackUint8();
						nChunkSize = (n0 << 24) | (n1 << 16) | (n2 << 8) | (n3 << 0);
						if (pBlobView.GetOffset() + nChunkSize < pBlobView.GetEnd() && pBlobView.UnpackUint8() == 'I' && pBlobView.UnpackUint8() == 'H' && pBlobView.UnpackUint8() == 'D' && pBlobView.UnpackUint8() == 'R')
						{
							m_pImageInfo = new PngImageInfo();
							n0 = pBlobView.UnpackUint8();
							n1 = pBlobView.UnpackUint8();
							n2 = pBlobView.UnpackUint8();
							n3 = pBlobView.UnpackUint8();
							m_pImageInfo.m_nWidth = ((int)(n0)) << 24 | ((int)(n1)) << 16 | ((int)(n2)) << 8 | ((int)(n3)) << 0;
							n0 = pBlobView.UnpackUint8();
							n1 = pBlobView.UnpackUint8();
							n2 = pBlobView.UnpackUint8();
							n3 = pBlobView.UnpackUint8();
							m_pImageInfo.m_nHeight = (n0 << 24) | (n1 << 16) | (n2 << 8) | (n3 << 0);
						}
					}
				}
				return m_pImageInfo;
			}

			~PngLoader()
			{
			}

		}
	}
}

namespace NumberDuck
{
	class MD4
	{
		protected const int BLOCK_SIZE = 64;
		protected uint[] m_nBuffer = new uint[4];
		public void Process(BlobView pBlobView)
		{
			int nIncomingSize = pBlobView.GetSize() - pBlobView.GetOffset();
			Reset();
			while (pBlobView.GetOffset() + BLOCK_SIZE <= pBlobView.GetSize())
			{
				ProcessChunk(pBlobView);
			}
			int nRemainingBlockSize = nIncomingSize % BLOCK_SIZE;
			int nPaddingSize = (BLOCK_SIZE - 8) - nRemainingBlockSize;
			if (nPaddingSize <= 0)
				nPaddingSize += BLOCK_SIZE;
			Blob pFinalBlob = new Blob(true);
			BlobView pFinalBlobView = pFinalBlob.GetBlobView();
			pFinalBlobView.Pack(pBlobView, pBlobView.GetSize() - pBlobView.GetOffset());
			pFinalBlobView.PackUint8(0x80);
			for (int i = 0; i < nPaddingSize - 1; i++)
				pFinalBlobView.PackUint8(0);
			for (int i = 0; i < 4; i++)
				pFinalBlobView.PackUint8((byte)((nIncomingSize * 8) >> (8 * i)));
			pFinalBlobView.PackUint32(0);
			pFinalBlobView.SetOffset(0);
			while (pFinalBlobView.GetOffset() < pFinalBlobView.GetSize())
			{
				ProcessChunk(pFinalBlobView);
			}
		}

		public void BlobWrite(BlobView pBlobView)
		{
			for (int i = 0; i < 4; i++)
				for (int j = 0; j < 4; j++)
					pBlobView.PackUint8((byte)(m_nBuffer[i] >> (8 * j)));
		}

		protected void Reset()
		{
			m_nBuffer[0] = 0x67452301;
			m_nBuffer[1] = 0xEFCDAB89;
			m_nBuffer[2] = 0x98BADCFE;
			m_nBuffer[3] = 0x10325476;
		}

		protected void ProcessChunk(BlobView pBlobView)
		{
			int i;
			uint A = m_nBuffer[0];
			uint B = m_nBuffer[1];
			uint C = m_nBuffer[2];
			uint D = m_nBuffer[3];
			uint[] m_nChunk = new uint[BLOCK_SIZE >> 2];
			Secret.nbAssert.Assert(pBlobView.GetOffset() + BLOCK_SIZE <= pBlobView.GetSize());
			for (i = 0; i < 16; i++)
			{
				uint c0 = pBlobView.UnpackUint8();
				uint c1 = pBlobView.UnpackUint8();
				uint c2 = pBlobView.UnpackUint8();
				uint c3 = pBlobView.UnpackUint8();
				m_nChunk[i] = c0 | (c1 << 8) | (c2 << 16) | (c3 << 24);
			}
			A = FF(A, B, C, D, m_nChunk[0], 3);
			D = FF(D, A, B, C, m_nChunk[1], 7);
			C = FF(C, D, A, B, m_nChunk[2], 11);
			B = FF(B, C, D, A, m_nChunk[3], 19);
			A = FF(A, B, C, D, m_nChunk[4], 3);
			D = FF(D, A, B, C, m_nChunk[5], 7);
			C = FF(C, D, A, B, m_nChunk[6], 11);
			B = FF(B, C, D, A, m_nChunk[7], 19);
			A = FF(A, B, C, D, m_nChunk[8], 3);
			D = FF(D, A, B, C, m_nChunk[9], 7);
			C = FF(C, D, A, B, m_nChunk[10], 11);
			B = FF(B, C, D, A, m_nChunk[11], 19);
			A = FF(A, B, C, D, m_nChunk[12], 3);
			D = FF(D, A, B, C, m_nChunk[13], 7);
			C = FF(C, D, A, B, m_nChunk[14], 11);
			B = FF(B, C, D, A, m_nChunk[15], 19);
			A = GG(A, B, C, D, m_nChunk[0], 3);
			D = GG(D, A, B, C, m_nChunk[4], 5);
			C = GG(C, D, A, B, m_nChunk[8], 9);
			B = GG(B, C, D, A, m_nChunk[12], 13);
			A = GG(A, B, C, D, m_nChunk[1], 3);
			D = GG(D, A, B, C, m_nChunk[5], 5);
			C = GG(C, D, A, B, m_nChunk[9], 9);
			B = GG(B, C, D, A, m_nChunk[13], 13);
			A = GG(A, B, C, D, m_nChunk[2], 3);
			D = GG(D, A, B, C, m_nChunk[6], 5);
			C = GG(C, D, A, B, m_nChunk[10], 9);
			B = GG(B, C, D, A, m_nChunk[14], 13);
			A = GG(A, B, C, D, m_nChunk[3], 3);
			D = GG(D, A, B, C, m_nChunk[7], 5);
			C = GG(C, D, A, B, m_nChunk[11], 9);
			B = GG(B, C, D, A, m_nChunk[15], 13);
			A = HH(A, B, C, D, m_nChunk[0], 3);
			D = HH(D, A, B, C, m_nChunk[8], 9);
			C = HH(C, D, A, B, m_nChunk[4], 11);
			B = HH(B, C, D, A, m_nChunk[12], 15);
			A = HH(A, B, C, D, m_nChunk[2], 3);
			D = HH(D, A, B, C, m_nChunk[10], 9);
			C = HH(C, D, A, B, m_nChunk[6], 11);
			B = HH(B, C, D, A, m_nChunk[14], 15);
			A = HH(A, B, C, D, m_nChunk[1], 3);
			D = HH(D, A, B, C, m_nChunk[9], 9);
			C = HH(C, D, A, B, m_nChunk[5], 11);
			B = HH(B, C, D, A, m_nChunk[13], 15);
			A = HH(A, B, C, D, m_nChunk[3], 3);
			D = HH(D, A, B, C, m_nChunk[11], 9);
			C = HH(C, D, A, B, m_nChunk[7], 11);
			B = HH(B, C, D, A, m_nChunk[15], 15);
			m_nBuffer[0] += A;
			m_nBuffer[1] += B;
			m_nBuffer[2] += C;
			m_nBuffer[3] += D;
		}

		protected uint FF(uint a, uint b, uint c, uint d, uint x, int s)
		{
			uint t = a + ((b & c) | (~b & d)) + x;
			return t << s | t >> (32 - s);
		}

		protected uint GG(uint a, uint b, uint c, uint d, uint x, int s)
		{
			uint t = a + ((b & (c | d)) | (c & d)) + x + 0x5A827999;
			return t << s | t >> (32 - s);
		}

		protected uint HH(uint a, uint b, uint c, uint d, uint x, int s)
		{
			uint t = a + (b ^ c ^ d) + x + 0x6ED9EBA1;
			return t << s | t >> (32 - s);
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class JpegImageInfo
		{
			public int m_nWidth;
			public int m_nHeight;
		}
		class JpegLoader
		{
			protected JpegImageInfo m_pImageInfo = null;
			public JpegImageInfo Load(Blob pBlob)
			{
				if (m_pImageInfo != null)
					{
						m_pImageInfo = null;
					}
				BlobView pBlobView = pBlob.GetBlobView();
				pBlobView.SetOffset(0);
				if (pBlob.GetSize() > 2)
				{
					byte cMarker;
					byte cType;
					cMarker = pBlobView.UnpackUint8();
					cType = pBlobView.UnpackUint8();
					if (cMarker == 0xFF && cType == 0xD8)
					{
						while (pBlobView.GetOffset() < pBlobView.GetEnd())
						{
							int nStart;
							int nSize;
							int cSize0;
							int cSize1;
							if (pBlobView.GetOffset() + 4 > pBlobView.GetEnd())
								break;
							cMarker = pBlobView.UnpackUint8();
							cType = pBlobView.UnpackUint8();
							nStart = pBlobView.GetOffset();
							cSize0 = pBlobView.UnpackUint8();
							cSize1 = pBlobView.UnpackUint8();
							nSize = cSize0 << 8 | cSize1;
							if (cType == 0xC0 || cType == 0xC2)
							{
								if (pBlobView.GetOffset() + 5 > pBlobView.GetEnd())
									break;
								pBlobView.UnpackUint8();
								byte nHeight0 = pBlobView.UnpackUint8();
								byte nHeight1 = pBlobView.UnpackUint8();
								byte nWidth0 = pBlobView.UnpackUint8();
								byte nWidth1 = pBlobView.UnpackUint8();
								m_pImageInfo = new JpegImageInfo();
								m_pImageInfo.m_nHeight = (int)(((uint)(nHeight0)) << 8 | ((uint)(nHeight1)));
								m_pImageInfo.m_nWidth = (int)(((uint)(nWidth0)) << 8 | ((uint)(nWidth1)));
								break;
							}
							else if (cType == 0xDA)
							{
								break;
							}
							pBlobView.SetOffset(nStart + nSize);
						}
					}
				}
				return m_pImageInfo;
			}

			~JpegLoader()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XlsxWorksheet : Worksheet
		{
			public XlsxWorksheet(Workbook pWorkbook) : base(pWorkbook)
			{
			}

			public bool Parse(XlsxWorkbookGlobals pWorkbookGlobals, XmlNode pWorksheetNode)
			{
				double dDefaultRowHeight = -1.0;
				{
					XmlNode pSheetFormatPrElement = pWorksheetNode.GetFirstChildElement("sheetFormatPr");
					if (pSheetFormatPrElement != null)
					{
						string szDefaultRowHeight = pSheetFormatPrElement.GetAttribute("defaultRowHeight");
						if (szDefaultRowHeight != null)
							dDefaultRowHeight = ExternalString.atof(szDefaultRowHeight);
					}
				}
				{
					XmlNode pColsNode = pWorksheetNode.GetFirstChildElement("cols");
					if (pColsNode != null)
					{
						XmlNode pColNode = pColsNode.GetFirstChildElement("col");
						while (pColNode != null)
						{
							string szMin = pColNode.GetAttribute("min");
							if (szMin == null)
								return false;
							ushort nMin = (ushort)(ExternalString.atol(szMin));
							string szMax = pColNode.GetAttribute("max");
							if (szMax == null)
								return false;
							ushort nMax = (ushort)(ExternalString.atol(szMax));
							string szWidth = pColNode.GetAttribute("width");
							if (szWidth == null)
								return false;
							double dWidth = ExternalString.atof(szWidth);
							bool bHidden = false;
							string szHidden = pColNode.GetAttribute("hidden");
							if (szHidden != null && szHidden[0] == '1')
								bHidden = true;
							for (ushort i = nMin; i <= nMax; i++)
							{
								SetColumnWidth((ushort)(i - 1), (ushort)(dWidth * 7.01));
								SetColumnHidden((ushort)(i - 1), bHidden);
							}
							pColNode = pColNode.GetNextSiblingElement("col");
						}
					}
				}
				{
					XmlNode pMergeCellsNode = pWorksheetNode.GetFirstChildElement("mergeCells");
					if (pMergeCellsNode != null)
					{
						XmlNode pMergeCellNode = pMergeCellsNode.GetFirstChildElement("mergeCell");
						while (pMergeCellNode != null)
						{
							string szRef = pMergeCellNode.GetAttribute("ref");
							if (szRef == null)
								return false;
							Secret.Area pArea = Secret.WorksheetImplementation.AddressToArea(szRef);
							CreateMergedCell(pArea.m_pTopLeft.m_nX, pArea.m_pTopLeft.m_nY, (ushort)(pArea.m_pBottomRight.m_nX - pArea.m_pTopLeft.m_nX + 1), (ushort)(pArea.m_pBottomRight.m_nY - pArea.m_pTopLeft.m_nY + 1));
							{
								pArea = null;
							}
							pMergeCellNode = pMergeCellNode.GetNextSiblingElement("mergeCell");
						}
					}
				}
				{
					XmlNode pSheetDataNode = pWorksheetNode.GetFirstChildElement("sheetData");
					if (pSheetDataNode != null)
					{
						XmlNode pRowNode;
						pRowNode = pSheetDataNode.GetFirstChildElement("row");
						while (pRowNode != null)
						{
							string szRow = pRowNode.GetAttribute("r");
							if (szRow == null)
								return false;
							ushort nRow = (ushort)(ExternalString.atol(szRow));
							double dHeight = dDefaultRowHeight;
							string szHeight = pRowNode.GetAttribute("ht");
							if (szHeight != null)
								dHeight = ExternalString.atof(szHeight);
							if (dHeight > 0)
								SetRowHeight((ushort)(nRow - 1), (ushort)(dHeight * 1.334));
							XmlNode pCellNode;
							pCellNode = pRowNode.GetFirstChildElement("c");
							while (pCellNode != null)
							{
								Cell pCell;
								{
									string szReference = pCellNode.GetAttribute("r");
									if (szReference == null)
										return false;
									pCell = GetCellByAddress(szReference);
									if (pCell == null)
										return false;
								}
								{
									string szStyle = pCellNode.GetAttribute("s");
									if (szStyle != null)
									{
										int nStyle = ExternalString.atoi(szStyle);
										Style pStyle = pWorkbookGlobals.GetStyleByIndex((ushort)(nStyle));
										pCell.SetStyle(pStyle);
									}
								}
								string szType = pCellNode.GetAttribute("t");
								if (szType != null)
								{
									if (ExternalString.Equal(szType, "s"))
									{
										XmlNode pValueElement;
										pValueElement = pCellNode.GetFirstChildElement("v");
										uint nIndex = (uint)(ExternalString.atol(pValueElement.GetText()));
										pCell.SetString(pWorkbookGlobals.GetSharedStringByIndex(nIndex));
									}
								}
								pCellNode = pCellNode.GetNextSiblingElement("c");
							}
							pRowNode = pRowNode.GetNextSiblingElement("row");
						}
					}
				}
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class XlsxWorkbookGlobals : WorkbookGlobals
		{
			public static bool LoadVector(Vector<XmlNode> pVector, XmlNode pStyleSheetNode, string szParent, string szChild)
			{
				XmlNode pParentNode = pStyleSheetNode.GetFirstChildElement(szParent);
				if (pParentNode == null)
					return false;
				XmlNode pChildNode = pParentNode.GetFirstChildElement(szChild);
				while (pChildNode != null)
				{
					pVector.PushBack(pChildNode);
					pChildNode = pChildNode.GetNextSiblingElement(szChild);
				}
				return true;
			}

			public static XmlNode GetElement(Vector<XmlNode> pVector, int nIndex)
			{
				if (nIndex < 0 || pVector.GetSize() <= nIndex)
					return null;
				return pVector.Get(nIndex);
			}

			public static uint ParseColor(XmlNode pColorNode)
			{
				string szIndexed = pColorNode.GetAttribute("indexed");
				string szRgb = pColorNode.GetAttribute("rgb");
				if (szIndexed != null)
				{
					int nIndexed = ExternalString.atoi(szIndexed);
					if (nIndexed < BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY)
						return BiffWorkbookGlobals.GetDefaultPaletteColorByIndex((ushort)(nIndexed));
					if (nIndexed == BiffWorkbookGlobals.PALETTE_INDEX_DEFAULT_FOREGROUND)
						return 0x000000;
				}
				else if (szRgb != null)
				{
					InternalString sTemp = new InternalString("0x");
					sTemp.AppendString(szRgb);
					uint nRgb = sTemp.ParseHex();
					{
						sTemp = null;
					}
					{
						return nRgb;
					}
				}
				nbAssert.Assert(false);
				return 0x000000;
			}

			public static bool ApplyFont(XmlNode pFontNode, Style pStyle)
			{
				Font pFont = pStyle.GetFont();
				{
					XmlNode pNameNode = pFontNode.GetFirstChildElement("name");
					if (pNameNode != null)
					{
						string szVal = pNameNode.GetAttribute("val");
						if (szVal == null)
							return false;
						pFont.SetName(szVal);
					}
				}
				{
					XmlNode pSzElement = pFontNode.GetFirstChildElement("sz");
					if (pSzElement != null)
					{
						string szVal = pSzElement.GetAttribute("val");
						if (szVal == null)
							return false;
						double dPoints = (double)(ExternalString.atoi(szVal));
						double dPixels = (dPoints * 96.0 + 72.0 / 2.0) / 72.0;
						pFont.SetSize((byte)(dPixels));
					}
				}
				{
					XmlNode pBElement = pFontNode.GetFirstChildElement("b");
					if (pBElement != null)
					{
						string szVal = pBElement.GetAttribute("val");
						if (szVal == null || szVal[0] == '1')
							pFont.SetBold(true);
						else
							pFont.SetBold(false);
					}
				}
				{
					XmlNode pUElement = pFontNode.GetFirstChildElement("u");
					if (pUElement != null)
					{
						string szVal = pUElement.GetAttribute("val");
						if (szVal == null || ExternalString.Equal(szVal, "single"))
							pFont.SetUnderline(Font.Underline.UNDERLINE_SINGLE);
						else if (ExternalString.Equal(szVal, "double"))
							pFont.SetUnderline(Font.Underline.UNDERLINE_DOUBLE);
						else if (ExternalString.Equal(szVal, "singleAccounting"))
							pFont.SetUnderline(Font.Underline.UNDERLINE_SINGLE_ACCOUNTING);
						else if (ExternalString.Equal(szVal, "doubleAccounting"))
							pFont.SetUnderline(Font.Underline.UNDERLINE_DOUBLE_ACCOUNTING);
					}
				}
				return true;
			}

			public static bool ApplyBorderLine(XmlNode pSideElement, Line pLine)
			{
				string szStyle = pSideElement.GetAttribute("style");
				if (szStyle != null)
				{
					if (ExternalString.Equal(szStyle, "dashDot"))
						pLine.SetType(Line.Type.TYPE_DASH_DOT);
					else if (ExternalString.Equal(szStyle, "dashDotDot"))
						pLine.SetType(Line.Type.TYPE_DASH_DOT_DOT);
					else if (ExternalString.Equal(szStyle, "dashed"))
						pLine.SetType(Line.Type.TYPE_DASHED);
					else if (ExternalString.Equal(szStyle, "dotted"))
						pLine.SetType(Line.Type.TYPE_DOTTED);
					else if (ExternalString.Equal(szStyle, "double"))
						pLine.SetType(Line.Type.TYPE_THICK);
					else if (ExternalString.Equal(szStyle, "hair"))
						pLine.SetType(Line.Type.TYPE_THIN);
					else if (ExternalString.Equal(szStyle, "medium"))
						pLine.SetType(Line.Type.TYPE_MEDIUM);
					else if (ExternalString.Equal(szStyle, "mediumDashDot"))
						pLine.SetType(Line.Type.TYPE_MEDIUM_DASH_DOT);
					else if (ExternalString.Equal(szStyle, "mediumDashDotDot"))
						pLine.SetType(Line.Type.TYPE_MEDIUM_DASH_DOT_DOT);
					else if (ExternalString.Equal(szStyle, "mediumDashed"))
						pLine.SetType(Line.Type.TYPE_MEDIUM_DASHED);
					else if (ExternalString.Equal(szStyle, "none"))
						pLine.SetType(Line.Type.TYPE_NONE);
					else if (ExternalString.Equal(szStyle, "slantDashDot"))
						pLine.SetType(Line.Type.TYPE_DASHED);
					else if (ExternalString.Equal(szStyle, "thick"))
						pLine.SetType(Line.Type.TYPE_THICK);
					else if (ExternalString.Equal(szStyle, "thin"))
						pLine.SetType(Line.Type.TYPE_THIN);
					else
						return false;
				}
				XmlNode pColorElement = pSideElement.GetFirstChildElement("color");
				if (pColorElement != null)
					pLine.GetColor().SetFromRgba(ParseColor(pColorElement));
				return true;
			}

			public static bool ApplyBorder(XmlNode pBorderElement, Style pStyle)
			{
				XmlNode pSideElement;
				pSideElement = pBorderElement.GetFirstChildElement("left");
				if (pSideElement != null)
					if (!ApplyBorderLine(pSideElement, pStyle.GetLeftBorderLine()))
						return false;
				pSideElement = pBorderElement.GetFirstChildElement("right");
				if (pSideElement != null)
					if (!ApplyBorderLine(pSideElement, pStyle.GetRightBorderLine()))
						return false;
				pSideElement = pBorderElement.GetFirstChildElement("top");
				if (pSideElement != null)
					if (!ApplyBorderLine(pSideElement, pStyle.GetTopBorderLine()))
						return false;
				pSideElement = pBorderElement.GetFirstChildElement("bottom");
				if (pSideElement != null)
					if (!ApplyBorderLine(pSideElement, pStyle.GetBottomBorderLine()))
						return false;
				return true;
			}

			public static bool ApplyFill(XmlNode pFillElement, Style pStyle)
			{
				XmlNode pPatternFillElement = pFillElement.GetFirstChildElement("patternFill");
				if (pPatternFillElement != null)
				{
					string szPatternType = pPatternFillElement.GetAttribute("patternType");
					if (szPatternType != null)
					{
					}
					XmlNode pColorElement = pPatternFillElement.GetFirstChildElement("fgColor");
					if (pColorElement != null)
						pStyle.GetBackgroundColor(true).SetFromRgba(ParseColor(pColorElement));
				}
				else
				{
				}
				return true;
			}

			public static bool ParseStyles(WorkbookGlobals pWorkbookGlobals, XmlNode pStyleSheetNode)
			{
				pWorkbookGlobals.m_pStyleVector.Clear();
				Vector<XmlNode> pNumFmtVector = new Vector<XmlNode>();
				Vector<XmlNode> pFontVector = new Vector<XmlNode>();
				Vector<XmlNode> pFillVector = new Vector<XmlNode>();
				Vector<XmlNode> pBorderVector = new Vector<XmlNode>();
				Vector<XmlNode> pCellStyleXfVector = new Vector<XmlNode>();
				Vector<XmlNode> pCellXfVector = new Vector<XmlNode>();
				Vector<XmlNode> pCellStyleVector = new Vector<XmlNode>();
				bool bResult;
				LoadVector(pNumFmtVector, pStyleSheetNode, "numFmts", "numFmt");
				LoadVector(pFontVector, pStyleSheetNode, "fonts", "font");
				LoadVector(pFillVector, pStyleSheetNode, "fills", "fill");
				LoadVector(pBorderVector, pStyleSheetNode, "borders", "border");
				LoadVector(pCellStyleXfVector, pStyleSheetNode, "cellStyleXfs", "xf");
				LoadVector(pCellXfVector, pStyleSheetNode, "cellXfs", "xf");
				LoadVector(pCellStyleVector, pStyleSheetNode, "cellStyles", "cellStyle");
				bResult = SubParseStyles(pWorkbookGlobals, pStyleSheetNode, pNumFmtVector, pFontVector, pFillVector, pBorderVector, pCellStyleXfVector, pCellXfVector, pCellStyleVector);
				{
					pNumFmtVector = null;
				}
				{
					pFontVector = null;
				}
				{
					pFillVector = null;
				}
				{
					pBorderVector = null;
				}
				{
					pCellStyleXfVector = null;
				}
				{
					pCellXfVector = null;
				}
				{
					pCellStyleVector = null;
				}
				{
					return bResult;
				}
			}

			public static bool SubParseStyles(WorkbookGlobals pWorkbookGlobals, XmlNode pStyleSheetNode, Vector<XmlNode> pNumFmtVector, Vector<XmlNode> pFontVector, Vector<XmlNode> pFillVector, Vector<XmlNode> pBorderVector, Vector<XmlNode> pCellStyleXfVector, Vector<XmlNode> pCellXfVector, Vector<XmlNode> pCellStyleVector)
			{
				for (int i = 0; i < pCellXfVector.GetSize(); i++)
				{
					Style pStyle = pWorkbookGlobals.CreateStyle();
					XmlNode pXfElement = pCellXfVector.Get(i);
					XmlNode pInheritXfElement = null;
					XmlNode pCellStyleElement = null;
					string szXfId = pXfElement.GetAttribute("xfId");
					if (szXfId != null)
					{
						int nXfId = ExternalString.atoi(szXfId);
						if (nXfId < 0 || pCellStyleXfVector.GetSize() <= nXfId)
							return false;
						pInheritXfElement = pCellStyleXfVector.Get(nXfId);
						if (nXfId < 0 || pCellStyleVector.GetSize() <= nXfId)
							return false;
						pCellStyleElement = pCellStyleVector.Get(nXfId);
					}
					{
						string szFontId = null;
						string szApplyFont = pXfElement.GetAttribute("applyFont");
						if (szApplyFont != null && szApplyFont[0] == '1')
							szFontId = pXfElement.GetAttribute("fontId");
						else
							szFontId = pInheritXfElement.GetAttribute("fontId");
						if (szFontId != null)
						{
							XmlNode pFontElement = GetElement(pFontVector, ExternalString.atoi(szFontId));
							if (pFontElement == null)
								return false;
							if (!ApplyFont(pFontElement, pStyle))
								return false;
						}
					}
					{
						string szBorderId = null;
						string szApplyBorder = pXfElement.GetAttribute("applyBorder");
						if (szApplyBorder != null && szApplyBorder[0] == '1')
							szBorderId = pXfElement.GetAttribute("borderId");
						else
							szBorderId = pInheritXfElement.GetAttribute("borderId");
						if (szBorderId != null)
						{
							XmlNode pBorderElement = GetElement(pBorderVector, ExternalString.atoi(szBorderId));
							if (pBorderElement == null)
								return false;
							if (!ApplyBorder(pBorderElement, pStyle))
								return false;
						}
					}
					{
						string szFillId = null;
						string szApplyFill = pXfElement.GetAttribute("applyFill");
						if (szApplyFill != null && szApplyFill[0] == '1')
							szFillId = pXfElement.GetAttribute("fillId");
						else
							szFillId = pInheritXfElement.GetAttribute("fillId");
						if (szFillId != null)
						{
							XmlNode pFillElement = GetElement(pFillVector, ExternalString.atoi(szFillId));
							if (pFillElement == null)
								return false;
							if (!ApplyFill(pFillElement, pStyle))
								return false;
						}
					}
					{
						XmlNode pAlignmentElement = null;
						string szApplyAlignment = pXfElement.GetAttribute("applyAlignment");
						if (szApplyAlignment != null && szApplyAlignment[0] == '1')
							pAlignmentElement = pXfElement.GetFirstChildElement("alignment");
						else
							pAlignmentElement = pInheritXfElement.GetFirstChildElement("alignment");
						if (pAlignmentElement != null)
						{
							string szHorizontal = pAlignmentElement.GetAttribute("horizontal");
							if (szHorizontal != null)
							{
								if (ExternalString.Equal(szHorizontal, "general"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_GENERAL);
								else if (ExternalString.Equal(szHorizontal, "left"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_LEFT);
								else if (ExternalString.Equal(szHorizontal, "center"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_CENTER);
								else if (ExternalString.Equal(szHorizontal, "right"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_RIGHT);
								else if (ExternalString.Equal(szHorizontal, "fill"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_FILL);
								else if (ExternalString.Equal(szHorizontal, "justify"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_JUSTIFY);
								else if (ExternalString.Equal(szHorizontal, "centerContinuous"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_CENTER_ACROSS_SELECTION);
								else if (ExternalString.Equal(szHorizontal, "distributed"))
									pStyle.SetHorizontalAlign(Style.HorizontalAlign.HORIZONTAL_ALIGN_DISTRIBUTED);
							}
							string szVertical = pAlignmentElement.GetAttribute("vertical");
							if (szVertical != null)
							{
								if (ExternalString.Equal(szVertical, "top"))
									pStyle.SetVerticalAlign(Style.VerticalAlign.VERTICAL_ALIGN_TOP);
								else if (ExternalString.Equal(szVertical, "center"))
									pStyle.SetVerticalAlign(Style.VerticalAlign.VERTICAL_ALIGN_CENTER);
								else if (ExternalString.Equal(szVertical, "bottom"))
									pStyle.SetVerticalAlign(Style.VerticalAlign.VERTICAL_ALIGN_BOTTOM);
								else if (ExternalString.Equal(szVertical, "justify"))
									pStyle.SetVerticalAlign(Style.VerticalAlign.VERTICAL_ALIGN_JUSTIFY);
								else if (ExternalString.Equal(szVertical, "distributed"))
									pStyle.SetVerticalAlign(Style.VerticalAlign.VERTICAL_ALIGN_DISTRIBUTED);
							}
						}
					}
					{
						string szNumFmtId = null;
						string szApplyNumberFormat = pXfElement.GetAttribute("applyNumberFormat");
						if (szApplyNumberFormat != null && szApplyNumberFormat[0] == '1')
							szNumFmtId = pXfElement.GetAttribute("numFmtId");
						else
							szNumFmtId = pInheritXfElement.GetAttribute("numFmtId");
						if (szNumFmtId != null)
						{
							int nNumFmtId = ExternalString.atoi(szNumFmtId);
							bool bFound = false;
							for (int j = 0; j < pNumFmtVector.GetSize(); j++)
							{
								XmlNode pNumberFormatElement = pNumFmtVector.Get(j);
								string szTestNumFmtId = pNumberFormatElement.GetAttribute("numFmtId");
								if (szTestNumFmtId == null)
									return false;
								if (ExternalString.Equal(szTestNumFmtId, szNumFmtId))
								{
									string szFormatCode = pNumberFormatElement.GetAttribute("formatCode");
									if (szFormatCode == null)
										return false;
									pStyle.SetFormat(szFormatCode);
									bFound = true;
									break;
								}
							}
							if (!bFound)
							{
								switch (nNumFmtId)
								{
									case 0:
									{
										pStyle.SetFormat("General");
										break;
									}

									case 1:
									{
										pStyle.SetFormat("0");
										break;
									}

									case 2:
									{
										pStyle.SetFormat("0.00");
										break;
									}

									case 3:
									{
										pStyle.SetFormat("#,##0");
										break;
									}

									case 4:
									{
										pStyle.SetFormat("#,##0.00");
										break;
									}

									case 9:
									{
										pStyle.SetFormat("0%");
										break;
									}

									case 10:
									{
										pStyle.SetFormat("0.00%");
										break;
									}

									case 11:
									{
										pStyle.SetFormat("0.00E+00");
										break;
									}

									case 12:
									{
										pStyle.SetFormat("# \\?/\\?");
										break;
									}

									case 13:
									{
										pStyle.SetFormat("# \\?\\?/\\?\\?");
										break;
									}

									case 14:
									{
										pStyle.SetFormat("mm-dd-yy");
										break;
									}

									case 15:
									{
										pStyle.SetFormat("d-mmm-yy");
										break;
									}

									case 16:
									{
										pStyle.SetFormat("d-mmm");
										break;
									}

									case 17:
									{
										pStyle.SetFormat("mmm-yy");
										break;
									}

									case 18:
									{
										pStyle.SetFormat("h:mm AM/PM");
										break;
									}

									case 19:
									{
										pStyle.SetFormat("h:mm:ss AM/PM");
										break;
									}

									case 20:
									{
										pStyle.SetFormat("h:mm");
										break;
									}

									case 21:
									{
										pStyle.SetFormat("h:mm:ss");
										break;
									}

									case 22:
									{
										pStyle.SetFormat("m/d/yy h:mm");
										break;
									}

									case 37:
									{
										pStyle.SetFormat("#,##0 ;(#,##0)");
										break;
									}

									case 38:
									{
										pStyle.SetFormat("#,##0 ;[Red](#,##0)");
										break;
									}

									case 39:
									{
										pStyle.SetFormat("#,##0.00;-#,##0.00");
										break;
									}

									case 40:
									{
										pStyle.SetFormat("#,##0.00;[Red](#,##0.00)");
										break;
									}

									case 45:
									{
										pStyle.SetFormat("mm:ss");
										break;
									}

									case 46:
									{
										pStyle.SetFormat("[h]:mm:ss");
										break;
									}

									case 47:
									{
										pStyle.SetFormat("mmss.0");
										break;
									}

									case 48:
									{
										pStyle.SetFormat("##0.0E+0");
										break;
									}

									case 49:
									{
										pStyle.SetFormat("@");
										break;
									}

									default:
									{
										return false;
									}

								}
							}
						}
					}
				}
				return true;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ColumnInfo
		{
			public ushort m_nWidth;
			public bool m_bHidden;
		}
		class RowInfo
		{
			public ushort m_nHeight;
		}
		class Coordinate
		{
			public Coordinate()
			{
				m_nX = 0;
				m_nY = 0;
				m_bXRelative = true;
				m_bYRelative = true;
			}

			public Coordinate(ushort nX, ushort nY, bool bXRelative = true, bool bYRelative = true)
			{
				nbAssert.Assert(nY <= Worksheet.MAX_ROW);
				nbAssert.Assert(nX <= Worksheet.MAX_COLUMN);
				m_nX = nX;
				m_nY = nY;
				m_bXRelative = bXRelative;
				m_bYRelative = bYRelative;
			}

			public Coordinate CreateClone()
			{
				return new Coordinate(m_nX, m_nY, m_bXRelative, m_bYRelative);
			}

			public ushort m_nX;
			public ushort m_nY;
			public bool m_bXRelative;
			public bool m_bYRelative;
		}
		class WorksheetRange
		{
			public WorksheetRange(ushort nFirst, ushort nLast)
			{
				m_nFirst = nFirst;
				m_nLast = nLast;
			}

			public WorksheetRange CreateClone()
			{
				return new WorksheetRange(m_nFirst, m_nLast);
			}

			public ushort m_nFirst;
			public ushort m_nLast;
		}
		class Coordinate3d
		{
			public ushort m_nWorksheetFirst;
			public ushort m_nWorksheetLast;
			public Coordinate m_pCoordinate;
			public Coordinate3d(ushort nWorksheetFirst, ushort nWorksheetLast, Coordinate pCoordinate)
			{
				m_nWorksheetFirst = nWorksheetFirst;
				m_nWorksheetLast = nWorksheetLast;
				m_pCoordinate = pCoordinate;
			}

			public Coordinate3d CreateClone()
			{
				return new Coordinate3d(m_nWorksheetFirst, m_nWorksheetLast, m_pCoordinate.CreateClone());
			}

			public void ToString(WorksheetImplementation pWorksheetImplementation, InternalString sOut)
			{
				pWorksheetImplementation.WorksheetRangeToAddress(m_nWorksheetFirst, m_nWorksheetLast, sOut);
				sOut.AppendChar('!');
				WorksheetImplementation.CoordinateToAddress(m_pCoordinate, sOut);
			}

			~Coordinate3d()
			{
			}

		}
		class Area
		{
			public Coordinate m_pTopLeft;
			public Coordinate m_pBottomRight;
			public Area(Coordinate pTopLeft, Coordinate pBottomRight)
			{
				m_pTopLeft = pTopLeft;
				m_pBottomRight = pBottomRight;
			}

			public Area CreateClone()
			{
				return new Area(m_pTopLeft.CreateClone(), m_pBottomRight.CreateClone());
			}

			~Area()
			{
			}

		}
		class Area3d
		{
			public WorksheetRange m_pWorksheetRange;
			public Area m_pArea;
			public Area3d(ushort nWorksheetFirst, ushort nWorksheetLast, Area pArea)
			{
				m_pWorksheetRange = new WorksheetRange(nWorksheetFirst, nWorksheetLast);
				m_pArea = pArea;
			}

			public Area3d CreateClone()
			{
				return new Area3d(m_pWorksheetRange.m_nFirst, m_pWorksheetRange.m_nLast, m_pArea.CreateClone());
			}

			~Area3d()
			{
			}

		}
		class WorksheetImplementation
		{
			public Workbook m_pWorkbook;
			public InternalString m_sName;
			public Worksheet.Orientation m_eOrientation;
			public bool m_bShowGridlines;
			public bool m_bPrintGridlines;
			public Table<Cell> m_pCellTable;
			public Table<ColumnInfo> m_pColumnInfoTable;
			public Table<RowInfo> m_pRowInfoTable;
			public ushort m_nDefaultRowHeight;
			public OwnedVector<Picture> m_pPictureVector;
			public OwnedVector<Chart> m_pChartVector;
			public OwnedVector<MergedCell> m_pMergedCellVector;
			public Worksheet m_pWorksheet;
			public WorksheetImplementation(Workbook pWorkbook, Worksheet pWorksheet)
			{
				m_sName = new Secret.InternalString("");
				m_pWorkbook = pWorkbook;
				m_pWorksheet = pWorksheet;
				m_eOrientation = Worksheet.Orientation.ORIENTATION_PORTRAIT;
				m_bPrintGridlines = false;
				m_bShowGridlines = true;
				m_pCellTable = new Table<Cell>();
				m_pColumnInfoTable = new Table<ColumnInfo>();
				m_pRowInfoTable = new Table<RowInfo>();
				m_nDefaultRowHeight = 255;
				m_pPictureVector = new OwnedVector<Picture>();
				m_pChartVector = new OwnedVector<Chart>();
				m_pMergedCellVector = new OwnedVector<MergedCell>();
			}

			~WorksheetImplementation()
			{
				{
					m_pPictureVector = null;
				}
			}

			public static ushort TwipsToPixels(ushort nTwips)
			{
				return (ushort)(((uint)(nTwips) * 546 + 8190 / 2) / 8190);
			}

			public static ushort PixelsToTwips(ushort nPixels)
			{
				return (ushort)((uint)(nPixels) * 8190 / 546);
			}

			public static void CoordinateToAddress(Coordinate pCoordinate, InternalString sOut)
			{
				if (!pCoordinate.m_bXRelative)
					sOut.AppendChar('$');
				if (pCoordinate.m_nX > 26)
					sOut.AppendChar((char)('A' + (pCoordinate.m_nX / 26 - 1)));
				sOut.AppendChar((char)('A' + pCoordinate.m_nX % 26));
				if (!pCoordinate.m_bYRelative)
					sOut.AppendChar('$');
				sOut.AppendUint32((uint)(pCoordinate.m_nY) + 1);
			}

			public static void AreaToAddress(Area pArea, InternalString sOut)
			{
				CoordinateToAddress(pArea.m_pTopLeft, sOut);
				sOut.AppendChar(':');
				CoordinateToAddress(pArea.m_pBottomRight, sOut);
			}

			public static Coordinate AddressToCoordinate(string szAddress)
			{
				ushort nX = 0;
				bool bXRelative = true;
				ushort nY = 0;
				bool bYRelative = true;
				ushort nMultiplier;
				ushort nBase;
				int nIndex = 0;
				Vector<int> nTempVector = new Vector<int>();
				InternalString sAddress = new InternalString(szAddress);
				if (sAddress.GetChar(nIndex) == '$')
				{
					bXRelative = false;
					nIndex++;
				}
				while (nIndex < sAddress.GetLength())
				{
					char cChar = sAddress.GetChar(nIndex);
					if (cChar >= 'A' && cChar <= 'Z')
						nTempVector.PushBack(cChar - 'A');
					else if (cChar >= 'a' && cChar <= 'z')
						nTempVector.PushBack(cChar - 'a');
					else
						break;
					nIndex++;
				}
				if (nTempVector.GetSize() == 0)
				{
					{
						nTempVector = null;
					}
					{
						sAddress = null;
					}
					{
						return null;
					}
				}
				nMultiplier = 1;
				nBase = 26;
				for (int i = 0; i < nTempVector.GetSize(); i++)
				{
					nX = (ushort)(nX + (nTempVector.Get(nTempVector.GetSize() - i - 1) + 1) * nMultiplier);
					nMultiplier = (ushort)(nMultiplier * nBase);
				}
				nX--;
				if (sAddress.GetChar(nIndex) == '$')
				{
					bYRelative = false;
					nIndex++;
				}
				nTempVector.Clear();
				while (nIndex < sAddress.GetLength())
				{
					char cChar = sAddress.GetChar(nIndex);
					if (cChar >= '0' && cChar <= '9')
						nTempVector.PushBack(cChar - '0');
					else
						break;
					nIndex++;
				}
				if (nTempVector.GetSize() == 0)
				{
					{
						return null;
					}
				}
				if (nTempVector.GetSize() == 1 && nTempVector.Get(0) == '0')
				{
					{
						return null;
					}
				}
				nMultiplier = 1;
				nBase = 10;
				for (int i = 0; i < nTempVector.GetSize(); i++)
				{
					nY = (ushort)(nY + nTempVector.Get(nTempVector.GetSize() - i - 1) * nMultiplier);
					nMultiplier = (ushort)(nMultiplier * nBase);
				}
				nY--;
				if (sAddress.GetLength() != nIndex)
				{
					{
						return null;
					}
				}
				if (nX > 255 || nY > 65535)
				{
					return null;
				}
				{
					return new Coordinate(nX, nY, bXRelative, bYRelative);
				}
			}

			public static Area AddressToArea(string szAddress)
			{
				Area pArea = null;
				InternalString sAddress = new InternalString(szAddress);
				int nIndex = sAddress.FindChar(':');
				if (nIndex >= 0)
				{
					InternalString sFirst = new InternalString(szAddress);
					sFirst.SubStr(0, nIndex);
					InternalString sSecond = new InternalString(szAddress);
					sSecond.SubStr(nIndex + 1, sSecond.GetLength() - (nIndex + 1));
					Coordinate pFirst = AddressToCoordinate(sFirst.GetExternalString());
					Coordinate pSecond = AddressToCoordinate(sSecond.GetExternalString());
					if (pFirst != null && pSecond != null)
					{
						{
							NumberDuck.Secret.Coordinate __1625950533 = pFirst;
							pFirst = null;
							{
								NumberDuck.Secret.Coordinate __3967857369 = pSecond;
								pSecond = null;
								pArea = new Area(__1625950533, __3967857369);
							}
						}
					}
				}
				{
					NumberDuck.Secret.Area __4245081970 = pArea;
					pArea = null;
					{
						return __4245081970;
					}
				}
			}

			public void WorksheetRangeToAddress(ushort nFirst, ushort nLast, InternalString sOut)
			{
				nbAssert.Assert(nFirst >= 0 && nFirst <= m_pWorkbook.GetNumWorksheet());
				nbAssert.Assert(nLast >= nFirst && nLast <= m_pWorkbook.GetNumWorksheet());
				InternalString sWorksheetFirst = m_pWorkbook.GetWorksheetByIndex(nFirst).m_pImpl.m_sName;
				InternalString sWorksheetLast = m_pWorkbook.GetWorksheetByIndex(nLast).m_pImpl.m_sName;
				bool bHasSpace = sWorksheetFirst.FindChar(' ') != -1 || (nFirst != nLast && sWorksheetLast.FindChar(' ') != -1);
				if (bHasSpace)
					sOut.AppendChar('\'');
				sOut.AppendString(sWorksheetFirst.GetExternalString());
				if (nFirst != nLast)
				{
					sOut.AppendChar(':');
					sOut.AppendString(sWorksheetLast.GetExternalString());
				}
				if (bHasSpace)
					sOut.AppendChar('\'');
			}

			public void Area3dToAddress(Area3d pArea3d, InternalString sOut)
			{
				WorksheetRangeToAddress(pArea3d.m_pWorksheetRange.m_nFirst, pArea3d.m_pWorksheetRange.m_nLast, sOut);
				sOut.AppendChar('!');
				AreaToAddress(pArea3d.m_pArea, sOut);
			}

			public WorksheetRange ParseWorksheetRange(InternalString sString)
			{
				if (sString.GetLength() > 2)
				{
					InternalString sFirst = sString.CreateClone();
					InternalString sLast;
					char nFirstChar = sFirst.GetChar(0);
					char nLastChar = sFirst.GetChar(sFirst.GetLength() - 1);
					if (nFirstChar == '\'' && nLastChar == '\'')
						sFirst.SubStr(1, sFirst.GetLength() - 2);
					sLast = sFirst.CreateClone();
					int nIndex = sFirst.FindChar(':');
					if (nIndex != -1)
					{
						sFirst.SubStr(0, nIndex);
						sLast.SubStr(nIndex + 1, sLast.GetLength() - (nIndex + 1));
					}
					int nWorksheetFirst = -1;
					int nWorksheetLast = -1;
					Workbook pWorkbook = GetWorkbook();
					for (ushort i = 0; i < pWorkbook.GetNumWorksheet(); i++)
					{
						string szName = pWorkbook.GetWorksheetByIndex(i).GetName();
						if (nWorksheetFirst == -1 && ExternalString.Equal(sFirst.GetExternalString(), szName))
							nWorksheetFirst = i;
						if (nWorksheetLast == -1 && ExternalString.Equal(sLast.GetExternalString(), szName))
							nWorksheetLast = i;
					}
					if (nWorksheetFirst >= 0 && nWorksheetLast >= 0)
					{
						return new WorksheetRange((ushort)(nWorksheetFirst), (ushort)(nWorksheetLast));
					}
				}
				return null;
			}

			public Coordinate3d ParseCoordinate3d(InternalString sString)
			{
				int nIndex = sString.FindChar('!');
				if (nIndex >= 0)
				{
					InternalString sWorksheetRange = sString.CreateClone();
					sWorksheetRange.SubStr(0, nIndex);
					InternalString sCell = sString.CreateClone();
					sCell.SubStr(nIndex + 1, sCell.GetLength() - (nIndex + 1));
					WorksheetRange pWorksheetRange = ParseWorksheetRange(sWorksheetRange);
					if (pWorksheetRange != null)
					{
						Coordinate pCoordinate = AddressToCoordinate(sCell.GetExternalString());
						if (pCoordinate != null)
						{
							Coordinate3d pCoordinate3d;
							{
								NumberDuck.Secret.Coordinate __3642692973 = pCoordinate;
								pCoordinate = null;
								pCoordinate3d = new Coordinate3d(pWorksheetRange.m_nFirst, pWorksheetRange.m_nLast, __3642692973);
							}
							{
								NumberDuck.Secret.Coordinate3d __1094936853 = pCoordinate3d;
								pCoordinate3d = null;
								{
									return __1094936853;
								}
							}
						}
						{
							pWorksheetRange = null;
						}
					}
					{
						sCell = null;
					}
					{
						sWorksheetRange = null;
					}
				}
				return null;
			}

			public Area3d ParseArea3d(InternalString sString)
			{
				int nIndex = sString.FindChar('!');
				if (nIndex >= 0)
				{
					InternalString sWorksheetRange = sString.CreateClone();
					sWorksheetRange.SubStr(0, nIndex);
					InternalString sArea = sString.CreateClone();
					sArea.SubStr(nIndex + 1, sArea.GetLength() - (nIndex + 1));
					WorksheetRange pWorksheetRange = ParseWorksheetRange(sWorksheetRange);
					if (pWorksheetRange != null)
					{
						Area pArea = AddressToArea(sArea.GetExternalString());
						if (pArea != null)
						{
							Area3d pArea3d;
							{
								NumberDuck.Secret.Area __4245081970 = pArea;
								pArea = null;
								pArea3d = new Area3d(pWorksheetRange.m_nFirst, pWorksheetRange.m_nLast, __4245081970);
							}
							{
								NumberDuck.Secret.Area3d __2738670685 = pArea3d;
								pArea3d = null;
								{
									return __2738670685;
								}
							}
						}
					}
				}
				return null;
			}

			public ColumnInfo GetColumnInfo(ushort nColumn)
			{
				TableElement<ColumnInfo> pElement = m_pColumnInfoTable.Get(nColumn, 0);
				if (pElement != null)
					return pElement.m_xObject;
				return null;
			}

			public ColumnInfo GetOrCreateColumnInfo(ushort nColumn)
			{
				TableElement<ColumnInfo> pElement = m_pColumnInfoTable.GetOrCreate(nColumn, 0);
				if (pElement.m_xObject == null)
				{
					ColumnInfo pColumnInfo = new ColumnInfo();
					pColumnInfo.m_bHidden = false;
					pColumnInfo.m_nWidth = Worksheet.DEFAULT_COLUMN_WIDTH;
					{
						NumberDuck.Secret.ColumnInfo __1173438266 = pColumnInfo;
						pColumnInfo = null;
						pElement.m_xObject = __1173438266;
					}
				}
				return pElement.m_xObject;
			}

			public RowInfo GetRowInfo(ushort nRow)
			{
				TableElement<RowInfo> pElement = m_pRowInfoTable.Get(0, nRow);
				if (pElement != null)
					return pElement.m_xObject;
				return null;
			}

			public RowInfo GetOrCreateRowInfo(ushort nRow)
			{
				TableElement<RowInfo> pElement = m_pRowInfoTable.GetOrCreate(0, nRow);
				if (pElement.m_xObject == null)
				{
					RowInfo pRowInfo = new RowInfo();
					pRowInfo.m_nHeight = Worksheet.DEFAULT_ROW_HEIGHT;
					{
						NumberDuck.Secret.RowInfo __3798332131 = pRowInfo;
						pRowInfo = null;
						pElement.m_xObject = __3798332131;
					}
				}
				return pElement.m_xObject;
			}

			public Workbook GetWorkbook()
			{
				return m_pWorkbook;
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class WorkbookImplementation
		{
			public WorkbookGlobals m_pWorkbookGlobals;
			public OwnedVector<Worksheet> m_pWorksheetVector;
			~WorkbookImplementation()
			{
				while (m_pWorksheetVector.GetSize() > 0)
				{
					Worksheet pWorksheet = m_pWorksheetVector.Remove(0);
					{
						pWorksheet = null;
					}
				}
			}

		}
	}
}

namespace NumberDuck
{
	class Workbook
	{
		public enum License
		{
			AGPL,
			Commercial,
		}

		public enum FileType
		{
			XLS,
			XLSX,
		}

		public Secret.WorkbookImplementation m_pImpl;
		public Workbook(License eLicense)
		{
			m_pImpl = new Secret.WorkbookImplementation();
			m_pImpl.m_pWorkbookGlobals = null;
			m_pImpl.m_pWorksheetVector = new Secret.OwnedVector<Worksheet>();
			Clear();
		}

		public void Clear()
		{
			m_pImpl.m_pWorksheetVector.Clear();
			if (m_pImpl.m_pWorkbookGlobals != null)
				{
				}
			m_pImpl.m_pWorkbookGlobals = new Secret.WorkbookGlobals();
			CreateWorksheet();
		}

		public bool Load(string szFileName)
		{
			m_pImpl.m_pWorksheetVector.Clear();
			if (m_pImpl.m_pWorkbookGlobals != null)
			{
				{
				}
				m_pImpl.m_pWorkbookGlobals = null;
			}
			bool bLoaded = false;
			Secret.CompoundFile pCompoundFile = new Secret.CompoundFile();
			if (pCompoundFile.Load(szFileName))
			{
				Secret.Stream pStream = pCompoundFile.GetStreamByName("Workbook");
				if (pStream != null)
				{
					pStream.SetOffset(0);
					while (pStream.GetOffset() < pStream.GetStreamSize())
					{
						Secret.BiffRecord pBiffRecord = Secret.BiffRecord.CreateBiffRecord(pStream);
						if (pBiffRecord.GetType() == Secret.BiffRecord.Type.TYPE_BOF)
						{
							Secret.BofRecord pBofRecord = (Secret.BofRecord)(pBiffRecord);
							switch (pBofRecord.GetBofType())
							{
								case Secret.BofRecord.BofType.BOF_TYPE_WORKBOOK_GLOBALS:
								{
									{
										NumberDuck.Secret.BiffRecord __3036547922 = pBiffRecord;
										pBiffRecord = null;
										m_pImpl.m_pWorkbookGlobals = new Secret.BiffWorkbookGlobals(__3036547922, pStream);
									}
									continue;
								}

								case Secret.BofRecord.BofType.BOF_TYPE_SHEET:
								{
									Worksheet pWorksheet;
									{
										NumberDuck.Secret.BiffRecord __3036547922 = pBiffRecord;
										pBiffRecord = null;
										pWorksheet = new Secret.BiffWorksheet(this, (Secret.BiffWorkbookGlobals)(m_pImpl.m_pWorkbookGlobals), __3036547922, pStream);
									}
									{
										NumberDuck.Worksheet __3928651719 = pWorksheet;
										pWorksheet = null;
										m_pImpl.m_pWorksheetVector.PushBack(__3928651719);
									}
									{
										continue;
									}
								}

								default:
								{
									Secret.nbAssert.Assert(false);
									break;
								}

							}
						}
						Secret.nbAssert.Assert(false);
					}
					bLoaded = true;
				}
			}
			{
				pCompoundFile = null;
			}
			if (!bLoaded)
			{
				Secret.InternalString sFileName = new Secret.InternalString("");
				Secret.Zip pZip = new Secret.Zip();
				bool bContinue = pZip.LoadFile(szFileName);
				if (bContinue)
				{
					Secret.OwnedVector<Secret.InternalString> sNameVector = new Secret.OwnedVector<Secret.InternalString>();
					{
						Blob pXmlBlob = new Blob(true);
						BlobView pXmlBlobView = pXmlBlob.GetBlobView();
						if (pZip.ExtractFileByName("[Content_Types].xml", pXmlBlobView))
						{
							Secret.XmlFile pXmlFile = new Secret.XmlFile();
							pXmlBlobView.SetOffset(0);
							if (pXmlFile.Load(pXmlBlobView))
							{
							}
							{
								pXmlFile = null;
							}
						}
						{
							pXmlBlob = null;
						}
					}
					m_pImpl.m_pWorkbookGlobals = new Secret.XlsxWorkbookGlobals();
					if (bContinue)
					{
						Secret.XmlFile pXmlFile = new Secret.XmlFile();
						Blob pXmlBlob = new Blob(true);
						BlobView pXmlBlobView = pXmlBlob.GetBlobView();
						Secret.XmlNode pWorkbookNode = null;
						Secret.XmlNode pSheetsNode = null;
						Secret.XmlNode pSheetNode = null;
						bContinue = bContinue && pZip.ExtractFileByName("xl/workbook.xml", pXmlBlobView);
						if (bContinue)
						{
							pXmlBlobView.SetOffset(0);
							if (!pXmlFile.Load(pXmlBlobView))
								bContinue = false;
						}
						if (bContinue)
						{
							pWorkbookNode = pXmlFile.GetFirstChildElement("workbook");
							if (pWorkbookNode == null)
								bContinue = false;
						}
						if (bContinue)
						{
							pSheetsNode = pWorkbookNode.GetFirstChildElement("sheets");
							if (pSheetsNode == null)
								bContinue = false;
						}
						if (bContinue)
						{
							pSheetNode = pSheetsNode.GetFirstChildElement("sheet");
							while (pSheetNode != null)
							{
								string szName = pSheetNode.GetAttribute("name");
								if (szName == null)
								{
									bContinue = false;
									break;
								}
								sNameVector.PushBack(new Secret.InternalString(szName));
								pSheetNode = pSheetNode.GetNextSiblingElement("sheet");
							}
						}
						{
							pXmlBlob = null;
						}
						{
							pXmlFile = null;
						}
					}
					if (bContinue)
					{
						Secret.XmlFile pXmlFile = new Secret.XmlFile();
						Blob pXmlBlob = new Blob(true);
						BlobView pXmlBlobView = pXmlBlob.GetBlobView();
						Secret.XmlNode pSstNode = null;
						Secret.XmlNode pSiNode = null;
						bContinue = bContinue && pZip.ExtractFileByName("xl/sharedStrings.xml", pXmlBlobView);
						if (bContinue)
						{
							pXmlBlobView.SetOffset(0);
							if (!pXmlFile.Load(pXmlBlobView))
								bContinue = false;
						}
						if (bContinue)
						{
							pSstNode = pXmlFile.GetFirstChildElement("sst");
							if (pSstNode == null)
								bContinue = false;
						}
						if (bContinue)
						{
							pSiNode = pSstNode.GetFirstChildElement("si");
							if (pSiNode == null)
								bContinue = false;
						}
						while (bContinue && pSiNode != null)
						{
							Secret.XmlNode pTNode;
							pTNode = pSiNode.GetFirstChildElement("t");
							if (pTNode == null)
							{
								bContinue = false;
								break;
							}
							string szTemp = pTNode.GetText();
							m_pImpl.m_pWorkbookGlobals.PushSharedString(szTemp);
							pSiNode = pSiNode.GetNextSiblingElement("si");
						}
						{
							pXmlBlob = null;
						}
						{
							pXmlFile = null;
						}
					}
					if (bContinue)
					{
						Secret.XmlFile pXmlFile = new Secret.XmlFile();
						Blob pXmlBlob = new Blob(true);
						BlobView pXmlBlobView = pXmlBlob.GetBlobView();
						Secret.XmlNode pStyleSheetNode = null;
						bContinue = bContinue && pZip.ExtractFileByName("xl/styles.xml", pXmlBlobView);
						if (bContinue)
						{
							pXmlBlobView.SetOffset(0);
							if (!pXmlFile.Load(pXmlBlobView))
								bContinue = false;
						}
						if (bContinue)
						{
							pStyleSheetNode = pXmlFile.GetFirstChildElement("styleSheet");
							if (pStyleSheetNode == null)
								bContinue = false;
						}
						bContinue = bContinue && Secret.XlsxWorkbookGlobals.ParseStyles(m_pImpl.m_pWorkbookGlobals, pStyleSheetNode);
						{
							pXmlBlob = null;
						}
						{
							pXmlFile = null;
						}
					}
					if (bContinue)
					{
						int nWorksheetIndex = 0;
						int nNumFail = 0;
						while (bContinue)
						{
							sFileName.Set("xl/worksheets/sheet");
							sFileName.AppendInt(nWorksheetIndex);
							sFileName.AppendString(".xml");
							Secret.XmlFile pXmlFile = new Secret.XmlFile();
							Blob pXmlBlob = new Blob(true);
							BlobView pXmlBlobView = pXmlBlob.GetBlobView();
							if (!pZip.ExtractFileByName(sFileName.GetExternalString(), pXmlBlobView))
							{
								nNumFail++;
								if (nNumFail > 5)
								{
									{
										pXmlBlob = null;
									}
									{
										pXmlFile = null;
									}
									{
										break;
									}
								}
							}
							else
							{
								Secret.XmlNode pWorksheetNode;
								pXmlBlobView.SetOffset(0);
								if (!pXmlFile.Load(pXmlBlobView))
								{
									{
										pXmlBlob = null;
									}
									{
										pXmlFile = null;
									}
									bContinue = false;
									{
										break;
									}
								}
								pWorksheetNode = pXmlFile.GetFirstChildElement("worksheet");
								Secret.XlsxWorksheet pWorksheet = new Secret.XlsxWorksheet(this);
								pWorksheet.SetName(sNameVector.Get(m_pImpl.m_pWorksheetVector.GetSize()).GetExternalString());
								if (!pWorksheet.Parse((Secret.XlsxWorkbookGlobals)(m_pImpl.m_pWorkbookGlobals), pWorksheetNode))
								{
									{
										pWorksheet = null;
									}
									{
										pXmlBlob = null;
									}
									{
										pXmlFile = null;
									}
									bContinue = false;
									{
										break;
									}
								}
								{
									NumberDuck.Secret.XlsxWorksheet __3928651719 = pWorksheet;
									pWorksheet = null;
									m_pImpl.m_pWorksheetVector.PushBack(__3928651719);
								}
							}
							nWorksheetIndex++;
							{
								pXmlBlob = null;
							}
							{
								pXmlFile = null;
							}
						}
					}
					bLoaded = bContinue;
					sNameVector.Clear();
					{
						sNameVector = null;
					}
				}
				{
					pZip = null;
				}
				{
					sFileName = null;
				}
			}
			if (bLoaded)
			{
				return true;
			}
			Clear();
			{
				return false;
			}
		}

		public bool Save(string szFileName, FileType eFileType)
		{
			m_pImpl.m_pWorkbookGlobals.Clear();
			if (eFileType == FileType.XLS)
			{
				Secret.Vector<Secret.BiffRecordContainer> pWorksheetBiffRecordContainerVector = new Secret.Vector<Secret.BiffRecordContainer>();
				for (int i = 0; i < m_pImpl.m_pWorksheetVector.GetSize(); i++)
				{
					Worksheet pWorksheet = m_pImpl.m_pWorksheetVector.Get(i);
					Secret.BiffRecordContainer pBiffRecordContainer = new Secret.BiffRecordContainer();
					Secret.BiffWorksheet.Write(pWorksheet, m_pImpl.m_pWorkbookGlobals, (ushort)(i), pBiffRecordContainer);
					m_pImpl.m_pWorkbookGlobals.PushBiffWorksheetStreamSize(pBiffRecordContainer.GetSize());
					pWorksheetBiffRecordContainerVector.PushBack(pBiffRecordContainer);
					pBiffRecordContainer = null;
				}
				Secret.CompoundFile pCompoundFile = new Secret.CompoundFile();
				Secret.Stream pStream = pCompoundFile.CreateStream("Workbook", Secret.Stream.Type.TYPE_USER_STREAM);
				Secret.BiffWorkbookGlobals.Write(m_pImpl.m_pWorkbookGlobals, m_pImpl.m_pWorksheetVector, pStream);
				for (int i = 0; i < pWorksheetBiffRecordContainerVector.GetSize(); i++)
				{
					Secret.BiffRecordContainer pBiffRecordContainer = pWorksheetBiffRecordContainerVector.Get(i);
					pBiffRecordContainer.Write(pStream);
					{
						pBiffRecordContainer = null;
					}
				}
				{
					pWorksheetBiffRecordContainerVector = null;
				}
				bool bResult = pCompoundFile.Save(szFileName);
				{
					pCompoundFile = null;
				}
				{
					return bResult;
				}
			}
			else
			{
				return false;
			}
		}

		public uint GetNumWorksheet()
		{
			return (uint)(m_pImpl.m_pWorksheetVector.GetSize());
		}

		public Worksheet GetWorksheetByIndex(uint nIndex)
		{
			if (nIndex >= GetNumWorksheet())
				return null;
			return m_pImpl.m_pWorksheetVector.Get((int)(nIndex));
		}

		public Worksheet CreateWorksheet()
		{
			Worksheet pWorksheet = new Secret.BiffWorksheet(this);
			int i = 0;
			Secret.InternalString sName = new Secret.InternalString("");
			while (true)
			{
				sName.Set("Sheet");
				sName.AppendInt(++i);
				if (pWorksheet.SetName(sName.GetExternalString()))
					break;
			}
			{
				sName = null;
			}
			Worksheet pTempWorksheet = pWorksheet;
			{
				NumberDuck.Worksheet __3928651719 = pWorksheet;
				pWorksheet = null;
				m_pImpl.m_pWorksheetVector.PushBack(__3928651719);
			}
			{
				return pTempWorksheet;
			}
		}

		public void PurgeWorksheet(uint nIndex)
		{
			Worksheet pWorksheet = GetWorksheetByIndex(nIndex);
			if (pWorksheet != null)
			{
				m_pImpl.m_pWorksheetVector.Erase((int)(nIndex));
			}
			if (GetNumWorksheet() == 0)
				CreateWorksheet();
		}

		public uint GetNumStyle()
		{
			return m_pImpl.m_pWorkbookGlobals.GetNumStyle();
		}

		public Style GetStyleByIndex(uint nIndex)
		{
			return m_pImpl.m_pWorkbookGlobals.GetStyleByIndex((ushort)(nIndex));
		}

		public Style GetDefaultStyle()
		{
			return GetStyleByIndex(0);
		}

		public Style CreateStyle()
		{
			return m_pImpl.m_pWorkbookGlobals.CreateStyle();
		}

		~Workbook()
		{
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ValueImplementation
		{
			public Value.Type m_eType;
			public InternalString m_sString;
			public double m_fFloat;
			public bool m_bBoolean;
			public Formula m_pFormula;
			public Worksheet m_pWorksheet;
			public Area m_pArea;
			public Area3d m_pArea3d;
			public Value m_pValue;
			public ValueImplementation()
			{
				m_pFormula = null;
				m_pValue = null;
				m_sString = new InternalString("");
				m_pArea = null;
				m_pArea3d = null;
			}

			public static Value CreateFloatValue(double fFloat)
			{
				Value pValue = new Value();
				pValue.m_pImpl.SetFloat(fFloat);
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CreateStringValue(string szString)
			{
				Value pValue = new Value();
				pValue.m_pImpl.SetString(szString);
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CreateBooleanValue(bool bBoolean)
			{
				Value pValue = new Value();
				pValue.m_pImpl.SetBoolean(bBoolean);
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CreateErrorValue()
			{
				Value pValue = new Value();
				pValue.m_pImpl.m_eType = Value.Type.TYPE_ERROR;
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CreateAreaValue(Area pArea)
			{
				Value pValue = new Value();
				pValue.m_pImpl.m_eType = Value.Type.TYPE_AREA;
				pValue.m_pImpl.m_pArea = pArea;
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CreateArea3dValue(Area3d pArea3d)
			{
				Value pValue = new Value();
				pValue.m_pImpl.m_eType = Value.Type.TYPE_AREA_3D;
				pValue.m_pImpl.m_pArea3d = pArea3d;
				{
					NumberDuck.Value __482980084 = pValue;
					pValue = null;
					return __482980084;
				}
			}

			public static Value CopyValue(Value pValue)
			{
				Value pNewValue = new Value();
				nbAssert.Assert(pValue.m_pImpl.m_eType != Value.Type.TYPE_FORMULA);
				pNewValue.m_pImpl.m_eType = pValue.m_pImpl.m_eType;
				pNewValue.m_pImpl.m_sString.Set(pValue.m_pImpl.m_sString.GetExternalString());
				pNewValue.m_pImpl.m_fFloat = pValue.m_pImpl.m_fFloat;
				pNewValue.m_pImpl.m_bBoolean = pValue.m_pImpl.m_bBoolean;
				if (pValue.m_pImpl.m_pArea != null)
					pNewValue.m_pImpl.m_pArea = pValue.m_pImpl.m_pArea.CreateClone();
				if (pValue.m_pImpl.m_pArea3d != null)
					pNewValue.m_pImpl.m_pArea3d = pValue.m_pImpl.m_pArea3d.CreateClone();
				{
					NumberDuck.Value __4153834605 = pNewValue;
					pNewValue = null;
					return __4153834605;
				}
			}

			public void Clear()
			{
				m_eType = Value.Type.TYPE_EMPTY;
			}

			public void SetString(string szString)
			{
				m_eType = Value.Type.TYPE_STRING;
				m_sString.Set(szString);
			}

			public void SetFloat(double fFloat)
			{
				m_eType = Value.Type.TYPE_FLOAT;
				m_fFloat = fFloat;
			}

			public void SetBoolean(bool bBoolean)
			{
				m_eType = Value.Type.TYPE_BOOLEAN;
				m_bBoolean = bBoolean;
			}

			public void SetFormulaFromString(string szFormula, Worksheet pWorksheet)
			{
				SetFormula(new Formula(szFormula, pWorksheet.m_pImpl), pWorksheet);
			}

			public void SetFormula(Formula pFormula, Worksheet pWorksheet)
			{
				nbAssert.Assert(pFormula != null);
				nbAssert.Assert(pWorksheet != null);
				m_eType = Value.Type.TYPE_FORMULA;
				if (m_pFormula != null)
					{
						m_pFormula = null;
					}
				m_pFormula = pFormula;
				m_pWorksheet = pWorksheet;
			}

			~ValueImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class StyleImplementation
		{
			public Font m_pFont;
			public Style.HorizontalAlign m_eHorizontalAlign;
			public Style.VerticalAlign m_eVerticalAlign;
			public Color m_pBackgroundColor;
			public Style.FillPattern m_eFillPattern;
			public Color m_pFillPatternColor;
			public Line m_pTopBorderLine;
			public Line m_pRightBorderLine;
			public Line m_pBottomBorderLine;
			public Line m_pLeftBorderLine;
			public InternalString m_sFormat;
			public uint m_nFormatIndex;
			public StyleImplementation()
			{
				m_pFont = new Font();
				m_eHorizontalAlign = Style.HorizontalAlign.HORIZONTAL_ALIGN_GENERAL;
				m_eVerticalAlign = Style.VerticalAlign.VERTICAL_ALIGN_BOTTOM;
				m_pBackgroundColor = null;
				m_eFillPattern = Style.FillPattern.FILL_PATTERN_NONE;
				m_pFillPatternColor = null;
				m_pTopBorderLine = new Line();
				m_pTopBorderLine.SetType(Line.Type.TYPE_NONE);
				m_pRightBorderLine = new Line();
				m_pRightBorderLine.SetType(Line.Type.TYPE_NONE);
				m_pBottomBorderLine = new Line();
				m_pBottomBorderLine.SetType(Line.Type.TYPE_NONE);
				m_pLeftBorderLine = new Line();
				m_pLeftBorderLine.SetType(Line.Type.TYPE_NONE);
				m_sFormat = new InternalString("General");
			}

			~StyleImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SharedString
		{
			public InternalString m_sString;
			public int m_nChecksum;
			public int m_nIndex;
			~SharedString()
			{
			}

		}
		class SharedStringContainer
		{
			protected OwnedVector<SharedString> m_pSharedStringVector;
			protected Vector<SharedString> m_pSharedStringSortedVector;
			public SharedStringContainer()
			{
				m_pSharedStringVector = new OwnedVector<SharedString>();
				m_pSharedStringSortedVector = new Vector<SharedString>();
			}

			~SharedStringContainer()
			{
			}

			public void Clear()
			{
				m_pSharedStringVector.Clear();
				m_pSharedStringSortedVector.Clear();
			}

			public string Get(int nIndex)
			{
				return m_pSharedStringVector.Get(nIndex).m_sString.GetExternalString();
			}

			public int GetIndex(string sxString)
			{
				int nChecksum = ExternalString.GetChecksum(sxString);
				int i = 0;
				while (i < m_pSharedStringSortedVector.GetSize())
				{
					SharedString pTest = m_pSharedStringSortedVector.Get(i);
					if (pTest.m_nChecksum > nChecksum)
						break;
					if (pTest.m_nChecksum == nChecksum && ExternalString.Equal(sxString, pTest.m_sString.GetExternalString()))
						return pTest.m_nIndex;
					i++;
				}
				return Push(sxString);
			}

			public int Push(string sxString)
			{
				int nIndex = m_pSharedStringVector.GetSize();
				SharedString pSharedString = new SharedString();
				pSharedString.m_sString = new InternalString(sxString);
				pSharedString.m_nIndex = nIndex;
				pSharedString.m_nChecksum = ExternalString.GetChecksum(sxString);
				int i = 0;
				while (i < m_pSharedStringSortedVector.GetSize())
				{
					SharedString pTest = m_pSharedStringSortedVector.Get(i);
					if (pTest.m_nChecksum >= pSharedString.m_nChecksum)
						break;
					i++;
				}
				m_pSharedStringSortedVector.Insert(i, pSharedString);
				{
					NumberDuck.Secret.SharedString __3093361001 = pSharedString;
					pSharedString = null;
					m_pSharedStringVector.PushBack(__3093361001);
				}
				{
					return nIndex;
				}
			}

			public int GetSize()
			{
				return m_pSharedStringVector.GetSize();
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class SeriesImplementation
		{
			public Worksheet m_pWorksheet;
			public Formula m_pNameFormula;
			public Formula m_pValuesFormula;
			public Line m_pLine;
			public Fill m_pFill;
			public Marker m_pMarker;
			public SeriesImplementation(Worksheet pWorksheet, Formula pValuesFormula)
			{
				m_pWorksheet = pWorksheet;
				m_pNameFormula = null;
				m_pValuesFormula = pValuesFormula;
				m_pLine = new Line();
				m_pFill = new Fill();
				m_pMarker = new Marker();
			}

			public void SetNameFormula(Formula pFormula)
			{
				nbAssert.Assert(pFormula != null);
				{
					m_pNameFormula = null;
				}
				m_pNameFormula = pFormula;
			}

			public void SetValuesFormula(Formula pFormula)
			{
				nbAssert.Assert(pFormula != null);
				{
					m_pValuesFormula = null;
				}
				m_pValuesFormula = pFormula;
			}

			public void SetClassicStyle(Chart.Type eChartType, ushort nIndex)
			{
				ushort nColorIndex = (ushort)(BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + 1 + (24 + nIndex - 1) % (BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY - 1));
				m_pLine.GetColor().SetFromRgba(BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(nColorIndex));
				nColorIndex = (ushort)(BiffWorkbookGlobals.NUM_DEFAULT_PALETTE_ENTRY + 1 + (16 + nIndex - 1) % (BiffWorkbookGlobals.NUM_CUSTOM_PALETTE_ENTRY - 1));
				m_pFill.GetForegroundColor().SetFromRgba(BiffWorkbookGlobals.GetDefaultPaletteColorByIndex(nColorIndex));
				m_pMarker.GetBorderColor(true).SetFromColor(m_pLine.GetColor());
				m_pMarker.GetFillColor(true).SetFromColor(m_pLine.GetColor());
				switch (eChartType)
				{
					case Chart.Type.TYPE_COLUMN:
					case Chart.Type.TYPE_COLUMN_STACKED:
					case Chart.Type.TYPE_COLUMN_STACKED_100:
					case Chart.Type.TYPE_BAR:
					case Chart.Type.TYPE_BAR_STACKED:
					case Chart.Type.TYPE_BAR_STACKED_100:
					case Chart.Type.TYPE_AREA:
					case Chart.Type.TYPE_AREA_STACKED:
					case Chart.Type.TYPE_AREA_STACKED_100:
					{
						m_pLine.GetColor().Set(0x00, 0x00, 0x00);
						break;
					}

					case Chart.Type.TYPE_LINE:
					case Chart.Type.TYPE_LINE_STACKED:
					case Chart.Type.TYPE_LINE_STACKED_100:
					{
						break;
					}

					case Chart.Type.TYPE_SCATTER:
					{
						m_pLine.SetType(Line.Type.TYPE_NONE);
						break;
					}

					default:
					{
						nbAssert.Assert(false);
						break;
					}

				}
				switch (nIndex % 9)
				{
					case 0:
					{
						m_pMarker.SetType(Marker.Type.TYPE_DIAMOND);
						break;
					}

					case 1:
					{
						m_pMarker.SetType(Marker.Type.TYPE_SQUARE);
						break;
					}

					case 2:
					{
						m_pMarker.SetType(Marker.Type.TYPE_TRIANGLE);
						break;
					}

					case 3:
					{
						m_pMarker.SetType(Marker.Type.TYPE_X);
						break;
					}

					case 4:
					{
						m_pMarker.SetType(Marker.Type.TYPE_ASTERISK);
						break;
					}

					case 5:
					{
						m_pMarker.SetType(Marker.Type.TYPE_CIRCULAR);
						break;
					}

					case 6:
					{
						m_pMarker.SetType(Marker.Type.TYPE_PLUS);
						break;
					}

					case 7:
					{
						m_pMarker.SetType(Marker.Type.TYPE_SHORT_BAR);
						break;
					}

					case 8:
					{
						m_pMarker.SetType(Marker.Type.TYPE_LONG_BAR);
						break;
					}

				}
				if (m_pMarker.GetType() == Marker.Type.TYPE_X || m_pMarker.GetType() == Marker.Type.TYPE_ASTERISK || m_pMarker.GetType() == Marker.Type.TYPE_PLUS)
					m_pMarker.ClearFillColor();
			}

			~SeriesImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	class Series
	{
		public Secret.SeriesImplementation m_pImpl;
		public Series(Worksheet pWorksheet, Secret.Formula pValuesFormula)
		{
			m_pImpl = new Secret.SeriesImplementation(pWorksheet, pValuesFormula);
		}

		public string GetName()
		{
			if (m_pImpl.m_pNameFormula != null)
				return m_pImpl.m_pNameFormula.ToChartNameString(m_pImpl.m_pWorksheet.m_pImpl);
			return "=";
		}

		public bool SetName(string szName)
		{
			Secret.Formula pFormula = new Secret.Formula(szName, m_pImpl.m_pWorksheet.m_pImpl);
			if (pFormula.ValidateForChartName(m_pImpl.m_pWorksheet.m_pImpl))
			{
				{
				}
				{
					NumberDuck.Secret.Formula __879619620 = pFormula;
					pFormula = null;
					m_pImpl.m_pNameFormula = __879619620;
				}
				{
					return true;
				}
			}
			{
				pFormula = null;
			}
			{
				return false;
			}
		}

		public string GetValues()
		{
			if (m_pImpl.m_pValuesFormula != null)
				return m_pImpl.m_pValuesFormula.ToChartString(m_pImpl.m_pWorksheet.m_pImpl);
			return "=";
		}

		public bool SetValues(string szValues)
		{
			Secret.Formula pFormula = new Secret.Formula(szValues, m_pImpl.m_pWorksheet.m_pImpl);
			if (pFormula.ValidateForChart(m_pImpl.m_pWorksheet.m_pImpl))
			{
				{
				}
				{
					NumberDuck.Secret.Formula __879619620 = pFormula;
					pFormula = null;
					m_pImpl.m_pValuesFormula = __879619620;
				}
				{
					return true;
				}
			}
			{
				pFormula = null;
			}
			{
				return false;
			}
		}

		public Line GetLine()
		{
			return m_pImpl.m_pLine;
		}

		public Fill GetFill()
		{
			return m_pImpl.m_pFill;
		}

		public Marker GetMarker()
		{
			return m_pImpl.m_pMarker;
		}

		~Series()
		{
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class PictureImplementation
		{
			public uint m_nX;
			public uint m_nY;
			public uint m_nSubX;
			public uint m_nSubY;
			public uint m_nWidth;
			public uint m_nHeight;
			public InternalString m_sUrl;
			public Blob m_pBlob;
			public Picture.Format m_eFormat;
			public PictureImplementation(Blob pBlob, Picture.Format eFormat)
			{
				m_nX = 0;
				m_nY = 0;
				m_nSubX = 0;
				m_nSubY = 0;
				m_nWidth = 200;
				m_nHeight = 100;
				m_sUrl = new InternalString("");
				BlobView pBlobView = pBlob.GetBlobView();
				pBlobView.SetOffset(0);
				m_pBlob = new Blob(false);
				m_pBlob.Resize(pBlob.GetSize(), false);
				m_pBlob.GetBlobView().Pack(pBlobView, pBlob.GetSize());
				m_eFormat = eFormat;
			}

			~PictureImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MergedCellImplementation
		{
			public uint m_nX;
			public uint m_nY;
			public uint m_nWidth;
			public uint m_nHeight;
		}
	}
}

namespace NumberDuck
{
	class MergedCell
	{
		protected Secret.MergedCellImplementation m_pImpl;
		public MergedCell(uint nX, uint nY, uint nWidth, uint nHeight)
		{
			m_pImpl = new Secret.MergedCellImplementation();
			SetX(nX);
			SetY(nY);
			SetWidth(nWidth);
			SetHeight(nHeight);
		}

		public uint GetX()
		{
			return m_pImpl.m_nX;
		}

		public void SetX(uint nX)
		{
			if (nX > 0xFF)
				nX = 0xFF;
			m_pImpl.m_nX = nX;
		}

		public uint GetY()
		{
			return m_pImpl.m_nY;
		}

		public void SetY(uint nY)
		{
			if (nY > 0xFFFF)
				nY = 0xFFFF;
			m_pImpl.m_nY = nY;
		}

		public uint GetWidth()
		{
			return m_pImpl.m_nWidth;
		}

		public void SetWidth(uint nWidth)
		{
			if (nWidth == 0)
				nWidth = 1;
			if (m_pImpl.m_nX + nWidth > 0xFF + 1)
				nWidth = 0xFF - m_pImpl.m_nX + 1;
			m_pImpl.m_nWidth = nWidth;
		}

		public uint GetHeight()
		{
			return m_pImpl.m_nHeight;
		}

		public void SetHeight(uint nHeight)
		{
			if (nHeight == 0)
				nHeight = 1;
			if (m_pImpl.m_nY + nHeight > 0xFFFF + 1)
				nHeight = 0xFFFF - m_pImpl.m_nY + 1;
			m_pImpl.m_nHeight = nHeight;
		}

		~MergedCell()
		{
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class MarkerImplementation
		{
			public Marker.Type m_eType;
			public Color m_pFillColor;
			public Color m_pBorderColor;
			public int m_nSize;
			public MarkerImplementation()
			{
				m_eType = Marker.Type.TYPE_SQUARE;
				m_nSize = WorksheetImplementation.TwipsToPixels(100);
				m_pFillColor = new Color(255, 0, 255);
				m_pBorderColor = new Color(0, 255, 0);
			}

			public bool Equals(MarkerImplementation pMarkerImplementation)
			{
				return pMarkerImplementation != null && m_eType == pMarkerImplementation.m_eType && (m_pFillColor == null && pMarkerImplementation.m_pFillColor == null || m_pFillColor != null && m_pFillColor.Equals(pMarkerImplementation.m_pFillColor)) && (m_pBorderColor == null && pMarkerImplementation.m_pBorderColor == null || m_pBorderColor != null && m_pBorderColor.Equals(pMarkerImplementation.m_pBorderColor));
			}

			public new Marker.Type GetType()
			{
				return m_eType;
			}

			public void SetType(Marker.Type eType)
			{
				if (eType >= Marker.Type.TYPE_NONE && eType < Marker.Type.NUM_TYPE)
					m_eType = eType;
			}

			public Color GetFillColor(bool bCreateIfMissing)
			{
				if (m_pFillColor == null && bCreateIfMissing)
					m_pFillColor = new Color(255, 0, 255);
				return m_pFillColor;
			}

			public void SetFillColor(Color pColor)
			{
				if (m_pFillColor == null)
					m_pFillColor = new Color(255, 0, 255);
				m_pFillColor.SetFromColor(pColor);
			}

			public void ClearFillColor()
			{
				{
					m_pFillColor = null;
				}
				m_pFillColor = null;
			}

			public Color GetBorderColor(bool bCreateIfMissing)
			{
				if (m_pBorderColor == null && bCreateIfMissing)
					m_pBorderColor = new Color(0, 255, 0);
				return m_pBorderColor;
			}

			public void SetBorderColor(Color pColor)
			{
				if (m_pBorderColor == null)
					m_pBorderColor = new Color(0, 255, 0);
				m_pBorderColor.SetFromColor(pColor);
			}

			public void ClearBorderColor()
			{
				{
					m_pBorderColor = null;
				}
				m_pBorderColor = null;
			}

			public int GetSize()
			{
				return m_nSize;
			}

			public void SetSize(int nSize)
			{
				if (nSize >= WorksheetImplementation.TwipsToPixels(40) && nSize <= WorksheetImplementation.TwipsToPixels(1440))
					m_nSize = nSize;
			}

			~MarkerImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LineImplementation
		{
			public Line.Type m_eType;
			public Color m_pColor;
			public LineImplementation()
			{
				m_eType = Line.Type.TYPE_THIN;
				m_pColor = new Color(0xFF, 0x00, 0x00);
			}

			public bool Equals(LineImplementation pLineImplementation)
			{
				return pLineImplementation != null && m_eType == pLineImplementation.m_eType && m_pColor.Equals(pLineImplementation.m_pColor);
			}

			public new Line.Type GetType()
			{
				return m_eType;
			}

			public void SetType(Line.Type eType)
			{
				if (eType >= Line.Type.TYPE_NONE && eType < Line.Type.NUM_TYPE)
					m_eType = eType;
			}

			~LineImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class LegendImplementation
		{
			public bool m_bHidden;
			public Line m_pBorderLine;
			public Fill m_pFill;
			public LegendImplementation()
			{
				m_bHidden = false;
				m_pBorderLine = new Line();
				m_pFill = new Fill();
			}

			~LegendImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	class Legend
	{
		protected Secret.LegendImplementation m_pImpl;
		public Legend()
		{
			m_pImpl = new Secret.LegendImplementation();
		}

		public bool GetHidden()
		{
			return m_pImpl.m_bHidden;
		}

		public void SetHidden(bool bHidden)
		{
			m_pImpl.m_bHidden = bHidden;
		}

		public Line GetBorderLine()
		{
			return m_pImpl.m_pBorderLine;
		}

		public Fill GetFill()
		{
			return m_pImpl.m_pFill;
		}

		~Legend()
		{
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FontImplementation
		{
			public InternalString m_sName;
			public int m_nSizeTwips;
			public Color m_pColor;
			public bool m_bBold;
			public bool m_bItalic;
			public Font.Underline m_eUnderline;
			public FontImplementation()
			{
				m_sName = new InternalString("Arial");
				m_nSizeTwips = 14 * 15;
				m_pColor = null;
				m_bBold = false;
				m_bItalic = false;
				m_eUnderline = Font.Underline.UNDERLINE_NONE;
			}

			~FontImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class FillImplementation
		{
			public Fill.Type m_eType;
			public Color m_pForegroundColor;
			public Color m_pBackgroundColor;
			public FillImplementation()
			{
				m_eType = Fill.Type.TYPE_SOLID;
				m_pForegroundColor = new Color(0, 255, 0);
				m_pBackgroundColor = new Color(0, 0, 255);
			}

			public bool Equals(FillImplementation pFillImplementation)
			{
				return pFillImplementation != null && m_eType == pFillImplementation.m_eType && m_pForegroundColor.Equals(pFillImplementation.m_pForegroundColor) && m_pBackgroundColor.Equals(pFillImplementation.m_pBackgroundColor);
			}

			public new Fill.Type GetType()
			{
				return m_eType;
			}

			public void SetType(Fill.Type eType)
			{
				if (eType >= Fill.Type.TYPE_NONE && eType < Fill.Type.NUM_TYPE)
					m_eType = eType;
			}

			~FillImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	class Color
	{
		protected byte m_nRed;
		protected byte m_nGreen;
		protected byte m_nBlue;
		public Color(byte nRed, byte nGreen, byte nBlue)
		{
			m_nRed = nRed;
			m_nGreen = nGreen;
			m_nBlue = nBlue;
		}

		public bool Equals(Color pColor)
		{
			if (pColor == null)
				return false;
			return m_nRed == pColor.m_nRed && m_nGreen == pColor.m_nGreen && m_nBlue == pColor.m_nBlue;
		}

		public byte GetRed()
		{
			return m_nRed;
		}

		public byte GetGreen()
		{
			return m_nGreen;
		}

		public byte GetBlue()
		{
			return m_nBlue;
		}

		public void Set(byte nRed, byte nGreen, byte nBlue)
		{
			m_nRed = nRed;
			m_nGreen = nGreen;
			m_nBlue = nBlue;
		}

		public void SetRed(byte nRed)
		{
			m_nRed = nRed;
		}

		public void SetGreen(byte nGreen)
		{
			m_nGreen = nGreen;
		}

		public void SetBlue(byte nBlue)
		{
			m_nBlue = nBlue;
		}

		public void SetFromColor(Color pColor)
		{
			m_nRed = pColor.m_nRed;
			m_nGreen = pColor.m_nGreen;
			m_nBlue = pColor.m_nBlue;
		}

		public void SetFromRgba(uint nRgba)
		{
			m_nRed = (byte)(nRgba & 0xFF);
			m_nGreen = (byte)((nRgba >> 8) & 0xFF);
			m_nBlue = (byte)((nRgba >> 16) & 0xFF);
		}

		public uint GetRgba()
		{
			return ((uint)(m_nRed) << 0) | ((uint)(m_nGreen) << 8) | ((uint)(m_nBlue) << 16) | ((uint)(0xFF) << 24);
		}

	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class ChartImplementation
		{
			public Worksheet m_pWorksheet;
			public uint m_nX;
			public uint m_nY;
			public uint m_nSubX;
			public uint m_nSubY;
			public uint m_nWidth;
			public uint m_nHeight;
			public Chart.Type m_eType;
			public Formula m_pCategoriesFormula;
			public InternalString m_sTitle;
			public InternalString m_sHorizontalAxisLabel;
			public InternalString m_sVerticalAxisLabel;
			public Legend m_pLegend;
			public Line m_pFrameBorderLine;
			public Fill m_pFrameFill;
			public Line m_pPlotBorderLine;
			public Fill m_pPlotFill;
			public Line m_pHorizontalAxisLine;
			public Line m_pHorizontalGridLine;
			public Line m_pVerticalAxisLine;
			public Line m_pVerticalGridLine;
			public OwnedVector<Series> m_pSeriesVector;
			public ChartImplementation(Worksheet pWorksheet, Chart.Type eType)
			{
				m_pWorksheet = pWorksheet;
				m_nX = 0;
				m_nY = 0;
				m_nSubX = 0;
				m_nSubY = 0;
				m_nWidth = 200;
				m_nHeight = 100;
				m_eType = Chart.Type.TYPE_COLUMN;
				if (eType >= Chart.Type.TYPE_COLUMN && eType < Chart.Type.NUM_TYPE)
					m_eType = eType;
				m_pCategoriesFormula = null;
				m_pLegend = new Legend();
				m_sTitle = new InternalString("");
				m_sHorizontalAxisLabel = new InternalString("");
				m_sVerticalAxisLabel = new InternalString("");
				m_pFrameBorderLine = new Line();
				m_pFrameFill = new Fill();
				m_pPlotBorderLine = new Line();
				m_pPlotFill = new Fill();
				m_pHorizontalAxisLine = new Line();
				m_pHorizontalGridLine = new Line();
				m_pVerticalAxisLine = new Line();
				m_pVerticalGridLine = new Line();
				SetClassicStyle();
				m_pLegend.SetHidden(false);
				m_pLegend.GetBorderLine().GetColor().Set(0x00, 0x00, 0x00);
				m_pLegend.GetFill().GetForegroundColor().Set(0xFF, 0xFF, 0xFF);
				m_pSeriesVector = new OwnedVector<Series>();
			}

			public Series CreateSeries(Formula pValuesFormula)
			{
				nbAssert.Assert(pValuesFormula != null);
				Series pSeries = new Series(m_pWorksheet, pValuesFormula);
				pSeries.m_pImpl.SetClassicStyle(m_eType, (ushort)(m_pSeriesVector.GetSize()));
				Series pTemp = pSeries;
				{
					NumberDuck.Series __1756674346 = pSeries;
					pSeries = null;
					m_pSeriesVector.PushBack(__1756674346);
				}
				{
					return pTemp;
				}
			}

			public void PurgeSeries(int nIndex)
			{
				if (nIndex >= 0 && nIndex < m_pSeriesVector.GetSize())
				{
					m_pSeriesVector.Erase(nIndex);
				}
			}

			public void SetCategoriesFormula(Formula pCategoriesFormula)
			{
				nbAssert.Assert(pCategoriesFormula != null);
				{
					m_pCategoriesFormula = null;
				}
				m_pCategoriesFormula = pCategoriesFormula;
			}

			public void SetClassicStyle()
			{
				m_pFrameBorderLine.GetColor().Set(0xFF, 0xFF, 0xFF);
				m_pFrameFill.GetForegroundColor().Set(0xFF, 0xFF, 0xFF);
				m_pPlotBorderLine.GetColor().Set(0x84, 0x84, 0x84);
				m_pPlotFill.GetForegroundColor().Set(0xD0, 0xD0, 0xD0);
				m_pHorizontalAxisLine.GetColor().Set(0x00, 0x00, 0x00);
				m_pHorizontalGridLine.SetType(Line.Type.TYPE_NONE);
				m_pVerticalAxisLine.GetColor().Set(0x00, 0x00, 0x00);
				m_pVerticalGridLine.SetType(Line.Type.TYPE_NONE);
				m_pLegend.SetHidden(true);
				m_pLegend.GetBorderLine().GetColor().Set(0x00, 0x00, 0x00);
				m_pLegend.GetFill().GetForegroundColor().Set(0xFF, 0xFF, 0xFF);
			}

			public void InsertColumn(ushort nWorksheet, ushort nColumn)
			{
				if (m_pCategoriesFormula != null)
					m_pCategoriesFormula.InsertColumn(nWorksheet, nColumn);
				for (int i = 0; i < m_pSeriesVector.GetSize(); i++)
				{
					Series pSeries = m_pSeriesVector.Get(i);
					pSeries.m_pImpl.m_pNameFormula.InsertColumn(nWorksheet, nColumn);
					pSeries.m_pImpl.m_pValuesFormula.InsertColumn(nWorksheet, nColumn);
				}
			}

			public void DeleteColumn(ushort nWorksheet, ushort nColumn)
			{
				if (m_pCategoriesFormula != null)
					m_pCategoriesFormula.DeleteColumn(nWorksheet, nColumn);
				for (int i = 0; i < m_pSeriesVector.GetSize(); i++)
				{
					Series pSeries = m_pSeriesVector.Get(i);
					pSeries.m_pImpl.m_pNameFormula.DeleteColumn(nWorksheet, nColumn);
					pSeries.m_pImpl.m_pValuesFormula.DeleteColumn(nWorksheet, nColumn);
				}
			}

			public void InsertRow(ushort nWorksheet, ushort nRow)
			{
				if (m_pCategoriesFormula != null)
					m_pCategoriesFormula.InsertRow(nWorksheet, nRow);
				for (int i = 0; i < m_pSeriesVector.GetSize(); i++)
				{
					Series pSeries = m_pSeriesVector.Get(i);
					pSeries.m_pImpl.m_pNameFormula.InsertRow(nWorksheet, nRow);
					pSeries.m_pImpl.m_pValuesFormula.InsertRow(nWorksheet, nRow);
				}
			}

			public void DeleteRow(ushort nWorksheet, ushort nRow)
			{
				if (m_pCategoriesFormula != null)
					m_pCategoriesFormula.DeleteRow(nWorksheet, nRow);
				for (int i = 0; i < m_pSeriesVector.GetSize(); i++)
				{
					Series pSeries = m_pSeriesVector.Get(i);
					pSeries.m_pImpl.m_pNameFormula.DeleteRow(nWorksheet, nRow);
					pSeries.m_pImpl.m_pValuesFormula.DeleteRow(nWorksheet, nRow);
				}
			}

			~ChartImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class CellImplementation
		{
			public Worksheet m_pWorksheet;
			public Value m_pValue;
			public Style m_pStyle;
			public Workbook GetWorkbook()
			{
				return m_pWorksheet.m_pImpl.GetWorkbook();
			}

			public void SetFormula(Formula pFormula)
			{
				m_pValue.m_pImpl.SetFormula(pFormula, m_pWorksheet);
			}

			~CellImplementation()
			{
			}

		}
	}
}

namespace NumberDuck
{
	namespace Secret
	{
		class OwnedVector<T> where T : class
		{
			protected Vector<T> m_pVector;
			public OwnedVector()
			{
				m_pVector = new Vector<T>();
			}

			~OwnedVector()
			{
				Clear();
			}

			public T PushBack(T xObject)
			{
				m_pVector.PushBack(xObject);
				return m_pVector.Get(m_pVector.GetSize() - 1);
			}

			public int GetSize()
			{
				return m_pVector.GetSize();
			}

			public T Get(int nIndex)
			{
				return m_pVector.Get(nIndex);
			}

			public void Clear()
			{
				while (m_pVector.GetSize() > 0)
				{
					T pTemp = m_pVector.PopBack();
					{
						pTemp = null;
					}
				}
			}

			public void Insert(int nIndex, T pObject)
			{
				m_pVector.Insert(nIndex, pObject);
			}

			public T Remove(int nIndex)
			{
				T pTemp = m_pVector.Get(nIndex);
				m_pVector.Erase(nIndex);
				return pTemp;
			}

			public void Erase(int nIndex)
			{
				{
				}
				m_pVector.Erase(nIndex);
			}

			public T PopBack()
			{
				return m_pVector.PopBack();
			}

			public T PopFront()
			{
				return m_pVector.PopFront();
			}

		}
	}
}

namespace NumberDuck
{
    namespace Secret
    {
        class nbAssert
        {
            public static void Assert(bool result, string test, string file, int line)
            {
                if (!result)
                    throw new System.Exception(test + "\n" + file + ":" + line);
            }

            public static void Assert(bool result)
            {
                Assert(result, "", "", 0);
            }
        }
    }
}

namespace NumberDuck
{
    class Blob
    {
        private const int DEFAULT_SIZE = 1024 * 32;

        internal byte[] m_pBuffer;
        internal int m_nSize;
        private bool m_bAutoResize;
        private BlobView m_pBlobView;

        public Blob(bool bAutoResize)
        {
            m_pBuffer = new byte[DEFAULT_SIZE];
            m_nSize = 0;
            m_bAutoResize = bAutoResize;
            m_pBlobView = new BlobView(this, 0, 0);
        }

        public bool Load(string fileName)
        {
            try
            {
                byte[] temp = System.IO.File.ReadAllBytes(fileName);
                m_pBuffer = temp;
            }
            catch
            {
                return false;
            }

            m_nSize = m_pBuffer.Length;
            m_pBlobView.m_nEnd = m_nSize;
            return true;
        }

        public bool Save(string fileName)
        {
            try
            {
                using (System.IO.FileStream fs = new System.IO.FileStream(fileName, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                {
                    fs.Write(m_pBuffer, 0, (int)m_nSize);
                }
            }
            catch //(Exception e)
            {
                return false;
            }
            return true;
        }

        public void Resize(int nSize, bool bAutoResize)
        {
            byte[] pOldBuffer = m_pBuffer;
            int nBufferSize = m_pBuffer.Length;

            if (bAutoResize)
                Secret.nbAssert.Assert(m_bAutoResize);

            if (nSize > nBufferSize)
            {
                while (nSize > nBufferSize)
                {
                    // if we are over 100mb, just use the target size, otherwise we'll blow out the RAMs
                    if (nSize > 1024 * 1024 * 100)
                        nBufferSize = nSize;
                    else
                        nBufferSize <<= 1;
                }

                m_pBuffer = new byte[nBufferSize];
                pOldBuffer.CopyTo(m_pBuffer, 0);
                pOldBuffer = null;
            }
            m_nSize = nSize;

            m_pBlobView.m_nEnd = m_nSize;

        }

        public int GetSize()
        {
            return m_nSize;
        }

        public System.IO.Stream CreateStream(int nStart, int nEnd)
        {
            return new System.IO.MemoryStream(m_pBuffer, nStart, nEnd - nStart);
        }

        public uint GetMsoCrc32()
        {
            // https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/324014d1-39aa-4038-bbf4-f3781732f767
            // https://github.com/chakannom/algorithm/blob/master/MsoCRC32Compute/MsoCRC32Compute/MsoCRC32Compute.cpp
            uint[] nCache = new uint[256];
            {
                uint i;
                for (i = 0; i < 256; i++)
                {
                    uint nBit;
                    uint nValue = i << 24;
                    for (nBit = 0; nBit < 8; nBit++)
                    {
                        if ((nValue & (0x1 << 31)) > 0)
                            nValue = (nValue << 1) ^ 0xAF;
                        else
                            nValue = nValue << 1;
                    }

                    nCache[i] = nValue & 0xFFFF;
                }
            }

            uint nCrcValue = 0;
            {
                uint i;
                for (i = 0; i < m_nSize; i++)
                {
                    uint nIndex = nCrcValue;
                    nIndex = nIndex >> 24;
                    nIndex = nIndex ^ m_pBuffer[i];
                    nCrcValue = nCrcValue << 8;
                    nCrcValue = nCrcValue ^ nCache[nIndex];
                    //printf("%x %lX \n", m_pBuffer[i], nCrcValue);
                }
            }

            return nCrcValue;
        }

        public void Md4Hash(BlobView pOut)
        {
            for (int i = 0; i < 16; i++)
                pOut.PackUint8(0);
            /*unsigned char nTemp[16];
            MD4_CTX ctx;
            MD4_Init(&ctx);
            MD4_Update(&ctx, (void*)m_pBuffer, m_nSize);
            MD4_Final(nTemp, &ctx);

            pOut->PackData(nTemp, 16);*/
        }


        /*public void Unpack(byte[] pTo, int nToOffset, int nFromOffset, int nSize)
		{
			Buffer.BlockCopy(pData, 0, m_pBuffer, (int)nOffset, (int)nSize);
		}*/


        public void UnpackData(byte[] pData, int nOffset, int nSize)
        {
            Secret.nbAssert.Assert(nOffset + nSize <= m_nSize);
            System.Buffer.BlockCopy(m_pBuffer, nOffset, pData, 0, nSize);
        }


        public void PackUint8(byte val, int nOffset) { m_pBuffer[nOffset] = val; }


        public short UnpackInt16(int nOffset) { Secret.nbAssert.Assert(nOffset + 2 <= m_nSize); return System.BitConverter.ToInt16(m_pBuffer, nOffset); }
        public int UnpackInt32(int nOffset) { Secret.nbAssert.Assert(nOffset + 4 <= m_nSize); return System.BitConverter.ToInt32(m_pBuffer, nOffset); }
        public byte UnpackUint8(int nOffset) { Secret.nbAssert.Assert(nOffset + 1 <= m_nSize); return m_pBuffer[nOffset]; }
        public ushort UnpackUint16(int nOffset) { Secret.nbAssert.Assert(nOffset + 2 <= m_nSize); return System.BitConverter.ToUInt16(m_pBuffer, nOffset); }
        public uint UnpackUint32(int nOffset) { Secret.nbAssert.Assert(nOffset + 4 <= m_nSize); return System.BitConverter.ToUInt32(m_pBuffer, nOffset); }
        public double UnpackDouble(int nOffset) { Secret.nbAssert.Assert(nOffset + 8 <= m_nSize); return System.BitConverter.ToDouble(m_pBuffer, nOffset); }


        public void PackData(byte[] pData, int nOffset, int nSize)
        {
            Secret.nbAssert.Assert(nOffset + nSize <= m_nSize);
            System.Buffer.BlockCopy(pData, 0, m_pBuffer, nOffset, nSize);
        }


        public void Pack(byte[] pData, int nDataOffset, int nOffset, int nSize)
        {
            Secret.nbAssert.Assert(nSize > 0);
            Secret.nbAssert.Assert(nOffset + nSize <= m_nSize);
            System.Buffer.BlockCopy(pData, nDataOffset, m_pBuffer, nOffset, nSize);
        }




        //m_pBlob->GetBlobView()->Pack(pBlobView, nSize);


        //public coid UnpackData()

        /*void Blob :: UnpackData(unsigned char* pData, unsigned int nOffset, unsigned int nSize)
		{
			CLIFFY_ASSERT(nOffset + nSize <= m_nSize);
			memcpy(pData, m_pBuffer + nOffset, nSize);
		}*/

        public BlobView GetBlobView()
        {
            return m_pBlobView;
        }

        public bool Equal(Blob pOther)
        {
            if (m_nSize != pOther.m_nSize)
                return false;

            for (int i = 0; i < m_nSize; i++)
                if (m_pBuffer[i] != pOther.m_pBuffer[i])
                    return false;
            return true;
        }
    }

    class BlobView
    {
        internal int m_nStart;
        internal int m_nEnd;
        internal int m_nOffset;
        internal Blob m_pBlob;


        public BlobView(Blob pBlob, int nStart, int nEnd)
        {
            Secret.nbAssert.Assert(nStart <= nEnd);
            Secret.nbAssert.Assert(nEnd <= pBlob.GetSize());

            m_pBlob = pBlob;
            m_nStart = nStart;
            m_nEnd = nEnd;
            m_nOffset = 0;
        }



        public Blob GetBlob()
        {
            return m_pBlob;
        }



        public int GetStart() { return m_nStart; }
        public int GetEnd() { return m_nEnd; }
        public int GetSize()
        {
            int nEnd = m_nEnd;
            if (nEnd == 0)
                nEnd = m_pBlob.GetSize();
            return nEnd - m_nStart;
        }

        public int GetOffset() { return m_nOffset; }

        public void SetOffset(int nOffset)
        {
            // todo: cap?
            m_nOffset = nOffset;
        }

        public System.IO.Stream CreateStream()
        {
            if (this == m_pBlob.GetBlobView())
                return m_pBlob.CreateStream(m_nStart + m_nOffset, m_nStart + m_pBlob.GetSize());
            return m_pBlob.CreateStream(m_nStart + m_nOffset, m_nEnd);
        }

        public void Pack(BlobView pBlobView, int nSize)
        {
            PackAt(m_nOffset, pBlobView, nSize);
            m_nOffset += nSize;
        }

        public void PackAt(int nOffset, BlobView pBlobView, int nSize)
        {
            byte[] pData = new byte[nSize];
            pBlobView.UnpackData(pData, nSize);
            PackDataAt(nOffset, pData, nSize);
        }

        public void PackDataAt(int nOffset, byte[] pData, int nSize)
        {
            int nBlobOffset = m_nStart + nOffset;
            if (this == m_pBlob.GetBlobView())
            {
                if (nBlobOffset + nSize > m_pBlob.GetSize())
                {
                    m_pBlob.Resize(nBlobOffset + nSize, true);
                    //this.m_nEnd = m_pBlob.GetSize();
                }
            }
            else
            {
                Secret.nbAssert.Assert(nBlobOffset + nSize <= m_nEnd);
            }
            m_pBlob.PackData(pData, nBlobOffset, nSize);
        }


        public void Pack(byte[] pData, int nDataSize)
        {
            PackDataAt(m_nOffset, pData, nDataSize);
            m_nOffset += nDataSize;
        }

        public void PackInt16(short val) { Pack(System.BitConverter.GetBytes(val), 2); }
        public void PackInt32(int val) { Pack(System.BitConverter.GetBytes(val), 4); }
        public void PackUint8(byte val) { Pack(System.BitConverter.GetBytes(val), 1); }
        public void PackUint16(ushort val) { Pack(System.BitConverter.GetBytes(val), 2); }
        public void PackUint32(uint val) { Pack(System.BitConverter.GetBytes(val), 4); }
        public void PackDouble(double val) { Pack(System.BitConverter.GetBytes(val), 8); }


        public void Unpack(BlobView pBlobView, int nSize)
        {
            UnpackAt(m_nOffset, pBlobView, nSize);
            m_nOffset += nSize;
        }

        public void UnpackAt(int nOffset, BlobView pBlobView, int nSize)
        {
            byte[] pData = new byte[nSize];
            UnpackDataAt(nOffset, pData, nSize);

            pBlobView.Pack(pData, nSize);
        }


        public void UnpackData(byte[] pData, int nSize)
        {
            UnpackDataAt(m_nOffset, pData, nSize);
            m_nOffset += nSize;
        }

        public void UnpackDataAt(int nOffset, byte[] pData, int nSize)
        {
            int nBlobOffset = m_nStart + nOffset;
            int nEnd = m_nEnd;
            if (nEnd == 0)
                nEnd = m_pBlob.GetSize();
            Secret.nbAssert.Assert(nBlobOffset + nSize <= nEnd);
            m_pBlob.UnpackData(pData, nBlobOffset, nSize);
        }


        /*public void Unpack(BlobView pBlobView, int nSize)
		{
			CliffyAssert.Assert(m_nStart + m_nOffset + nSize < m_nEnd);



			m_pBlob.Unpack(m_nStart + m_nOffset, nSize, pBlobView->m_pBlob, pBlobView.m_nStart + m_nOffset);

			//m_pBlob.UnpackAt
		}*/

        /*void BlobView :: Unpack(BlobView* pBlobView, unsigned int nSize)
		{
			UnpackAt(m_nOffset, pBlobView, nSize);
			m_nOffset += nSize;
		}

		void BlobView :: UnpackAt(unsigned int nOffset, BlobView* pBlobView, unsigned int nSize)
		{
			unsigned char* pData = new unsigned char[nSize];
			UnpackDataAt(nOffset, pData,nSize);

			pBlobView->PackData(pData, nSize);
			delete [] pData;
		}*/



        public short UnpackInt16() { short nTemp = m_pBlob.UnpackInt16(m_nStart + m_nOffset); m_nOffset += 2; return nTemp; }
        public int UnpackInt32() { int nTemp = m_pBlob.UnpackInt32(m_nStart + m_nOffset); m_nOffset += 4; return nTemp; }

        public int UnpackInt32At(int nOffset) { return m_pBlob.UnpackInt32(m_nStart + nOffset); }

        public byte UnpackUint8() { byte nTemp = m_pBlob.UnpackUint8(m_nStart + m_nOffset++); return nTemp; }
        public ushort UnpackUint16() { ushort nTemp = m_pBlob.UnpackUint16(m_nStart + m_nOffset); m_nOffset += 2; return nTemp; }
        public uint UnpackUint32() { uint nTemp = m_pBlob.UnpackUint32(m_nStart + m_nOffset); m_nOffset += 4; return nTemp; }
        public double UnpackDouble() { double fTemp = m_pBlob.UnpackDouble(m_nStart + m_nOffset); m_nOffset += 8; return fTemp; }






        public byte GetChecksum()
        {
            const int WIDTH = 8;
            const byte TOPBIT = (1 << (WIDTH - 1));
            const byte POLYNOMIAL = 0xD8;  /* 11011 followed by 0's */

            byte remainder = 0;
            int nEnd = m_nEnd;
            if (nEnd == 0)
                nEnd = m_pBlob.m_nSize;

            for (int currentByte = m_nStart; currentByte < nEnd; currentByte++)
            {
                remainder ^= m_pBlob.m_pBuffer[currentByte];
                for (int bit = 8; bit > 0; --bit)
                {
                    if ((remainder & TOPBIT) > 0)
                    {
                        remainder = (byte)((remainder << 1) ^ POLYNOMIAL);
                    }
                    else
                    {
                        remainder = (byte)(remainder << 1);
                    }
                }
            }
            return (remainder);
        }


    }
}


namespace NumberDuck
{
    namespace Secret
    {
        class Console
        {
            public static void Log(string sLog)
            {
                System.Console.WriteLine(sLog);
            }
        }
    }
}

namespace NumberDuck
{
	namespace Secret
	{
		class ExternalString
		{
			public static bool Equal(string szA, string szB)
			{
				return string.Equals(szA, szB);
			}

			public static int GetChecksum(string szString)
			{
				int nResult = 0xABC123;
				for (int i = 0; i < szString.Length; i++)
				{
					char c = szString[i];
					nResult = (nResult ^ c) << 1;
				}
				return nResult;
			}

			public static long hextol(string szString)
			{
				return long.Parse(szString, System.Globalization.NumberStyles.HexNumber);
			}

			public static int atoi(string szString)
			{
				return int.Parse(szString);
			}

			public static long atol(string szString)
			{
				return long.Parse(szString);
			}

			public static double atof(string szString)
			{
				return double.Parse(szString);
			}

		}
	}
}

namespace NumberDuck
{
    namespace Secret
    {
        class File
        {
            public static InternalString GetContents(string sxPath)
            {
                try
                {
                    return new InternalString(System.IO.File.ReadAllText(sxPath));
                }
                catch
                {

                }
                return null;
            }

            public static void PutContents(string sPath, string sContents)
            {
                System.IO.File.WriteAllText(sPath, sContents);
            }

            public static OwnedVector<InternalString> GetRecursiveFileVector(string sPath)
            {
                OwnedVector<InternalString> sFileVector = new OwnedVector<InternalString>();
                Vector<InternalString> sDirectoryVector = new Vector<InternalString>();

                sDirectoryVector.PushBack(new InternalString(sPath));

                while (sDirectoryVector.GetSize() > 0)
                {
                    string sDirectory = sDirectoryVector.PopBack().GetExternalString();

                    string[] sDirectories = System.IO.Directory.GetDirectories(sDirectory);
                    for (int i = 0; i < sDirectories.Length; i++)
                        sDirectoryVector.PushBack(new InternalString(sDirectories[i]));

                    string[] sFiles = System.IO.Directory.GetFiles(sDirectory);
                    for (int i = 0; i < sFiles.Length; i++)
                    {
                        string sFile = sFiles[i];
                        if (sFile.EndsWith(".nll") || sFile.EndsWith(".nll.def"))
                            sFileVector.PushBack(new InternalString(sFile));
                    }
                }

                return sFileVector;
            }

            public static InternalString GetFileDirectory(string sPath)
            {
                return new InternalString(System.IO.Path.GetDirectoryName(sPath));
            }

            public static void CreateDirectory(string sPath)
            {
                System.IO.Directory.CreateDirectory(sPath);
            }
        }
    }
}

namespace NumberDuck
{
	namespace Secret
	{
		class InternalString
		{
			private System.Text.StringBuilder m_pStringBuilder;

			public InternalString(string szString)
			{
				Set(szString);
			}

			public InternalString CreateClone()
			{
				return new InternalString(m_pStringBuilder.ToString());
			}

			public void Set(string szString)
			{
				m_pStringBuilder = new System.Text.StringBuilder(szString);
			}

			public string GetExternalString()
			{
				return m_pStringBuilder.ToString();
			}

			public void Append(string sString)
			{
				m_pStringBuilder.Append(sString);
			}

			public void AppendChar(char nChar)
			{
				m_pStringBuilder.Append(nChar);
			}

			public void AppendString(string szString)
			{
				m_pStringBuilder.Append(szString);
			}

			public void AppendInt(int nInt)
			{
				m_pStringBuilder.Append("" + nInt);
			}

			public void AppendUint32(uint nUint)
			{
				AppendUnsignedInt(nUint);
			}

			public void AppendUnsignedInt(uint nUint)
			{
				m_pStringBuilder.Append("" + nUint);
			}

			public void AppendDouble(double fDouble)
			{
				m_pStringBuilder.Append(fDouble.ToString("G6"));
			}

			public void PrependString(string sString)
			{
				m_pStringBuilder = new System.Text.StringBuilder(sString + m_pStringBuilder.ToString());
			}

			public void PrependChar(char cChar)
			{
				m_pStringBuilder = new System.Text.StringBuilder(cChar + m_pStringBuilder.ToString());
			}

			public void SubStr(int nStart, int nLength)
			{
				m_pStringBuilder = new System.Text.StringBuilder(m_pStringBuilder.ToString().Substring(nStart, nLength));
			}

			public void CropFront(int nLength)
			{
				m_pStringBuilder = new System.Text.StringBuilder(m_pStringBuilder.ToString().Substring(nLength));
			}

			public int GetLength()
			{
				return m_pStringBuilder.Length;
			}

			public char GetChar(int nIndex)
			{
				return m_pStringBuilder[nIndex];
			}

			public void BlobWriteUtf8(BlobView pBlobView, bool bZeroTerminator)
			{
				byte[] pData = System.Text.Encoding.UTF8.GetBytes(m_pStringBuilder.ToString());
				pBlobView.Pack(pData, pData.Length);
				if (bZeroTerminator)
					pBlobView.PackUint8(0);
			}

			public void BlobWrite16Bit(BlobView pBlobView, bool bZeroTerminator)
			{
				byte[] pData = System.Text.Encoding.Unicode.GetBytes(m_pStringBuilder.ToString());
				pBlobView.Pack(pData, pData.Length);
				if (bZeroTerminator)
					pBlobView.PackUint16(0);
			}

			public bool IsAscii()
			{
				return System.Text.Encoding.UTF8.GetByteCount(m_pStringBuilder.ToString()) == m_pStringBuilder.Length;
			}

			public bool IsEqual(string sString)
			{
				return m_pStringBuilder.ToString().Equals(sString);
			}

			public bool StartsWith(string szString)
			{
				return m_pStringBuilder.ToString().StartsWith(szString);
			}

			public bool EndsWith(string szString)
			{
				return m_pStringBuilder.ToString().EndsWith(szString);
			}

			public double ParseDouble()
			{
				return double.Parse(m_pStringBuilder.ToString());
			}

			public uint ParseHex()
			{
				string sTemp = m_pStringBuilder.ToString();
				if (!sTemp.StartsWith("0x", System.StringComparison.Ordinal))
					return 0;
				sTemp = sTemp.Substring(2);
				return uint.Parse(sTemp, System.Globalization.NumberStyles.HexNumber);
			}

			public int FindChar(char nChar)
			{
				return m_pStringBuilder.ToString().IndexOf((char)nChar);
			}

			public void Replace(string sxFind, string sxReplace)
			{
				m_pStringBuilder = new System.Text.StringBuilder(m_pStringBuilder.ToString().Replace(sxFind, sxReplace));
			}
		}
	}
}

namespace NumberDuck
{
    namespace Secret
    {
        class Utils
        {
            public static double Pow(double fBase, double fExponent)
            {
                return System.Math.Pow(fBase, fExponent);
            }

            public static double ByteConvertUint64ToDouble(ulong nValue)
            {
                byte[] nByteArray = System.BitConverter.GetBytes(nValue);
                return System.BitConverter.ToDouble(nByteArray, 0);
            }

            public static int ByteConvertUint32ToInt32(uint nValue)
            {
                byte[] nByteArray = System.BitConverter.GetBytes(nValue);
                return System.BitConverter.ToInt32(nByteArray, 0);
            }

            public static uint ByteConvertInt32ToUint32(int nValue)
            {
                byte[] nByteArray = System.BitConverter.GetBytes(nValue);
                return System.BitConverter.ToUInt32(nByteArray, 0);
            }

            public static void Indent(int nTabDepth, InternalString sOut)
            {
                for (int i = 0; i < nTabDepth; i++)
                    sOut.AppendChar('\t');
            }
        }
    }
}


namespace NumberDuck
{
	namespace Secret
	{

		class Vector<T>
		{
			private System.Collections.Generic.List<T> m_pList;


			public Vector()
			{
				Clear();
			}

			public void PushFront(T xObject)
			{
				m_pList.Insert(0, xObject);
			}

			public void PushBack(T xObject)
			{
				m_pList.Add(xObject);
			}

			public int GetSize()
			{
				return m_pList.Count;
			}

			public T Get(int nIndex)
			{
				return m_pList[nIndex];
			}

			public void Clear()
			{
				m_pList = new System.Collections.Generic.List<T>();
			}

			public void Set(int nIndex, T xObject)
			{
				m_pList[nIndex] = xObject;
			}

			public void Insert(int nIndex, T xObject)
			{
				m_pList.Insert(nIndex, xObject);
			}

			public void Erase(int nIndex)
			{
				m_pList.RemoveAt(nIndex);
			}

			public T PopBack()
			{
				int nIndex = m_pList.Count - 1;
				T xBack = m_pList[nIndex];
				m_pList.RemoveAt(nIndex);
				return xBack;
			}

			public T PopFront()
			{
				T xFront = m_pList[0];
				m_pList.RemoveAt(0);
				return xFront;
			}
		}
	}
}

﻿
namespace NumberDuck
{
    namespace Secret
    {
        class XmlFile : XmlNode
        {
            public XmlFile() : base(null)
            {
            }

            public bool Load(BlobView pBlobView)
            {
                try
                {
                    System.Xml.XmlDocument document = new System.Xml.XmlDocument();
                    document.Load(pBlobView.CreateStream());
                    m_pNode = document;
                }
                catch (System.Exception)
                {
                    return false;
                }
                return true;
            }
        }
    }


}

﻿
namespace NumberDuck
{
    namespace Secret
    {
        public class XmlNode
        {
            protected System.Xml.XmlNode m_pNode;

            internal XmlNode(System.Xml.XmlNode pNode)
            {
                m_pNode = pNode;
            }

            public XmlNode GetFirstChildElement(string szName)
            {
                if (m_pNode.ChildNodes.Count > 0)
                {
                    if (szName == null)
                        return new XmlNode(m_pNode.ChildNodes[0]);

                    for (int i = 0; i < m_pNode.ChildNodes.Count; i++)
                    {
                        System.Xml.XmlNode pNode = m_pNode.ChildNodes[i];
                        if (pNode.Name.Equals(szName))
                            return new XmlNode(pNode);
                    }
                }
                return null;
            }

            public XmlNode GetNextSiblingElement(string szName)
            {
                if (m_pNode.NextSibling != null)
                    return new XmlNode(m_pNode.NextSibling);
                return null;
            }

            public string GetValue()
            {
                return m_pNode.Name;
            }

            public string GetText()
            {
                return m_pNode.InnerText;
            }


            public string GetAttribute(string szName)
            {
                for (int i = 0; i < m_pNode.Attributes.Count; i++)
                    if (m_pNode.Attributes[i].Name.Equals(szName))
                        return m_pNode.Attributes[i].Value;
                return null;
            }
        }
    }
}

﻿
namespace NumberDuck
{
    namespace Secret
    {
        class Zip
        {
            System.IO.Compression.ZipArchive m_pZipArchive = null;

            public Zip()
            {

            }

            public bool LoadBlobView(BlobView pBlobView)
            {
                System.IO.Stream stream = pBlobView.CreateStream();

                try
                {
                    m_pZipArchive = new System.IO.Compression.ZipArchive(stream);
                }
                catch
                {
                    m_pZipArchive = null;
                    return false;
                }
                return true;
            }

            public bool LoadFile(string szFileName)
            {
                Blob pBlob = new Blob(false);
                if (!pBlob.Load(szFileName))
                    return false;
                return LoadBlobView(pBlob.GetBlobView());
            }

            public int GetNumFile()
            {
                return m_pZipArchive.Entries.Count;
            }

            public ZipFileInfo GetFileInfo(int nIndex)
            {
                System.IO.Compression.ZipArchiveEntry entry = m_pZipArchive.Entries[nIndex];
                return new ZipFileInfo(entry.FullName, (int)entry.Length);
            }

            public bool ExtractFileByIndex(int nIndex, BlobView pOutBlobView)
            {
                try
                {
                    System.IO.Compression.ZipArchiveEntry entry = m_pZipArchive.Entries[nIndex];
                    System.IO.Stream stream = entry.Open();

                    byte[] buffer = new byte[entry.Length];
                    stream.Read(buffer, 0, (int)entry.Length);
                    pOutBlobView.PackDataAt(pOutBlobView.GetOffset(), buffer, (int)entry.Length);
                }
                catch
                {
                    return false;
                }
                return true;
            }

            public bool ExtractFileByName(string szFileName, BlobView pOutBlobView)
            {
                for (int i = 0; i < m_pZipArchive.Entries.Count; i++)
                {
                    if (m_pZipArchive.Entries[i].FullName.Equals(szFileName))
                        return ExtractFileByIndex(i, pOutBlobView);
                }
                return false;
            }

            public bool ExtractFileByIndexToString(int nIndex, InternalString sOut)
            {
                try
                {
                    System.IO.Compression.ZipArchiveEntry entry = m_pZipArchive.Entries[nIndex];
                    System.IO.Stream stream = entry.Open();

                    System.IO.StreamReader reader = new System.IO.StreamReader(stream);
                    sOut.Append(reader.ReadToEnd());
                }
                catch
                {
                    return false;
                }
                return true;
            }


        }
    }
}


﻿
namespace NumberDuck
{
    namespace Secret
    {
        public class ZipFileInfo
        {
            private string m_sFileName;
            private int m_nSize;
            //private uint m_nCrc32;

            internal ZipFileInfo(string sFileName, int nSize /*, uint nCrc32*/)
            {
                m_sFileName = sFileName;
                m_nSize = nSize;
                //m_nCrc32 = nCrc32;
            }

            public string GetFileName()
            {
                return m_sFileName;
            }

            public int GetSize()
            {
                return m_nSize;
            }

            /*public uint GetCrc32()
            {
                return m_nCrc32;
            }*/
        }
    }
}

