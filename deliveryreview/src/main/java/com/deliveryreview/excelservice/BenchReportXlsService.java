package com.deliveryreview.excelservice;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.deliveryreview.request.BenchReportRequest;

public class BenchReportXlsService {

	public  Map<Workbook, File> exportBenchReport(BenchReportRequest downloadXls, Workbook workbook, boolean isConsolidateReport) throws IOException {
		Map<Workbook, File> testMap = new HashedMap<Workbook, File>();
		File file = new File("");
		FileOutputStream fileOut = null;
		if (isConsolidateReport) {
			file = new File("Delivery_Review-01Dec_v0.1.xlsx");
		} else {

			file = new File("DR_Bench_Report.xlsx");

		}
		//File file = new File("DR_Bench_Report.xlsx");

		try {
			//workbook = new XSSFWorkbook();
			fileOut = new FileOutputStream(file);
			Sheet benchRepSheet = workbook.createSheet("Bench Report");

			try {

				CellStyle mainStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), mainStyle, false, 11, IndexedColors.BLACK.getIndex());
				mainStyle.setAlignment(HorizontalAlignment.CENTER);
				mainStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				mainStyle.setWrapText(true);
				setBorder(mainStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

				CellStyle tableHeaderStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), tableHeaderStyle, false, 11, IndexedColors.BLACK.getIndex());
				tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
				tableHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				tableHeaderStyle.setWrapText(true);
				setBorder(tableHeaderStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				setBGColor(workbook, tableHeaderStyle, new Color(147, 153, 151), FillPatternType.SOLID_FOREGROUND);

				CellStyle designsStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), designsStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, designsStyle, new Color(94, 219, 92), FillPatternType.SOLID_FOREGROUND);
				designsStyle.setAlignment(HorizontalAlignment.CENTER);
				designsStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				designsStyle.setWrapText(true);
				setBorder(designsStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);


				Row benchRepHeader1 = benchRepSheet.createRow(0);


				Cell roleHeader = benchRepHeader1.createCell(0);
				roleHeader.setCellValue("Roles");
				roleHeader.setCellStyle(mainStyle);
				benchRepSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
				benchRepSheet.autoSizeColumn(0);

				int cellNo = 0;
				for(int i = 0; i < downloadXls.getMonthList().size(); i++) {

					Cell headerValues = benchRepHeader1.createCell(cellNo+1);			
					headerValues.setCellValue(downloadXls.getMonthList().get(i));
					headerValues.setCellStyle(tableHeaderStyle);
					benchRepSheet.addMergedRegion(new CellRangeAddress(0, 0, cellNo+1, cellNo+2));
					benchRepSheet.autoSizeColumn(cellNo+1);
					cellNo = cellNo+2;
				}


				int currentRow = benchRepSheet.getLastRowNum();

				Row benchRepHeader2 = benchRepSheet.createRow(currentRow+1);

				int cellNo1 = 0;
				for(int i = 0; i < downloadXls.getMonthList().size(); i++) {

					Cell headerValues1 = benchRepHeader2.createCell(cellNo1+1);			
					headerValues1.setCellValue("Actual");
					headerValues1.setCellStyle(tableHeaderStyle);
					benchRepSheet.autoSizeColumn(cellNo1);
					Cell headerValues2 = benchRepHeader2.createCell(cellNo1+2);
					headerValues2.setCellValue("Normalised Bench");
					headerValues2.setCellStyle(tableHeaderStyle);
					cellNo1 = cellNo1+2;
				}

				int seRowNo = benchRepSheet.getLastRowNum()+1;
				Row seRow = benchRepSheet.createRow(seRowNo);
				Cell desigCell1 = seRow.createCell(0);
				desigCell1.setCellValue("Software Engg");
				desigCell1.setCellStyle(designsStyle);
				int seCell = 0;

				for(int i = 0; i < downloadXls.getSeList().size(); i++) {

					Cell seValues1 = seRow.createCell(seCell+1);			
					seValues1.setCellValue(downloadXls.getSeList().get(i).getValue());
					seValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(seCell);
					Cell seValues2 = seRow.createCell(seCell+2);
					seValues2.setCellValue(downloadXls.getSeList().get(i).getRatio());
					seValues2.setCellStyle(mainStyle);
					seCell = seCell+2;
				}

				int sseRowNo = benchRepSheet.getLastRowNum()+1;
				Row sseRow = benchRepSheet.createRow(sseRowNo);
				Cell desigCell2 = sseRow.createCell(0);
				desigCell2.setCellValue("Sen.Software Engg");
				desigCell2.setCellStyle(designsStyle);
				int sseCell = 0;

				for(int i = 0; i < downloadXls.getSseList().size(); i++) {

					Cell sseValues1 = sseRow.createCell(sseCell+1);			
					sseValues1.setCellValue(downloadXls.getSseList().get(i).getValue());
					sseValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(sseCell);
					Cell sseValues2 = sseRow.createCell(sseCell+2);
					sseValues2.setCellValue(downloadXls.getSseList().get(i).getRatio());
					sseValues2.setCellStyle(mainStyle);
					sseCell = sseCell+2;
				}

				int tsRowNo = benchRepSheet.getLastRowNum()+1;
				Row tsRow = benchRepSheet.createRow(tsRowNo);
				Cell desigCell3 = tsRow.createCell(0);
				desigCell3.setCellValue("Technical Specialist");
				desigCell3.setCellStyle(designsStyle);
				int tsCell = 0;

				for(int i = 0; i < downloadXls.getTsList().size(); i++) {

					Cell tsValues1 = tsRow.createCell(tsCell+1);			
					tsValues1.setCellValue(downloadXls.getTsList().get(i).getValue());
					tsValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(tsCell);
					Cell tsValues2 = tsRow.createCell(tsCell+2);
					tsValues2.setCellValue(downloadXls.getTsList().get(i).getRatio());
					tsValues2.setCellStyle(mainStyle);
					tsCell = tsCell+2;
				}

				int stsRowNo = benchRepSheet.getLastRowNum()+1;
				Row stsRow = benchRepSheet.createRow(stsRowNo);
				Cell desigCell4 = stsRow.createCell(0);
				desigCell4.setCellValue("Sen.Technical Specialist");
				desigCell4.setCellStyle(designsStyle);
				int stsCell = 0;

				for(int i = 0; i < downloadXls.getStsList().size(); i++) {

					Cell stsValues1 = stsRow.createCell(stsCell+1);			
					stsValues1.setCellValue(downloadXls.getStsList().get(i).getValue());
					stsValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(stsCell);
					Cell stsValues2 = stsRow.createCell(stsCell+2);
					stsValues2.setCellValue(downloadXls.getStsList().get(i).getRatio());
					stsValues2.setCellStyle(mainStyle);
					stsCell = stsCell+2;
				}

				int taRowNo = benchRepSheet.getLastRowNum()+1;
				Row taRow = benchRepSheet.createRow(taRowNo);
				Cell desigCell5 = taRow.createCell(0);
				desigCell5.setCellValue("Technical Architect");
				desigCell5.setCellStyle(designsStyle);
				int taCell = 0;

				for(int i = 0; i < downloadXls.getTaList().size(); i++) {

					Cell taValues1 = taRow.createCell(taCell+1);			
					taValues1.setCellValue(downloadXls.getTaList().get(i).getValue());
					taValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(taCell);
					Cell taValues2 = taRow.createCell(taCell+2);
					taValues2.setCellValue(downloadXls.getTaList().get(i).getRatio());
					taValues2.setCellStyle(mainStyle);
					taCell = taCell+2;
				}

				int staRowNo = benchRepSheet.getLastRowNum()+1;
				Row staRow = benchRepSheet.createRow(staRowNo);
				Cell desigCell6 = staRow.createCell(0);
				desigCell6.setCellValue("Sen.Technical Architect");
				desigCell6.setCellStyle(designsStyle);
				int staCell = 0;

				for(int i = 0; i < downloadXls.getStaList().size(); i++) {

					Cell staValues1 = staRow.createCell(staCell+1);			
					staValues1.setCellValue(downloadXls.getStaList().get(i).getValue());
					staValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(staCell);
					Cell staValues2 = staRow.createCell(staCell+2);
					staValues2.setCellValue(downloadXls.getStaList().get(i).getRatio());
					staValues2.setCellStyle(mainStyle);
					staCell = staCell+2;
				}

				int ptaRowNo = benchRepSheet.getLastRowNum()+1;
				Row ptaRow = benchRepSheet.createRow(ptaRowNo);
				Cell desigCell7 = ptaRow.createCell(0);
				desigCell7.setCellValue("Princ.Technical Architect");
				desigCell7.setCellStyle(designsStyle);
				int ptaCell = 0;

				for(int i = 0; i < downloadXls.getPtaList().size(); i++) {

					Cell ptaValues1 = ptaRow.createCell(ptaCell+1);			
					ptaValues1.setCellValue(downloadXls.getPtaList().get(i).getValue());
					ptaValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(ptaCell);
					Cell ptaValues2 = ptaRow.createCell(ptaCell+2);
					ptaValues2.setCellValue(downloadXls.getPtaList().get(i).getRatio());
					ptaValues2.setCellStyle(mainStyle);
					ptaCell = ptaCell+2;
				}

				int aptaRowNo = benchRepSheet.getLastRowNum()+1;
				Row aptaRow = benchRepSheet.createRow(aptaRowNo);
				Cell desigCell8 = aptaRow.createCell(0);
				desigCell8.setCellValue("Asso.Princ.Technical Architect");
				desigCell8.setCellStyle(designsStyle);
				int aptaCell = 0;

				for(int i = 0; i < downloadXls.getAptaList().size(); i++) {

					Cell aptaValues1 = aptaRow.createCell(aptaCell+1);			
					aptaValues1.setCellValue(downloadXls.getAptaList().get(i).getValue());
					aptaValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(aptaCell);
					Cell aptaValues2 = aptaRow.createCell(aptaCell+2);
					aptaValues2.setCellValue(downloadXls.getAptaList().get(i).getRatio());
					aptaValues2.setCellStyle(mainStyle);
					aptaCell = aptaCell+2;
				}

				int totalRowNo = benchRepSheet.getLastRowNum()+1;
				Row totalRow = benchRepSheet.createRow(totalRowNo);
				Cell desigCell9 = totalRow.createCell(0);
				desigCell9.setCellValue("Total");
				desigCell9.setCellStyle(mainStyle);
				int totalCell = 0;

				for(int i = 0; i < downloadXls.getTotalList().size(); i++) {

					Cell totalValues1 = totalRow.createCell(totalCell+1);			
					totalValues1.setCellValue(downloadXls.getTotalList().get(i).getValue());
					totalValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(totalCell);
					Cell totalValues2 = totalRow.createCell(totalCell+2);
					totalValues2.setCellValue(downloadXls.getTotalList().get(i).getRatio());
					totalValues2.setCellStyle(mainStyle);
					totalCell = totalCell+2;
				}

				int simpAvgRowNo = benchRepSheet.getLastRowNum()+1;
				Row simpAvgRow = benchRepSheet.createRow(simpAvgRowNo);
				Cell desigCell10 = simpAvgRow.createCell(0);
				desigCell10.setCellValue("Simple Average");
				desigCell10.setCellStyle(mainStyle);
				int simpAvgCell = 0;

				for(int i = 0; i < downloadXls.getSimpleAverageList().size(); i++) {

					Cell simpAvgValues1 = simpAvgRow.createCell(simpAvgCell+1);			
					simpAvgValues1.setCellValue(downloadXls.getSimpleAverageList().get(i).getValue()+"%");
					simpAvgValues1.setCellStyle(mainStyle);
					benchRepSheet.autoSizeColumn(simpAvgCell);
					Cell simpAvgValues2 = simpAvgRow.createCell(simpAvgCell+2);
					simpAvgValues2.setCellValue(downloadXls.getSimpleAverageList().get(i).getRatio()+"%");
					simpAvgValues2.setCellStyle(designsStyle);
					simpAvgCell = simpAvgCell+2;
				}

			}catch(Exception ex) {
				ex.printStackTrace();
			}
			
			testMap.put(workbook, file);

			if (isConsolidateReport) {
				workbook.write(fileOut);
				return testMap;
			} else {
				workbook.write(fileOut);
			}
			testMap.put(workbook, file);
			//workbook.write(fileOut);

		}catch(Exception ex) {
			ex.printStackTrace();
			return testMap;
		}finally {
			if (isConsolidateReport) {

			} else {
				if (fileOut != null)
					fileOut.close();

				if (workbook != null) {
					workbook.close();
				}
			}
		} 



		return testMap;
	}

	protected static void setMerge(Sheet sheet, int numRow, int untilRow, int numCol, int untilCol, boolean border) {
		CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
		sheet.addMergedRegion(cellMerge);
		if (border) {
			setBordersToMergedCells(sheet, cellMerge);
		}
	}

	protected static void setBordersToMergedCells(Sheet sheet, CellRangeAddress rangeAddress) {
		RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
		RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
	}

	protected static void setBGColor(Workbook workbook, CellStyle cellStyle, Color color, FillPatternType fp) {
		byte[] rgb = new byte[] { (byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue() };
		if (cellStyle instanceof XSSFCellStyle) {
			XSSFCellStyle xssfreportHeaderCellStyle = (XSSFCellStyle) cellStyle;
			xssfreportHeaderCellStyle.setFillForegroundColor(new XSSFColor(rgb, null));
		} else if (cellStyle instanceof HSSFCellStyle) {
			cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIME.getIndex());
			HSSFWorkbook hssfworkbook = (HSSFWorkbook) workbook;
			HSSFPalette palette = hssfworkbook.getCustomPalette();
			palette.setColorAtIndex(HSSFColor.HSSFColorPredefined.LIME.getIndex(), rgb[0], rgb[1], rgb[2]);
		}
		cellStyle.setFillPattern(fp);

	}

	protected static void setFont(Font font, CellStyle cellStyle, boolean isBold, int height, short color) {
		font.setBold(isBold);
		font.setFontHeightInPoints((short) height);
		font.setColor(color);
		cellStyle.setFont(font);
	}

	protected static void setBorder(CellStyle cellStyle, BorderStyle left, BorderStyle top, BorderStyle right,
			BorderStyle bottom) {
		cellStyle.setBorderLeft(left);
		cellStyle.setBorderTop(top);
		cellStyle.setBorderRight(right);
		cellStyle.setBorderBottom(bottom);
	}

}
