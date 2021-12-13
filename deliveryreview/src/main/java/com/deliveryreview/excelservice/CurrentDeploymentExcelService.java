package com.deliveryreview.excelservice;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
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
import org.json.JSONArray;

import com.deliveryreview.request.HeaderList;

public class CurrentDeploymentExcelService {

	public Map<Workbook, File> exportCurrDepReport(List<HeaderList> headerList, JSONArray activeConsultantsArray,
			JSONArray inActiveConsultantsRowsArray, JSONArray partnerEcoSystemRowsArray,
			JSONArray inActivePartnerEcoSystemRowsArray, boolean isConsolidateReport, Workbook workbook) throws IOException {

		Map<Workbook, File> currentDeploymentReportMap = new HashedMap<Workbook, File>();
		File file = new File("");
		
		FileOutputStream fileOut = null;
		Date todaysDate = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
		String strDate = formatter.format(todaysDate);
		if (isConsolidateReport) {
			file = new File("Delivery_Review-" + strDate + "_v0.1.xlsx");
		} else {

			file = new File("DR_Current_Deployment_Report.xlsx");

		}
		try {

			fileOut = new FileOutputStream(file);
			Sheet currDepSheet = workbook.createSheet("Current Deployment");

			try {

				CellStyle mainStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), mainStyle, false, 11, IndexedColors.BLACK.getIndex());
				mainStyle.setAlignment(HorizontalAlignment.CENTER);
				mainStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				mainStyle.setWrapText(true);
				setBorder(mainStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

				CellStyle tableHeaderStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), tableHeaderStyle, false, 11, IndexedColors.BLACK.getIndex());
				tableHeaderStyle.setWrapText(true);
				setBorder(tableHeaderStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				setBGColor(workbook, tableHeaderStyle, new Color(147, 153, 151), FillPatternType.SOLID_FOREGROUND);

				CellStyle billableStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), billableStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, billableStyle, new Color(94, 219, 92), FillPatternType.SOLID_FOREGROUND);
				billableStyle.setAlignment(HorizontalAlignment.CENTER);
				billableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				billableStyle.setWrapText(true);
				setBorder(billableStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

				CellStyle nonBillableStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), nonBillableStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, nonBillableStyle, new Color(229, 232, 28), FillPatternType.SOLID_FOREGROUND);
				nonBillableStyle.setAlignment(HorizontalAlignment.CENTER);
				nonBillableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				nonBillableStyle.setWrapText(true);
				setBorder(nonBillableStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

				CellStyle benchStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), benchStyle, false, 11, IndexedColors.RED.getIndex());
				benchStyle.setAlignment(HorizontalAlignment.CENTER);
				benchStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				benchStyle.setWrapText(true);
				setBorder(benchStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

				Row currDepRepHeader = currDepSheet.createRow(0);
				int srNo = 1;

				Cell srNoCol = currDepRepHeader.createCell(0);
				srNoCol.setCellValue("Sr.No");
				srNoCol.setCellStyle(mainStyle);
				for (int i = 0; i < headerList.size(); i++) {
					Cell headerValues = currDepRepHeader.createCell(i + 1);
					headerValues.setCellValue(headerList.get(i).getLabel());
					currDepSheet.autoSizeColumn(i + 1);
					headerValues.setCellStyle(mainStyle);
				}

				currDepSheet.createFreezePane(5, 1);

				int activeConsTable = currDepSheet.getLastRowNum() + 1;
				Row activeConsTableHeader = currDepSheet.createRow(activeConsTable);

				// Table for Active Consultants

				for (int i = 0; i < headerList.size(); i++) {
					Cell activeConstCell = activeConsTableHeader.createCell(i);
					activeConstCell.setCellValue("Section A : Active Consultants");
					activeConstCell.setCellStyle(tableHeaderStyle);
				}
				currDepSheet.addMergedRegion(
						new CellRangeAddress(activeConsTable, activeConsTable, 0, headerList.size() - 1));

				for (int i = 0; i < activeConsultantsArray.length(); i++) {
					int activeCons = currDepSheet.getLastRowNum() + 1;
					Row activeConsTableData = currDepSheet.createRow(activeCons);
					JSONArray resourceArr = activeConsultantsArray.getJSONArray(i);
					Cell activeConstResCellSrNo = activeConsTableData.createCell(0);
					activeConstResCellSrNo.setCellValue(srNo);
					activeConstResCellSrNo.setCellStyle(mainStyle);
					for (int j = 0; j < resourceArr.length(); j++) {
						Cell activeConstResCell = activeConsTableData.createCell(j + 1);

						if (resourceArr.getJSONObject(j).get("label").toString() == "null") {
							activeConstResCell.setCellValue("Still in Organization");
						} else {
							activeConstResCell.setCellValue(resourceArr.getJSONObject(j).get("label").toString());
						}
						if (resourceArr.getJSONObject(j).get("label").toString().equals("Billable")
								|| resourceArr.getJSONObject(j).get("label").toString().equals("Partially Billed")) {
							activeConstResCell.setCellStyle(billableStyle);
						} else if (resourceArr.getJSONObject(j).get("label").toString().equals("Non Billable")) {
							activeConstResCell.setCellStyle(nonBillableStyle);
						} else if (resourceArr.getJSONObject(j).get("label").toString().equals("Bench")) {
							activeConstResCell.setCellStyle(benchStyle);
						} else {
							activeConstResCell.setCellStyle(mainStyle);
						}

						currDepSheet.autoSizeColumn(j + 1);
					}
					srNo++;
				}

				// Table for Employee who has left this organisation in this period

				int inActiveConsTable = currDepSheet.getLastRowNum() + 3;
				Row inActiveConsTableHeader = currDepSheet.createRow(inActiveConsTable);
				for (int i = 0; i < headerList.size(); i++) {
					Cell inActiveConstCell = inActiveConsTableHeader.createCell(i);
					inActiveConstCell
							.setCellValue("Section B : Employee who has left this organisation in this period");
					inActiveConstCell.setCellStyle(tableHeaderStyle);
				}
				currDepSheet.addMergedRegion(
						new CellRangeAddress(inActiveConsTable, inActiveConsTable, 0, headerList.size() - 1));

				for (int i = 0; i < inActiveConsultantsRowsArray.length(); i++) {
					int inActiveCons = currDepSheet.getLastRowNum() + 1;
					Row inActiveConsTableData = currDepSheet.createRow(inActiveCons);
					JSONArray resourceArr1 = inActiveConsultantsRowsArray.getJSONArray(i);
					Cell InActiveConstResCellSrNo = inActiveConsTableData.createCell(0);
					InActiveConstResCellSrNo.setCellValue(srNo);
					InActiveConstResCellSrNo.setCellStyle(mainStyle);
					for (int j = 0; j < resourceArr1.length(); j++) {
						Cell InActiveConstResCell = inActiveConsTableData.createCell(j + 1);

						if (resourceArr1.getJSONObject(j).get("label").toString() == "null") {
							InActiveConstResCell.setCellValue("Still in Organization");
						} else {
							InActiveConstResCell.setCellValue(resourceArr1.getJSONObject(j).get("label").toString());
						}
						if (resourceArr1.getJSONObject(j).get("label").toString().equals("Billable")
								|| resourceArr1.getJSONObject(j).get("label").toString().equals("Partially Billed")) {
							InActiveConstResCell.setCellStyle(billableStyle);
						} else if (resourceArr1.getJSONObject(j).get("label").toString().equals("Non Billable")) {
							InActiveConstResCell.setCellStyle(nonBillableStyle);
						} else if (resourceArr1.getJSONObject(j).get("label").toString().equals("Bench")) {
							InActiveConstResCell.setCellStyle(benchStyle);
						} else {
							InActiveConstResCell.setCellStyle(mainStyle);
						}
						currDepSheet.autoSizeColumn(j + 1);
					}
					srNo++;
				}

				// Table for Partner EcoSystem

				int parActiveConsTable = currDepSheet.getLastRowNum() + 3;
				Row parActiveConsTableHeader = currDepSheet.createRow(parActiveConsTable);
				for (int i = 0; i < headerList.size(); i++) {
					Cell parActiveConstCell = parActiveConsTableHeader.createCell(i);
					parActiveConstCell.setCellValue("Section C : Partner EcoSystem");
					parActiveConstCell.setCellStyle(tableHeaderStyle);
				}
				currDepSheet.addMergedRegion(
						new CellRangeAddress(parActiveConsTable, parActiveConsTable, 0, headerList.size() - 1));

				for (int i = 0; i < partnerEcoSystemRowsArray.length(); i++) {
					int parActiveCons = currDepSheet.getLastRowNum() + 1;
					Row parActiveConsTableData = currDepSheet.createRow(parActiveCons);
					JSONArray resourceArr2 = partnerEcoSystemRowsArray.getJSONArray(i);
					Cell parActiveConstResCellSrNo = parActiveConsTableData.createCell(0);
					parActiveConstResCellSrNo.setCellValue(srNo);
					parActiveConstResCellSrNo.setCellStyle(mainStyle);
					for (int j = 0; j < resourceArr2.length(); j++) {
						Cell parActiveConstResCell = parActiveConsTableData.createCell(j + 1);

						if (resourceArr2.getJSONObject(j).get("label").toString() == "null") {
							parActiveConstResCell.setCellValue("Still in Organization");
						} else {
							parActiveConstResCell.setCellValue(resourceArr2.getJSONObject(j).get("label").toString());
						}
						if (resourceArr2.getJSONObject(j).get("label").toString().equals("Billable")
								|| resourceArr2.getJSONObject(j).get("label").toString().equals("Partially Billed")) {
							parActiveConstResCell.setCellStyle(billableStyle);
						} else if (resourceArr2.getJSONObject(j).get("label").toString().equals("Non Billable")) {
							parActiveConstResCell.setCellStyle(nonBillableStyle);
						} else if (resourceArr2.getJSONObject(j).get("label").toString().equals("Bench")) {
							parActiveConstResCell.setCellStyle(benchStyle);
						} else {
							parActiveConstResCell.setCellStyle(mainStyle);
						}
						currDepSheet.autoSizeColumn(j);
					}
					srNo++;
				}

				// Table for Employees from partner ecosystem who has left the organisation in
				// this period

				int parInActiveConsTable = currDepSheet.getLastRowNum() + 3;
				Row parInActiveConsTableHeader = currDepSheet.createRow(parInActiveConsTable);
				for (int i = 0; i < headerList.size(); i++) {
					Cell parInActiveConstCell = parInActiveConsTableHeader.createCell(i);
					parInActiveConstCell.setCellValue(
							"Section D : Employees from partner ecosystem who has left the organisation in this period");
					parInActiveConstCell.setCellStyle(tableHeaderStyle);
				}
				currDepSheet.addMergedRegion(
						new CellRangeAddress(parInActiveConsTable, parInActiveConsTable, 0, headerList.size() - 1));

				for (int i = 0; i < inActivePartnerEcoSystemRowsArray.length(); i++) {
					int inParActiveCons = currDepSheet.getLastRowNum() + 1;
					Row parInActiveConsTableData = currDepSheet.createRow(inParActiveCons);
					JSONArray resourceArr3 = inActivePartnerEcoSystemRowsArray.getJSONArray(i);
					Cell parInActiveConstResCellSrNo = parInActiveConsTableData.createCell(0);
					parInActiveConstResCellSrNo.setCellValue(srNo);
					parInActiveConstResCellSrNo.setCellStyle(mainStyle);
					for (int j = 0; j < resourceArr3.length(); j++) {
						Cell parInActiveConstResCell = parInActiveConsTableData.createCell(j + 1);

						if (resourceArr3.getJSONObject(j).get("label").toString() == "null") {
							parInActiveConstResCell.setCellValue("Still in Organization");
						} else {
							parInActiveConstResCell.setCellValue(resourceArr3.getJSONObject(j).get("label").toString());
						}
						if (resourceArr3.getJSONObject(j).get("label").toString().equals("Billable")
								|| resourceArr3.getJSONObject(j).get("label").toString().equals("Partially Billed")) {
							parInActiveConstResCell.setCellStyle(billableStyle);
						} else if (resourceArr3.getJSONObject(j).get("label").toString().equals("Non Billable")) {
							parInActiveConstResCell.setCellStyle(nonBillableStyle);
						} else if (resourceArr3.getJSONObject(j).get("label").toString().equals("Bench")) {
							parInActiveConstResCell.setCellStyle(benchStyle);
						} else {
							parInActiveConstResCell.setCellStyle(mainStyle);
						}
						currDepSheet.autoSizeColumn(j + 1);
					}

					srNo++;
				}

			} catch (Exception ex) {
				ex.printStackTrace();
			}

			currentDeploymentReportMap.put(workbook, file);

			if (isConsolidateReport) {
				workbook.write(fileOut);
				return currentDeploymentReportMap;
			} else {
				workbook.write(fileOut);
			}
			currentDeploymentReportMap.put(workbook, file);

		} catch (Exception ex) {
			ex.printStackTrace();
			return currentDeploymentReportMap;
		} finally {
			if (isConsolidateReport) {

			} else {
				if (fileOut != null)
					fileOut.close();

				if (workbook != null) {
					workbook.close();
				}
			}
		}
		return currentDeploymentReportMap;

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