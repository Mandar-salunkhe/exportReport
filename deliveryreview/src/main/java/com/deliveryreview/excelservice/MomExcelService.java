package com.deliveryreview.excelservice;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.deliveryreview.request.MomRequest;
import com.deliveryreview.utility.ImageUtil;

public class MomExcelService {
	private static final Logger logger = LogManager.getLogger(MomExcelService.class);

	public String exportMomReport(MomRequest momDetails) throws IOException {
		String status = "";
		logger.debug("Exporting Excel");
		Workbook workbook = null;
		FileOutputStream fileOut = null;
		File file = new File("Mom_Details_Sample.xlsx");

		try {
			workbook = new XSSFWorkbook();
			fileOut = new FileOutputStream(file);
			CreationHelper createHelper = workbook.getCreationHelper();
			Sheet momSheet = workbook.createSheet("Mom-Date.xlsx");

			try {
				
				Row momHeaderBlankRow = momSheet.createRow(0);
				CellStyle momHeaderBlankRowStyle = workbook.createCellStyle();
				setBGColor(workbook, momHeaderBlankRowStyle, new Color(255, 255, 255), FillPatternType.SOLID_FOREGROUND);
				Cell momHeaderBlankRowCell = momHeaderBlankRow.createCell(0);
				momHeaderBlankRowCell.setCellStyle(momHeaderBlankRowStyle);
				
				momSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 15));
				
				
				momSheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 2));

				CellStyle momMainStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momMainStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momMainStyle, new Color(184, 204, 228), FillPatternType.SOLID_FOREGROUND);
				momMainStyle.setAlignment(HorizontalAlignment.CENTER);
				momMainStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				momMainStyle.setWrapText(true);
				//setBorder(momMainStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
//				momMainStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
//				momMainStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
//				momMainStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());
//				momMainStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());

				CellStyle momSecondaryStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momSecondaryStyle, false, 11, IndexedColors.BLACK.getIndex());
				momSecondaryStyle.setAlignment(HorizontalAlignment.LEFT);
				momSecondaryStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				momSecondaryStyle.setWrapText(true);

				CellStyle momGreenBgStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momGreenBgStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momGreenBgStyle, new Color(146, 208, 80), FillPatternType.SOLID_FOREGROUND);
				//setBorder(momGreenBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				momGreenBgStyle.setAlignment(HorizontalAlignment.CENTER);
				momGreenBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				CellStyle momYellowBgStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momYellowBgStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momYellowBgStyle, new Color(255, 255, 0), FillPatternType.SOLID_FOREGROUND);
				//setBorder(momYellowBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				momYellowBgStyle.setAlignment(HorizontalAlignment.CENTER);
				momYellowBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				CellStyle momRedBgStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momRedBgStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momRedBgStyle, new Color(255, 0, 0), FillPatternType.SOLID_FOREGROUND);
				//setBorder(momRedBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				momRedBgStyle.setAlignment(HorizontalAlignment.CENTER);
				momRedBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				CellStyle momOrangeBgStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momOrangeBgStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momOrangeBgStyle, new Color(255, 165, 0), FillPatternType.SOLID_FOREGROUND);
				//setBorder(momOrangeBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				momOrangeBgStyle.setAlignment(HorizontalAlignment.CENTER);
				momOrangeBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				
				CellStyle momBrownBgStyle = workbook.createCellStyle();
				setFont(workbook.createFont(), momBrownBgStyle, false, 11, IndexedColors.BLACK.getIndex());
				setBGColor(workbook, momBrownBgStyle, new Color(148, 138, 84), FillPatternType.SOLID_FOREGROUND);
				//setBorder(momBrownBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
				momBrownBgStyle.setAlignment(HorizontalAlignment.CENTER);
				momBrownBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				/*
				 * Row momHeaderRow1 = momSheet.createRow((short) 1); Cell momHeaderCell1 =
				 * momHeaderRow1.createCell(1); momHeaderCell1.setCellStyle(momMainStyle); Cell
				 * momHeaderCell2 = momHeaderRow1.createCell(2);
				 * momHeaderCell2.setCellStyle(momMainStyle);
				 * 
				 * Row momHeaderRow2 = momSheet.createRow((short) 2); Cell momHeaderCell3 =
				 * momHeaderRow2.createCell(1); momHeaderCell3.setCellStyle(momMainStyle); Cell
				 * momHeaderCell4 = momHeaderRow2.createCell(2);
				 * momHeaderCell4.setCellStyle(momMainStyle);
				 */

				/*
				 * Row momHeaderRow3 = momSheet.createRow((short) 3); Cell momHeaderCell5 =
				 * momHeaderRow3.createCell(1); momHeaderCell5.setCellStyle(momMainStyle); Cell
				 * momHeaderCell6 = momHeaderRow3.createCell(2);
				 * momHeaderCell6.setCellStyle(momMainStyle);
				 */
				
				
				// Logo
				byte[] bytes = IOUtils.toByteArray(ImageUtil.getImage(ImageUtil.CLIENT_LOGO));
				int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
				CreationHelper helper = workbook.getCreationHelper();

				ClientAnchor anchor = helper.createClientAnchor();
				anchor.setCol1(1);
				anchor.setRow1(1);
				anchor.setRow2(2);
				Drawing<?> drawing = momSheet.createDrawingPatriarch();
				Picture pict = drawing.createPicture(anchor, pictureIdx);
				pict.resize(2);

				Row momHeaderRow1 = momSheet.createRow(1);
				Row momHeaderRow2 = momSheet.createRow(2);
				Row momHeaderRow3 = momSheet.createRow(3);

				Cell cell = momHeaderRow1.createCell(3);
				cell.setCellValue("Minutes Of Meetings");
				cell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(1, 3, 3, 10));

				Cell cell11 = momHeaderRow1.createCell(11);
				cell11.setCellValue("");
				cell11.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(1, 1, 11, 14));

				Cell momDateRow2Cell11 = momHeaderRow2.createCell(11);
				momDateRow2Cell11.setCellValue("Date");
				momDateRow2Cell11.setCellStyle(momMainStyle);
				Cell momDateCell12 = momHeaderRow2.createCell(12);
				momDateCell12.setCellValue(momDetails.getDate().toString());

				Cell momDateRow2Cell13 = momHeaderRow2.createCell(13);
				momDateRow2Cell13.setCellValue("");
				momDateRow2Cell13.setCellStyle(momMainStyle);

				Cell momDateRow2Cell14 = momHeaderRow2.createCell(14);
				momDateRow2Cell14.setCellValue("");
				momDateRow2Cell14.setCellStyle(momMainStyle);

				Cell momDateRow3Cell13 = momHeaderRow3.createCell(11);
				momDateRow3Cell13.setCellValue("");
				momDateRow3Cell13.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(3, 3, 11, 14));

				Row momHeaderRow4 = momSheet.createRow(4);
				momHeaderRow4.setZeroHeight(true);  
				Cell blankCell1 = momHeaderRow4.createCell(1);
				blankCell1.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(4, 4, 1, 14));

				int k = 5;

				Row momParticipantRow = momSheet.createRow(5);
				Cell participantBlankCell = momParticipantRow.createCell(3);
				participantBlankCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(5, 5, 3, 13));
				
				Cell participantHeaderCell1 = momParticipantRow.createCell(1);

				participantHeaderCell1.setCellValue("Participants");
				participantHeaderCell1.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(k, (k + momDetails.getParticipants().size()), 1, 2));

				Cell participantHeaderCell2 = momParticipantRow.createCell(14);
				participantHeaderCell2.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(k, (k + momDetails.getParticipants().size()), 14, 14));

				int j = 6;

				for (int i = 0; i < momDetails.getParticipants().size(); i++) {
					momParticipantRow = momSheet.createRow(j);
					Cell participantValueCell = momParticipantRow.createCell(3);
					participantValueCell.setCellValue(momDetails.getParticipants().get(i).split("::")[1]);
					momSheet.addMergedRegion(new CellRangeAddress(j, j, 3, 13));
					j++;
				}

				Row blankRow = momSheet.createRow(13);
				Cell blankRowCell = blankRow.createCell(1);
				blankRowCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(13, 14, 1, 14));

				Row pointDiscussedHeaderRow = momSheet.createRow(15);

				Cell pointDiscussedHeaderCell = pointDiscussedHeaderRow.createCell(1);
				pointDiscussedHeaderCell.setCellValue("Point Discussed");
				pointDiscussedHeaderCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(15, 17, 1, 2));

				Cell pointDiscussedValueCell = pointDiscussedHeaderRow.createCell(3);
				pointDiscussedValueCell.setCellValue(momDetails.getPointsDiscussed());
				pointDiscussedValueCell.setCellStyle(momSecondaryStyle);
				momSheet.addMergedRegion(new CellRangeAddress(15, 17, 3, 13));

				Cell pointDiscuusedBlankCell = pointDiscussedHeaderRow.createCell(14);
				pointDiscuusedBlankCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(15, 17, 14, 14));

				Row blankRow1 = momSheet.createRow(18);
				Cell blankRowCell1 = blankRow1.createCell(1);
				blankRowCell1.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(18, 19, 1, 14));

				Row actionItemHeaderRow = momSheet.createRow(20);

				Cell actionItemsCell = actionItemHeaderRow.createCell(1);
				actionItemsCell.setCellValue("Point Agreed/Action Items");
				actionItemsCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 1, 2));

				Cell actionItemsCell1 = actionItemHeaderRow.createCell(3);
				actionItemsCell1.setCellValue("Task Details");
				actionItemsCell1.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 3, 8));

				Cell actionItemsCell2 = actionItemHeaderRow.createCell(9);
				actionItemsCell2.setCellValue("Responsible");
				actionItemsCell2.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 9, 9));

				Cell actionItemsCell3 = actionItemHeaderRow.createCell(10);
				actionItemsCell3.setCellValue("Status");
				actionItemsCell3.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 10, 10));

				Cell actionItemsCell4 = actionItemHeaderRow.createCell(11);
				actionItemsCell4.setCellValue("Completion Date");
				actionItemsCell4.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 11, 11));

				Cell actionItemsCell5 = actionItemHeaderRow.createCell(12);
				actionItemsCell5.setCellValue("Remarks");
				actionItemsCell5.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 12, 13));

				Cell actionItemsBlankCell = actionItemHeaderRow.createCell(14);
				actionItemsBlankCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(20, 21, 14, 14));

				Row actionItemFullDateRow = momSheet.createRow(22);

				Cell actionItemFullDateRowCell = actionItemFullDateRow.createCell(1);
				actionItemFullDateRowCell.setCellValue(momDetails.getDate().toString());
				actionItemFullDateRowCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(22, 22, 1, 14));

				int actionItemTaskDetailsRowNumber = 23;
				Row actionItemTableDataRow;
				for (int i = 0; i < momDetails.getMomActionList().size(); i++) {
					actionItemTableDataRow = momSheet.createRow(actionItemTaskDetailsRowNumber);

					Cell actionItemValueCell = actionItemTableDataRow.createCell(1);
					actionItemValueCell.setCellStyle(momMainStyle);

					Cell actionItemValueCell0 = actionItemTableDataRow.createCell(2);
					actionItemValueCell0.setCellValue(i + 1);
					actionItemValueCell0.setCellStyle(momMainStyle);

					Cell actionItemValueCell1 = actionItemTableDataRow.createCell(3);
					actionItemValueCell1.setCellValue(momDetails.getMomActionList().get(i).getTaskDetails().toString());
					actionItemValueCell1.setCellStyle(momSecondaryStyle);
					momSheet.addMergedRegion(new CellRangeAddress(actionItemTaskDetailsRowNumber, actionItemTaskDetailsRowNumber, 3, 8));

					Cell actionItemValueCell2 = actionItemTableDataRow.createCell(9);
					List<String> responsibleList = new ArrayList<>();
					for (int init = 0; init < momDetails.getMomActionList().get(i).getResponsible().size(); init++) {
						responsibleList.add(momDetails.getMomActionList().get(i).getResponsible().get(init).split("::")[2]);
					}
					actionItemValueCell2.setCellValue(responsibleList.toString().replace("[", "").replace("]", ""));
					actionItemValueCell2.setCellStyle(momSecondaryStyle);
					

					Cell actionItemValueCell3 = actionItemTableDataRow.createCell(10);
					actionItemValueCell3.setCellValue(momDetails.getMomActionList().get(i).getStatus());
					if (momDetails.getMomActionList().get(i).getStatus().equals("Done")) {
						actionItemValueCell3.setCellStyle(momGreenBgStyle);
					} else if (momDetails.getMomActionList().get(i).getStatus().equals("WIP")) {
						actionItemValueCell3.setCellStyle(momYellowBgStyle);
					} else if (momDetails.getMomActionList().get(i).getStatus().equals("Pending")) {
						actionItemValueCell3.setCellStyle(momRedBgStyle);
					} else if (momDetails.getMomActionList().get(i).getStatus().equals("TBD")) {
						actionItemValueCell3.setCellStyle(momOrangeBgStyle);
					} else if (momDetails.getMomActionList().get(i).getStatus().equals("NA")) {
						actionItemValueCell3.setCellStyle(momBrownBgStyle);
					}

					Cell actionItemValueCell4 = actionItemTableDataRow.createCell(11);
					actionItemValueCell4.setCellValue(momDetails.getMomActionList().get(i).getCompletionDate());

					Cell actionItemValueCell5 = actionItemTableDataRow.createCell(12);
					actionItemValueCell5.setCellValue(momDetails.getMomActionList().get(i).getRemarks());
					momSheet.addMergedRegion(new CellRangeAddress(actionItemTaskDetailsRowNumber,actionItemTaskDetailsRowNumber, 12, 13));
					actionItemValueCell5.setCellStyle(momSecondaryStyle);

					Cell actionItemValueCell6 = actionItemTableDataRow.createCell(14);
					actionItemValueCell6.setCellStyle(momMainStyle);

					actionItemTaskDetailsRowNumber++;
				}
				Row actionItemBlankRow = momSheet.createRow(actionItemTaskDetailsRowNumber);
				Cell actionItemBlankRowCell = actionItemBlankRow.createCell(1);
				actionItemBlankRowCell.setCellStyle(momMainStyle);
				momSheet.addMergedRegion(new CellRangeAddress(actionItemTaskDetailsRowNumber, actionItemTaskDetailsRowNumber, 1, 14));
				
				momSheet.addMergedRegion(new CellRangeAddress(1, 100, 0, 0));
				momSheet.addMergedRegion(new CellRangeAddress(1, 100, 15, 15));
				
				Row momSheetBlankCellRow = momSheet.createRow(actionItemTaskDetailsRowNumber+1);
				Cell momSheetBlankCell = momSheetBlankCellRow.createCell(1);
				momSheetBlankCell.setCellStyle(momHeaderBlankRowStyle);
				momSheet.addMergedRegion(new CellRangeAddress(actionItemTaskDetailsRowNumber+1, 100, 1, 14));
				
				momSheet.setColumnWidth(9, 25 * 256);
				momSheet.setColumnWidth(11, 25 * 175);
				momSheet.setColumnWidth(12, 25 * 256);
				
			} catch (Exception ex) {
				logger.error(ex.getMessage(), ex);
			}
			workbook.write(fileOut);
		} catch (Exception ex) {
			logger.error(ex.getMessage(), ex);
		} finally {
			if (fileOut != null)
				fileOut.close();

			if (workbook != null) {
				workbook.close();
			}

		}
		return status;

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
