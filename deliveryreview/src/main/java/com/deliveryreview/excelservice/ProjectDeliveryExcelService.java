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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.deliveryreview.request.ProjectDeliveryStatRequest;

public class ProjectDeliveryExcelService {

	static CellStyle momYellowBgStyle;

	public Map<Workbook, File> exportProjectDelivReport(List<ProjectDeliveryStatRequest> projectDeliveryDetails,
			boolean isConsolidateReport) throws IOException {
		Map<Workbook, File> projectMap = new HashedMap<Workbook, File>();
		FileOutputStream fileOut = null;
		Workbook workbook = null;
		File file = new File("");
		Date todaysDate = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
		String strDate = formatter.format(todaysDate);
		if (isConsolidateReport) {
			file = new File("Delivery_Review-" + strDate + "_v0.1.xlsx");

		} else {
			file = new File("DR_Project_Delivery_Review.xlsx");
		}

		try {
			fileOut = new FileOutputStream(file);
			workbook = new XSSFWorkbook();

			CellStyle mainStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), mainStyle, false, 11, IndexedColors.BLACK.getIndex());
			mainStyle.setAlignment(HorizontalAlignment.CENTER);
			mainStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			mainStyle.setWrapText(true);
			setBorder(mainStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

			CellStyle mainRedStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), mainRedStyle, false, 11, IndexedColors.RED.getIndex());
			mainRedStyle.setAlignment(HorizontalAlignment.CENTER);
			mainRedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			mainRedStyle.setWrapText(true);
			setBorder(mainRedStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

			CellStyle mainGreenStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), mainGreenStyle, false, 11, IndexedColors.GREEN.getIndex());
			mainGreenStyle.setAlignment(HorizontalAlignment.CENTER);
			mainGreenStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			mainGreenStyle.setWrapText(true);
			setBorder(mainGreenStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

			CellStyle mainAmberStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), mainAmberStyle, false, 11, IndexedColors.ORANGE.getIndex());
			mainAmberStyle.setAlignment(HorizontalAlignment.CENTER);
			mainAmberStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			mainAmberStyle.setWrapText(true);
			setBorder(mainAmberStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

			CellStyle nameHeaderStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), nameHeaderStyle, true, 11, IndexedColors.BLACK.getIndex());
			nameHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
			nameHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			nameHeaderStyle.setWrapText(true);
			setBorder(nameHeaderStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);

			Sheet deliveryReviewSheet = workbook.createSheet("Delivery Review");

			String[] colurs = { "Green", "Amber", "Red" };
			String[] defination = { "Seamless Execution", "Close Watch", "Needs Attention" };

			for (int i = 0; i < 3; i++) {
				Row clrCodeRow = deliveryReviewSheet.createRow(i);
				Cell colurCodeCell1 = clrCodeRow.createCell(2);
				colurCodeCell1.setCellValue(colurs[i]);
				setBackground(colurs[i], workbook);
				colurCodeCell1.setCellStyle(momYellowBgStyle);
				Cell colurCodeCell2 = clrCodeRow.createCell(3);
				colurCodeCell2.setCellValue(defination[i]);
				setBackground(colurs[i], workbook);
				colurCodeCell2.setCellStyle(momYellowBgStyle);
				deliveryReviewSheet.autoSizeColumn(i + 2);
			}

			int tableHeaderInt = deliveryReviewSheet.getLastRowNum() + 2;
			Row tableHeader = deliveryReviewSheet.createRow(tableHeaderInt);
			int headCounter = 0;
			String[] deliveryMetrics = { "Area", "Goal" };

			for (int i = 0; i < 2; i++) {

				Cell header1 = tableHeader.createCell(i + 2);
				if (i == 0) {
					header1.setCellValue("Sr.no");
					deliveryReviewSheet
							.addMergedRegion(new CellRangeAddress(tableHeaderInt, tableHeaderInt + 1, i + 2, i + 2));
					header1.setCellStyle(nameHeaderStyle);
				} else {
					header1.setCellValue("Delivery Metrics");
					deliveryReviewSheet
							.addMergedRegion(new CellRangeAddress(tableHeaderInt, tableHeaderInt, i + 2, i + 3));
					header1.setCellStyle(nameHeaderStyle);
				}
				headCounter = i + 2;
			}

			int subHeaderInt = deliveryReviewSheet.getLastRowNum() + 1;
			Row subHeader = deliveryReviewSheet.createRow(subHeaderInt);

			for (int i = 0; i < deliveryMetrics.length; i++) {
				Cell header2 = subHeader.createCell(i + 3);
				header2.setCellValue(deliveryMetrics[i]);
				header2.setCellStyle(nameHeaderStyle);
				headCounter++;
			}

			int prjtBreif = headCounter;

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell header2 = tableHeader.createCell(headCounter);
				header2.setCellValue(projectDeliveryDetails.get(i).getClientName() + " - "
						+ projectDeliveryDetails.get(i).getProjectName());
				if (projectDeliveryDetails.get(i).getStatus().equals("Green")) {
					setBackground("Green", workbook);
					header2.setCellStyle(momYellowBgStyle);
				} else if (projectDeliveryDetails.get(i).getStatus().equals("Amber")) {
					setBackground("Amber", workbook);
					header2.setCellStyle(momYellowBgStyle);
				} else if (projectDeliveryDetails.get(i).getStatus().equals("Red")) {
					setBackground("Red", workbook);
					header2.setCellStyle(momYellowBgStyle);
				}
				deliveryReviewSheet.addMergedRegion(
						new CellRangeAddress(tableHeaderInt, tableHeaderInt + 1, headCounter, headCounter));
				headCounter++;
			}

			int prjectBriefInt = deliveryReviewSheet.getLastRowNum() + 1;
			Row prjectBrief = deliveryReviewSheet.createRow(prjectBriefInt);
			int csatHead = prjtBreif;

			Cell srNo = prjectBrief.createCell(prjtBreif - 3);
			srNo.setCellValue(1);
			srNo.setCellStyle(mainStyle);

			Cell paramName = prjectBrief.createCell(prjtBreif - 2);
			paramName.setCellValue("Project Brief");
			paramName.setCellStyle(mainStyle);

			Cell goal = prjectBrief.createCell(prjtBreif - 1);
			goal.setCellValue("");
			goal.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell header2 = prjectBrief.createCell(prjtBreif);
				header2.setCellValue(projectDeliveryDetails.get(i).getProjectBrief());
				header2.setCellStyle(mainStyle);
				deliveryReviewSheet.setColumnWidth(prjtBreif, 25 * 300);
				prjtBreif++;
			}

			int csatInt = deliveryReviewSheet.getLastRowNum() + 1;
			Row csatIntCol = deliveryReviewSheet.createRow(csatInt);
			int usage = csatHead;

			Cell srNo2 = csatIntCol.createCell(csatHead - 3);
			srNo2.setCellValue(2);
			srNo2.setCellStyle(mainStyle);

			Cell paramName2 = csatIntCol.createCell(csatHead - 2);
			paramName2.setCellValue("CSAT Details");
			paramName2.setCellStyle(mainStyle);

			Cell goal1 = csatIntCol.createCell(csatHead - 1);
			goal1.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell csatDetailsCell = csatIntCol.createCell(csatHead);
				csatDetailsCell.setCellValue(projectDeliveryDetails.get(i).getCsatDetails());
				if (projectDeliveryDetails.get(i).getCsatPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getCsatPriorityStatus().equals("Red")) {
						csatDetailsCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getCsatPriorityStatus().equals("Green")) {
						csatDetailsCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getCsatPriorityStatus().equals("Amber")) {
						csatDetailsCell.setCellStyle(mainAmberStyle);
					} else {
						csatDetailsCell.setCellStyle(mainStyle);
					}
				} else {
					csatDetailsCell.setCellStyle(mainStyle);
				}

				deliveryReviewSheet.setColumnWidth(csatHead, 25 * 300);
				csatHead++;
			}

			int row3Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row3Col = deliveryReviewSheet.createRow(row3Data);
			int row4 = usage;

			Cell srNo3 = row3Col.createCell(usage - 3);
			srNo3.setCellValue(3);
			srNo3.setCellStyle(mainStyle);

			Cell paramName3 = row3Col.createCell(usage - 2);
			paramName3.setCellValue("% age of projects using P-A-S™ Assurance Tool Kit :  ");
			paramName3.setCellStyle(mainStyle);

			Cell goal2 = row3Col.createCell(usage - 1);
			goal2.setCellValue("");
			goal2.setCellStyle(mainStyle);
			
			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell pasUsingPCTCell = row3Col.createCell(usage);
				pasUsingPCTCell.setCellValue(projectDeliveryDetails.get(i).getPasUsingPct());
				if (projectDeliveryDetails.get(i).getPasUsingPctPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getPasUsingPctPriorityStatus().equals("Red")) {
						pasUsingPCTCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getPasUsingPctPriorityStatus().equals("Green")) {
						pasUsingPCTCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getPasUsingPctPriorityStatus().equals("Amber")) {
						pasUsingPCTCell.setCellStyle(mainAmberStyle);
					} else {
						pasUsingPCTCell.setCellStyle(mainStyle);
					}
				} else {
					pasUsingPCTCell.setCellStyle(mainStyle);
				}

				deliveryReviewSheet.setColumnWidth(usage, 25 * 300);
				usage++;
			}

			int row4Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row4Col = deliveryReviewSheet.createRow(row4Data);
			int row5 = row4;

			Cell srNo4 = row4Col.createCell(row4 - 3);
			srNo4.setCellValue(3);
			srNo4.setCellStyle(mainStyle);

			Cell paramName4 = row4Col.createCell(row4 - 2);
			paramName4.setCellValue("% age customer reference-ability (new customers)");
			paramName4.setCellStyle(mainStyle);

			Cell goal3 = row4Col.createCell(row4 - 1);
			goal3.setCellValue("");
			goal3.setCellStyle(mainStyle);
			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell customerRefCell = row4Col.createCell(row4);
				
				customerRefCell.setCellValue(projectDeliveryDetails.get(i).getCustomerRefAbilityPct());
				if (projectDeliveryDetails.get(i).getCustomerRefAbilityPctPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getCustomerRefAbilityPctPriorityStatus().equals("Red")) {
						customerRefCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getCustomerRefAbilityPctPriorityStatus().equals("Green")) {
						customerRefCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getCustomerRefAbilityPctPriorityStatus().equals("Amber")) {
						customerRefCell.setCellStyle(mainAmberStyle);
					} else {
						customerRefCell.setCellStyle(mainStyle);
					}
				} else {
					customerRefCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row4, 25 * 300);
				row4++;
			}

			int row6Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row6Col = deliveryReviewSheet.createRow(row6Data);
			int row7 = row5;

			Cell srNo6 = row6Col.createCell(row5 - 3);
			srNo6.setCellValue(4);
			srNo6.setCellStyle(mainStyle);

			Cell paramName6 = row6Col.createCell(row5 - 2);
			paramName6.setCellValue("% age project artifacts submitted to AveKen/Techninja");
			paramName6.setCellStyle(mainStyle);

			Cell goal4 = row6Col.createCell(row5 - 1);
			goal4.setCellValue("");
			goal4.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell projectArtifactCell = row6Col.createCell(row5);
				projectArtifactCell.setCellValue(projectDeliveryDetails.get(i).getProjectArtifactsPct());
				if (projectDeliveryDetails.get(i).getProjectArtifactsPctPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getProjectArtifactsPctPriorityStatus().equals("Red")) {
						projectArtifactCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getProjectArtifactsPctPriorityStatus().equals("Green")) {
						projectArtifactCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getProjectArtifactsPctPriorityStatus().equals("Amber")) {
						projectArtifactCell.setCellStyle(mainAmberStyle);
					} else {
						projectArtifactCell.setCellStyle(mainStyle);
					}
				} else {
					projectArtifactCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row5, 25 * 300);
				row5++;
			}

			int row7Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row7Col = deliveryReviewSheet.createRow(row7Data);
			int row8 = row7;

			Cell srNo7 = row7Col.createCell(row7 - 3);
			srNo7.setCellValue(5);
			srNo7.setCellStyle(mainStyle);

			Cell paramName7 = row7Col.createCell(row7 - 2);
			paramName7.setCellValue("Qualilty artifacts submitted to AveKen/Techninja");
			paramName7.setCellStyle(mainStyle);

			Cell goal5 = row7Col.createCell(row7 - 1);
			goal5.setCellValue("");
			goal5.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell qualityArtifactCell = row7Col.createCell(row7);
				qualityArtifactCell.setCellValue(projectDeliveryDetails.get(i).getQualityOfArtifacts());
				if (projectDeliveryDetails.get(i).getQualityOfArtifactsPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getQualityOfArtifactsPriorityStatus().equals("Red")) {
						qualityArtifactCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getQualityOfArtifactsPriorityStatus().equals("Green")) {
						qualityArtifactCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getQualityOfArtifactsPriorityStatus().equals("Amber")) {
						qualityArtifactCell.setCellStyle(mainAmberStyle);
					} else {
						qualityArtifactCell.setCellStyle(mainStyle);
					}
				} else {
					qualityArtifactCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row7, 25 * 300);
				row7++;
			}

			int row8Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row8Col = deliveryReviewSheet.createRow(row8Data);
			int row9 = row8;

			Cell srNo8 = row8Col.createCell(row8 - 3);
			srNo8.setCellValue(6);
			srNo8.setCellStyle(mainStyle);

			Cell paramName8 = row8Col.createCell(row8 - 2);
			paramName8.setCellValue(
					"% age repeat business (downstream) from new clients (measured by the value of repeat business vs initial)");
			paramName8.setCellStyle(mainStyle);

			Cell goal6 = row8Col.createCell(row8 - 1);
			goal6.setCellValue("");
			goal6.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell repeatBuisnessCell = row8Col.createCell(row8);
				repeatBuisnessCell.setCellValue(projectDeliveryDetails.get(i).getRepeatBusinessPct());
				if (projectDeliveryDetails.get(i).getRepeatBusinessPctPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getRepeatBusinessPctPriorityStatus().equals("Red")) {
						repeatBuisnessCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getRepeatBusinessPctPriorityStatus().equals("Green")) {
						repeatBuisnessCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getRepeatBusinessPctPriorityStatus().equals("Amber")) {
						repeatBuisnessCell.setCellStyle(mainAmberStyle);
					} else {
						repeatBuisnessCell.setCellStyle(mainStyle);
					}
				} else {
					repeatBuisnessCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row8, 25 * 300);
				row8++;
			}

			int row9Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row9Col = deliveryReviewSheet.createRow(row9Data);
			int row10 = row9;

			Cell srNo9 = row9Col.createCell(row9 - 3);
			srNo9.setCellValue(7);
			srNo9.setCellStyle(mainStyle);

			Cell paramName9 = row9Col.createCell(row9 - 2);
			paramName9.setCellValue("Outcome based project");
			paramName9.setCellStyle(mainStyle);

			Cell goal7 = row9Col.createCell(row9 - 1);
			goal7.setCellValue("");
			goal7.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell outcomeBasedCell = row9Col.createCell(row9);
				outcomeBasedCell.setCellValue(projectDeliveryDetails.get(i).getOutcomeBased());
				if (projectDeliveryDetails.get(i).getOutcomeBasedPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getOutcomeBasedPriorityStatus().equals("Red")) {
						outcomeBasedCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getOutcomeBasedPriorityStatus().equals("Green")) {
						outcomeBasedCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getOutcomeBasedPriorityStatus().equals("Amber")) {
						outcomeBasedCell.setCellStyle(mainAmberStyle);
					} else {
						outcomeBasedCell.setCellStyle(mainStyle);
					}
				} else {
					outcomeBasedCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row9, 25 * 300);
				row9++;
			}

			int row10Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row10Col = deliveryReviewSheet.createRow(row10Data);
			int row11 = row10;

			Cell srNo10 = row10Col.createCell(row10 - 3);
			srNo10.setCellValue(8);
			srNo10.setCellStyle(mainStyle);

			Cell paramName10 = row10Col.createCell(row10 - 2);
			paramName10.setCellValue("Deliver FP projects with-in the planned budgeted efforts");
			paramName10.setCellStyle(mainStyle);

			Cell goal8 = row10Col.createCell(row10 - 1);
			goal8.setCellValue("");
			goal8.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell deliverFPCell = row10Col.createCell(row10);
				deliverFPCell.setCellValue(projectDeliveryDetails.get(i).getDeliverFP());
				if (projectDeliveryDetails.get(i).getDeliverFPPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getDeliverFPPriorityStatus().equals("Red")) {
						deliverFPCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getDeliverFPPriorityStatus().equals("Green")) {
						deliverFPCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getDeliverFPPriorityStatus().equals("Amber")) {
						deliverFPCell.setCellStyle(mainAmberStyle);
					} else {
						deliverFPCell.setCellStyle(mainStyle);
					}
				} else {
					deliverFPCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row10, 25 * 300);
				row10++;
			}

			int row11Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row11Col = deliveryReviewSheet.createRow(row11Data);
			int row12 = row11;

			Cell srNo11 = row11Col.createCell(row11 - 3);
			srNo11.setCellValue(9);
			srNo11.setCellStyle(mainStyle);

			Cell paramName11 = row11Col.createCell(row11 - 2);
			paramName11.setCellValue("Role ratio(target Vs Actuals) -Reduce the costing of each offerings");
			paramName11.setCellStyle(mainStyle);

			Cell goal9 = row11Col.createCell(row11 - 1);
			goal9.setCellValue("");
			goal9.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell roleRatioCell = row11Col.createCell(row11);
				roleRatioCell.setCellValue(projectDeliveryDetails.get(i).getRoleRatio());
				roleRatioCell.setCellStyle(mainStyle);
				deliveryReviewSheet.setColumnWidth(row11, 25 * 300);
				row11++;
			}

			int row14Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row14Col = deliveryReviewSheet.createRow(row14Data);
			int row15 = row12;

			Cell srNo14 = row14Col.createCell(row12 - 3);
			srNo14.setCellValue(10);
			srNo14.setCellStyle(mainStyle);

			Cell paramName14 = row14Col.createCell(row12 - 2);
			paramName14.setCellValue("Cost Reduction / improve profitability through automation ");
			paramName14.setCellStyle(mainStyle);

			Cell goal10 = row14Col.createCell(row12 - 1);
			goal10.setCellValue("");
			goal10.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell costReductionCell = row14Col.createCell(row12);
				costReductionCell.setCellValue(projectDeliveryDetails.get(i).getCostReductOrImproveProfit());
				if (projectDeliveryDetails.get(i).getCostReductOrImproveProfitPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getCostReductOrImproveProfitPriorityStatus().equals("Red")) {
						costReductionCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getCostReductOrImproveProfitPriorityStatus().equals("Green")) {
						costReductionCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getCostReductOrImproveProfitPriorityStatus().equals("Amber")) {
						costReductionCell.setCellStyle(mainAmberStyle);
					} else {
						costReductionCell.setCellStyle(mainStyle);
					}
				} else {
					costReductionCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row12, 25 * 300);
				row12++;
			}

			int row15Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row15Col = deliveryReviewSheet.createRow(row15Data);
			int row16 = row15;

			Cell srNo15 = row15Col.createCell(row15 - 3);
			srNo15.setCellValue(11);
			srNo15.setCellStyle(mainStyle);

			Cell paramName15 = row15Col.createCell(row15 - 2);
			paramName15.setCellValue("Ensure no risk in the project");
			paramName15.setCellStyle(mainStyle);

			Cell goal11 = row15Col.createCell(row15 - 1);
			goal11.setCellValue("");
			goal11.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell noRiskCell = row15Col.createCell(row15);
				noRiskCell.setCellValue(projectDeliveryDetails.get(i).getNoRiskInProject());
				if (projectDeliveryDetails.get(i).getNoRiskInProjectPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getNoRiskInProjectPriorityStatus().equals("Red")) {
						noRiskCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getNoRiskInProjectPriorityStatus().equals("Green")) {
						noRiskCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getNoRiskInProjectPriorityStatus().equals("Amber")) {
						noRiskCell.setCellStyle(mainAmberStyle);
					} else {
						noRiskCell.setCellStyle(mainStyle);
					}
				} else {
					noRiskCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row15, 25 * 300);
				row15++;
			}

			int row16Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row16Col = deliveryReviewSheet.createRow(row16Data);
			int row17 = row16;

			Cell srNo16 = row16Col.createCell(row16 - 3);
			srNo16.setCellValue(12);
			srNo16.setCellStyle(mainStyle);

			Cell paramName16 = row16Col.createCell(row16 - 2);
			paramName16.setCellValue("Innovation / Best practices -  created and implemented across the Organization");
			paramName16.setCellStyle(mainStyle);

			Cell goal12 = row16Col.createCell(row16 - 1);
			goal12.setCellValue("");
			goal12.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell innovationCell = row16Col.createCell(row16);
				innovationCell.setCellValue(projectDeliveryDetails.get(i).getInnovationOrBestPractices());
				if (projectDeliveryDetails.get(i).getInnovationOrBestPracticesPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getInnovationOrBestPracticesPriorityStatus().equals("Red")) {
						innovationCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getInnovationOrBestPracticesPriorityStatus().equals("Green")) {
						innovationCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getInnovationOrBestPracticesPriorityStatus().equals("Amber")) {
						innovationCell.setCellStyle(mainAmberStyle);
					} else {
						innovationCell.setCellStyle(mainStyle);
					}
				} else {
					innovationCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row16, 25 * 300);
				row16++;
			}

			int row17Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row17Col = deliveryReviewSheet.createRow(row17Data);
			int row18 = row17;

			Cell srNo17 = row17Col.createCell(row17 - 3);
			srNo17.setCellValue(13);
			srNo17.setCellStyle(mainStyle);

			Cell paramName17 = row17Col.createCell(row17 - 2);
			paramName17.setCellValue("Revenue leakage");
			paramName17.setCellStyle(mainStyle);

			Cell goal13 = row17Col.createCell(row17 - 1);
			goal13.setCellValue("");
			goal13.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell revenueLeakageCell = row17Col.createCell(row17);
				revenueLeakageCell.setCellValue(projectDeliveryDetails.get(i).getRevenueLeakage());
				if (projectDeliveryDetails.get(i).getRevenueLeakagePriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getRevenueLeakagePriorityStatus().equals("Red")) {
						revenueLeakageCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getRevenueLeakagePriorityStatus().equals("Green")) {
						revenueLeakageCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getRevenueLeakagePriorityStatus().equals("Amber")) {
						revenueLeakageCell.setCellStyle(mainAmberStyle);
					} else {
						revenueLeakageCell.setCellStyle(mainStyle);
					}
				} else {
					revenueLeakageCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row17, 25 * 300);
				row17++;
			}

			int row18Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row18Col = deliveryReviewSheet.createRow(row18Data);
			int row19 = row18;

			Cell srNo18 = row18Col.createCell(row18 - 3);
			srNo18.setCellValue(14);
			srNo18.setCellStyle(mainStyle);

			Cell paramName18 = row18Col.createCell(row18 - 2);
			paramName18.setCellValue("Employee Satisfaction");
			paramName18.setCellStyle(mainStyle);

			Cell goal14 = row18Col.createCell(row18 - 1);
			goal14.setCellValue("");
			goal14.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell employeeSatCell = row18Col.createCell(row18);
				employeeSatCell.setCellValue(projectDeliveryDetails.get(i).getEmployeeSatisfaction());
				if (projectDeliveryDetails.get(i).getEmployeeSatisfactionPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getEmployeeSatisfactionPriorityStatus().equals("Red")) {
						employeeSatCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getEmployeeSatisfactionPriorityStatus().equals("Green")) {
						employeeSatCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getEmployeeSatisfactionPriorityStatus().equals("Amber")) {
						employeeSatCell.setCellStyle(mainAmberStyle);
					} else {
						employeeSatCell.setCellStyle(mainStyle);
					}
				} else {
					employeeSatCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row18, 25 * 300);
				row18++;
			}

			int row20Data = deliveryReviewSheet.getLastRowNum() + 1;
			Row row20Col = deliveryReviewSheet.createRow(row20Data);

			Cell srNo20 = row20Col.createCell(row19-3);
			srNo20.setCellValue(15);
			srNo20.setCellStyle(mainStyle);

			Cell paramName20 = row20Col.createCell(row19-2);
			paramName20.setCellValue("Learnings");
			paramName20.setCellStyle(mainStyle);

			Cell goal15 = row20Col.createCell(row19-1);
			goal15.setCellValue("");
			goal15.setCellStyle(mainStyle);

			for (int i = 0; i < projectDeliveryDetails.size(); i++) {
				Cell learningCell = row20Col.createCell(row19);
				learningCell.setCellValue(projectDeliveryDetails.get(i).getLearnings());
				if (projectDeliveryDetails.get(i).getLearningsPriorityStatus() != null) {
					if (projectDeliveryDetails.get(i).getLearningsPriorityStatus().equals("Red")) {
						learningCell.setCellStyle(mainRedStyle);
					} else if (projectDeliveryDetails.get(i).getLearningsPriorityStatus().equals("Green")) {
						learningCell.setCellStyle(mainGreenStyle);
					} else if (projectDeliveryDetails.get(i).getLearningsPriorityStatus().equals("Amber")) {
						learningCell.setCellStyle(mainAmberStyle);
					} else {
						learningCell.setCellStyle(mainStyle);
					}
				} else {
					learningCell.setCellStyle(mainStyle);
				}
				deliveryReviewSheet.setColumnWidth(row19, 25 * 300);
				row19++;
			}

			workbook.write(fileOut);

			projectMap.put(workbook, file);

			if (isConsolidateReport) {
				workbook.write(fileOut);
				return projectMap;
			} else {
				workbook.write(fileOut);
			}
			projectMap.put(workbook, file);
		} catch (Exception ex) {
			ex.printStackTrace();
			return projectMap;
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
		return projectMap;
	}

	protected static void setBackground(String colour, Workbook workbook) {
		if (colour.equals("Green")) {
			momYellowBgStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), momYellowBgStyle, true, 11, IndexedColors.BLACK.getIndex());
			setBGColor(workbook, momYellowBgStyle, new Color(0, 128, 0), FillPatternType.SOLID_FOREGROUND);
			momYellowBgStyle.setAlignment(HorizontalAlignment.CENTER);
			momYellowBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			momYellowBgStyle.setWrapText(true);
			setBorder(momYellowBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
		} else if (colour.equals("Red")) {
			momYellowBgStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), momYellowBgStyle, true, 11, IndexedColors.BLACK.getIndex());
			setBGColor(workbook, momYellowBgStyle, new Color(255, 0, 0), FillPatternType.SOLID_FOREGROUND);
			momYellowBgStyle.setAlignment(HorizontalAlignment.CENTER);
			momYellowBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			momYellowBgStyle.setWrapText(true);
			setBorder(momYellowBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
		} else if (colour.equals("Amber")) {
			momYellowBgStyle = workbook.createCellStyle();
			setFont(workbook.createFont(), momYellowBgStyle, true, 11, IndexedColors.BLACK.getIndex());
			setBGColor(workbook, momYellowBgStyle, new Color(255, 165, 0), FillPatternType.SOLID_FOREGROUND);
			momYellowBgStyle.setAlignment(HorizontalAlignment.CENTER);
			momYellowBgStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			momYellowBgStyle.setWrapText(true);
			setBorder(momYellowBgStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
		}

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
