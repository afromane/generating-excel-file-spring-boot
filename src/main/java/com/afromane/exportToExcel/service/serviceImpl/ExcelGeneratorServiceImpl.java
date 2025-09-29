package com.afromane.exportToExcel.service.serviceImpl;

import com.afromane.exportToExcel.model.PeopleReviewData;
import com.afromane.exportToExcel.service.ExcelGeneratorService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ExcelGeneratorServiceImpl implements ExcelGeneratorService {

    @Override

    public ByteArrayInputStream generatePeopleReviewReport(List<PeopleReviewData> dataList) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("People Review");

        // Set default column widths (14 columns: A to N)
        for (int i = 0; i < 14; i++) {
            sheet.setColumnWidth(i, 256 * 12); // 12 characters width
        }

        // Create styles
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dateStyle = createDateStyle(workbook);
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle subHeaderStyle = createSubHeaderStyle(workbook);
        CellStyle commentHeaderStyle = createCommentHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);

        // Row 0: Title
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("People Review ");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13)); // A1 to N1

        // Row 1: Date (in column N - index 13)
        Row dateRow = sheet.createRow(1);
        Cell dateCell = dateRow.createCell(13);
        dateCell.setCellValue("16/03/2023");
        dateCell.setCellStyle(dateStyle);

        // Row 2: Main headers
        Row mainHeaderRow = sheet.createRow(2);

        // Create individual main header cells
        String[] mainHeaders = {
                "Postes/occupants",
                "Date de Prise de Fonction",
                "Anciennete",
                "Potentiel d'Evolution",
                "Performance",
                "Statut"
        };

        for (int i = 0; i < mainHeaders.length; i++) {
            Cell cell = mainHeaderRow.createCell(i);
            cell.setCellValue(mainHeaders[i]);
            cell.setCellStyle(headerStyle);
        }

        // "Décision Talent Management" - spans columns G to M (indices 6-12)
        Cell decisionCell = mainHeaderRow.createCell(6);
        decisionCell.setCellValue("Décision Talent Management");
        decisionCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 6, 12));

        // Row 3: Subheaders
        Row subHeaderRow = sheet.createRow(3);

        // First 6 columns are empty (but styled)
        for (int i = 0; i < 6; i++) {
            Cell cell = subHeaderRow.createCell(i);
            cell.setCellStyle(subHeaderStyle);
        }

        // Decision Talent Management subheaders (columns G to N)
        String[] subHeaders = {
                "Revalorisation",
                "Promotion",
                "Prime Exceptionnelle",
                "Autres Avantages",
                "Formations",
                "Action Suivi pour 2024",
                "Commentaires"
        };

        for (int i = 0; i < subHeaders.length; i++) {
            Cell cell = subHeaderRow.createCell(6 + i);
            cell.setCellValue(subHeaders[i]);
            if (i == subHeaders.length - 1) { // Commentaires column
                cell.setCellStyle(commentHeaderStyle);
            } else {
                cell.setCellStyle(subHeaderStyle);
            }
        }

        // Add data rows starting from row 4
        int rowNum = 4;
        for (PeopleReviewData data : dataList) {
            Row dataRow = sheet.createRow(rowNum++);

            // Fill all 14 columns with data
            dataRow.createCell(0).setCellValue(data.getPostesOccupants());
            dataRow.createCell(1).setCellValue(data.getDatePriseFonction());
            dataRow.createCell(2).setCellValue(data.getAnciennete());
            dataRow.createCell(3).setCellValue(data.getPotentielEvolution());
            dataRow.createCell(4).setCellValue(data.getPerformance());
            dataRow.createCell(5).setCellValue(data.getStatut());
            dataRow.createCell(6).setCellValue(data.getRevalorisation());
            dataRow.createCell(7).setCellValue(data.getPromotion());
            dataRow.createCell(8).setCellValue(data.getPrimeExceptionnelle());
            dataRow.createCell(9).setCellValue(data.getAutresAvantages());
            dataRow.createCell(10).setCellValue(data.getFormations());
            dataRow.createCell(11).setCellValue(data.getActionSuivi2024());
            dataRow.createCell(12).setCellValue(data.getCommentaires());
            dataRow.createCell(13).setCellValue(""); // Empty cell for column N

            // Apply data style to all cells in the row
            for (int i = 0; i < 14; i++) {
                if (dataRow.getCell(i) != null) {
                    dataRow.getCell(i).setCellStyle(dataStyle);
                }
            }
        }

        // Auto-size columns for better readability
        for (int i = 0; i < 14; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write to byte array
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();

        return new ByteArrayInputStream(out.toByteArray());
    }

    private CellStyle createTitleStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        font.setItalic(true);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createDateStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    private CellStyle createSubHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    private CellStyle createCommentHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }

}
