package com.afromane.exportToExcel.service.serviceImpl;

import com.afromane.exportToExcel.model.PeopleReviewData;
import com.afromane.exportToExcel.service.MultiSheetPeopleReviewService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;

@Service
public class MultiSheetPeopleReviewServiceImpl implements MultiSheetPeopleReviewService {

    private static final String TEMPLATE_PATH = "templates/single.xlsx";

    @Override
    public ByteArrayInputStream generateMultiSheetReport(Map<String, List<PeopleReviewData>> sheetData) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            for (Map.Entry<String, List<PeopleReviewData>> entry : sheetData.entrySet()) {
                String sheetName = entry.getKey();
                List<PeopleReviewData> dataList = entry.getValue();

                Sheet sheet = cloneTemplateToSheet(workbook, sheetName);
                updateDate(sheet);
                updateTitle(sheet, sheetName);
                populateData(sheet, dataList, 6);
            }

            return writeWorkbookToByteArray(workbook);
        }
    }

    private Sheet cloneTemplateToSheet(Workbook destWorkbook, String sheetName) throws IOException {
        try (InputStream templateStream = new ClassPathResource(TEMPLATE_PATH).getInputStream();
             Workbook srcWorkbook = new XSSFWorkbook(templateStream)) {

            Sheet srcSheet = srcWorkbook.getSheetAt(0);
            Sheet destSheet = destWorkbook.createSheet(sheetName);

            // Copier les lignes
            for (int i = 0; i <= srcSheet.getLastRowNum(); i++) {
                Row srcRow = srcSheet.getRow(i);
                if (srcRow != null) {
                    Row destRow = destSheet.createRow(i);
                    copyRow(srcWorkbook, destWorkbook, srcRow, destRow);
                }
            }

            // Copier les merged regions
            for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
                destSheet.addMergedRegion(srcSheet.getMergedRegion(i));
            }

            return destSheet;
        }
    }

    private void copyRow(Workbook srcWorkbook, Workbook destWorkbook, Row srcRow, Row destRow) {
        destRow.setHeight(srcRow.getHeight());
        for (int i = 0; i < srcRow.getLastCellNum(); i++) {
            Cell srcCell = srcRow.getCell(i);
            if (srcCell != null) {
                Cell destCell = destRow.createCell(i);
                copyCell(srcWorkbook, destWorkbook, srcCell, destCell);
            }
        }
    }

    private void copyCell(Workbook srcWorkbook, Workbook destWorkbook, Cell srcCell, Cell destCell) {
        // ✅ Cloner le style dans le bon workbook
        CellStyle newStyle = cloneCellStyle(srcCell.getCellStyle(), srcWorkbook, destWorkbook);
        destCell.setCellStyle(newStyle);

        switch (srcCell.getCellType()) {
            case STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(srcCell.getCellFormula());
                break;
            case BLANK:
                // Rien à faire
                break;
            default:
                break;
        }
    }

    // ✅ Méthode corrigée : on passe srcWorkbook en paramètre
    private CellStyle cloneCellStyle(CellStyle srcStyle, Workbook srcWorkbook, Workbook destWorkbook) {
        CellStyle newStyle = destWorkbook.createCellStyle();
        newStyle.cloneStyleFrom(srcStyle);

        // Gérer la police (font) si nécessaire
        if (srcStyle instanceof XSSFCellStyle xssfSrcStyle) {
            // Récupérer la font du workbook source
            Font srcFont = srcWorkbook.getFontAt(xssfSrcStyle.getFontIndexAsInt());
            Font newFont = destWorkbook.createFont();

            // Copier les propriétés de la font
            newFont.setFontHeight(srcFont.getFontHeight());
            newFont.setFontName(srcFont.getFontName());
            newFont.setBold(srcFont.getBold());
            newFont.setItalic(srcFont.getItalic());
            newFont.setColor(srcFont.getColor());
            newFont.setUnderline(srcFont.getUnderline());
            newFont.setStrikeout(srcFont.getStrikeout());

            newStyle.setFont(newFont);
        }

        return newStyle;
    }

    private void updateDate(Sheet sheet) {
        Row dateRow = sheet.getRow(0);
        if (dateRow != null) {
            Cell dateCell = dateRow.getCell(14);
            if (dateCell != null) {
                String today = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                dateCell.setCellValue(today);
            }
        }
    }

    private void updateTitle(Sheet sheet, String directionName) {
        Row titleRow = sheet.getRow(1);
        if (titleRow != null) {
            Cell titleCell = titleRow.getCell(1);
            if (titleCell != null) {
                titleCell.setCellValue("People Review " + directionName);
            }
        }
    }

    private void populateData(Sheet sheet, List<PeopleReviewData> dataList, int startRowExcel) {
        int startIndex = startRowExcel - 1; // Excel ligne 6 → index 5

        // Nettoyer les anciennes données
        for (int i = sheet.getLastRowNum(); i >= startIndex; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }

        int currentRow = startIndex;
        for (PeopleReviewData data : dataList) {
            Row row = sheet.getRow(currentRow);
            if (row == null) {
                row = sheet.createRow(currentRow);
            }

            row.createCell(1).setCellValue(data.getPostesOccupants());
            row.createCell(2).setCellValue(data.getDatePriseFonction());
            row.createCell(3).setCellValue(data.getAnciennete());
            row.createCell(4).setCellValue(data.getPotentielEvolution());
            row.createCell(5).setCellValue(data.getPerformance());
            row.createCell(6).setCellValue(data.getStatut());
            row.createCell(7).setCellValue(data.getRevalorisation());
            row.createCell(8).setCellValue(data.getPromotion());
            row.createCell(9).setCellValue(data.getPrimeExceptionnelle());
            row.createCell(10).setCellValue(data.getAutresAvantages());
            row.createCell(11).setCellValue(data.getFormations());
            row.createCell(12).setCellValue(data.getActionSuivi2024());
            row.createCell(13).setCellValue(data.getCommentaires());

            currentRow++;
        }
    }

    private ByteArrayInputStream writeWorkbookToByteArray(Workbook workbook) throws IOException {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }
}