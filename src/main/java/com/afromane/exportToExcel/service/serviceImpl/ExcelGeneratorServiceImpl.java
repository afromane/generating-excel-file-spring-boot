package com.afromane.exportToExcel.service.serviceImpl;

import com.afromane.exportToExcel.model.PeopleReviewData;
import com.afromane.exportToExcel.service.ExcelGeneratorService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import java.io.InputStream;

@Service
public class ExcelGeneratorServiceImpl implements ExcelGeneratorService {

    private static final String TEMPLATE_PATH = "templates/single.xlsx";


@Override
public ByteArrayInputStream generatePeopleReviewReport(List<PeopleReviewData> dataList) throws IOException {
    // Charger votre fichier template existant
    InputStream templateStream = new ClassPathResource("templates/single.xlsx").getInputStream();
    Workbook workbook = new XSSFWorkbook(templateStream);
    Sheet sheet = workbook.getSheetAt(0);

    Row dateRow = sheet.getRow(0);
    if (dateRow != null) {
        Cell dateCell = dateRow.getCell(14);
        if (dateCell != null) {
            dateCell.setCellValue(java.time.LocalDate.now()
                    .format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        }
    }
    Row titleRow = sheet.getRow(1); // Ligne 2 affichée = index 1
    if (titleRow != null) {
        Cell titleCell = titleRow.getCell(1); // prends la 1ère cellule de la zone fusionnée
        if (titleCell != null) {
            String oldValue = titleCell.getStringCellValue();
            String direction = " DIRECTION RESSOURCES HUMAINES"; // ton texte à ajouter
            titleCell.setCellValue(oldValue + direction);
        }
    }


    // Ligne de départ pour les données (dans ton fichier → ligne 6, donc index 5)
    int rowNum = 6;

    for (PeopleReviewData data : dataList) {
        Row dataRow = sheet.getRow(rowNum);
        if (dataRow == null) {
            dataRow = sheet.createRow(rowNum);
        }

        // Insérer les données dans les colonnes déjà prévues
        dataRow.createCell(1).setCellValue(data.getPostesOccupants());
        dataRow.createCell(2).setCellValue(data.getDatePriseFonction());
        dataRow.createCell(3).setCellValue(data.getAnciennete());
        dataRow.createCell(4).setCellValue(data.getPotentielEvolution());
        dataRow.createCell(5).setCellValue(data.getPerformance());
        dataRow.createCell(6).setCellValue(data.getStatut());
        dataRow.createCell(7).setCellValue(data.getRevalorisation());
        dataRow.createCell(8).setCellValue(data.getPromotion());
        dataRow.createCell(9).setCellValue(data.getPrimeExceptionnelle());
        dataRow.createCell(10).setCellValue(data.getAutresAvantages());
        dataRow.createCell(11).setCellValue(data.getFormations());
        dataRow.createCell(12).setCellValue(data.getActionSuivi2024());
        dataRow.createCell(13).setCellValue(data.getCommentaires());

        rowNum++;
    }

    // Écriture en mémoire
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    workbook.write(out);
    workbook.close();
    templateStream.close();

    return new ByteArrayInputStream(out.toByteArray());
}


}
