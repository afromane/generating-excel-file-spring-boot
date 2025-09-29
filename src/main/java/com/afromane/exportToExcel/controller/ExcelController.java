package com.afromane.exportToExcel.controller;

import com.afromane.exportToExcel.model.PeopleReviewData;
import com.afromane.exportToExcel.service.ExcelGeneratorService;
import com.afromane.exportToExcel.service.MultiSheetPeopleReviewService;
import org.springframework.web.bind.annotation.RestController;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@RestController
public class ExcelController {

    @Autowired
    private ExcelGeneratorService excelReportService;
    @Autowired
    private MultiSheetPeopleReviewService multiSheetService;


    @GetMapping("/download-people-review")
    public ResponseEntity<byte[]> downloadPeopleReviewReport() throws IOException {
        List<PeopleReviewData> dataList = getSampleData();

        ByteArrayInputStream inputStream = excelReportService.generatePeopleReviewReport(dataList);
        byte[] excelData = inputStream.readAllBytes();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "people_review.xlsx");

        return new ResponseEntity<>(excelData, headers, HttpStatus.OK);
    }


    private List<PeopleReviewData> getSampleData() {
        return Arrays.asList(
                new PeopleReviewData(
                        "Manager RH", "01/01/2020", "3 ans", "Élevé", "Excellente", "Permanent",
                        "Oui", "Non", "Oui", "Non", "Oui", "Suivi trimestriel", "Très bon potentiel"
                ),
                new PeopleReviewData(
                        "Développeur Senior", "15/06/2021", "2 ans", "Moyen", "Bonne", "CDD",
                        "Non", "Oui", "Non", "Oui", "Non", "Formation technique", "Besoin de développement"
                )
        );
    }

    @GetMapping("/download-multi-sheet-report")
    public ResponseEntity<byte[]> downloadMultiSheetReport() throws IOException {
        // Préparer les données par direction
        Map<String, List<PeopleReviewData>> sheetData = Map.of(
                "DAF", getSampleDataDAF(),
                "DTA", getSampleDataDTA(),
                "DRH", getSampleDataDRH()
        );

        ByteArrayInputStream inputStream = multiSheetService.generateMultiSheetReport(sheetData);
        byte[] excelData = inputStream.readAllBytes();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "people_review_multi_sheet.xlsx");

        return new ResponseEntity<>(excelData, headers, HttpStatus.OK);
    }

    private List<PeopleReviewData> getSampleDataDAF() {
        return Arrays.asList(
                new PeopleReviewData("Directeur Financier", "01/01/2019", "4 ans", "Élevé", "Excellente", "Permanent",
                        "Oui", "Oui", "Oui", "Non", "Oui", "Plan de succession", "Top performer"),
                new PeopleReviewData("Contrôleur de Gestion", "15/03/2020", "3 ans", "Moyen", "Bonne", "Permanent",
                        "Oui", "Non", "Non", "Oui", "Non", "Formation Excel avancé", "Bon potentiel")
        );
    }

    private List<PeopleReviewData> getSampleDataDTA() {
        return Arrays.asList(
                new PeopleReviewData("Directeur Technique", "10/05/2018", "5 ans", "Élevé", "Excellente", "Permanent",
                        "Oui", "Oui", "Oui", "Oui", "Oui", "Mentorat", "Leader technique"),
                new PeopleReviewData("Architecte Système", "20/07/2021", "2 ans", "Élevé", "Très bonne", "CDD",
                        "Non", "Oui", "Oui", "Non", "Oui", "Certification cloud", "À convertir en CDI")
        );
    }

    private List<PeopleReviewData> getSampleDataDRH() {
        return Arrays.asList(
                new PeopleReviewData("Manager RH", "01/01/2020", "3 ans", "Élevé", "Excellente", "Permanent",
                        "Oui", "Non", "Oui", "Non", "Oui", "Suivi trimestriel", "Très bon potentiel"),
                new PeopleReviewData("Développeur Senior", "15/06/2021", "2 ans", "Moyen", "Bonne", "CDD",
                        "Non", "Oui", "Non", "Oui", "Non", "Formation technique", "Besoin de développement")
        );
    }

}