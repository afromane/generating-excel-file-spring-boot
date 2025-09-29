package com.afromane.exportToExcel.controller;

import com.afromane.exportToExcel.model.PeopleReviewData;
import com.afromane.exportToExcel.service.ExcelGeneratorService;
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

@RestController
public class ExcelController {

    @Autowired
    private ExcelGeneratorService excelReportService;


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

}