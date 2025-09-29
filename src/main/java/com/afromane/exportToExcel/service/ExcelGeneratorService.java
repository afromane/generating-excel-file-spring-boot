package com.afromane.exportToExcel.service;

import com.afromane.exportToExcel.model.PeopleReviewData;
import org.apache.poi.ss.formula.functions.T;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

public interface ExcelGeneratorService {
    ByteArrayInputStream generatePeopleReviewReport(List<PeopleReviewData> dataList)  throws IOException;

}
