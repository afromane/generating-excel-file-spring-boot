package com.afromane.exportToExcel.service;

import com.afromane.exportToExcel.model.PeopleReviewData;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface MultiSheetPeopleReviewService {
    ByteArrayInputStream generateMultiSheetReport(Map<String, List<PeopleReviewData>> sheetData) throws IOException;
}
