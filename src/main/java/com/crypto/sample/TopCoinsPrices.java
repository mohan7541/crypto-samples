package com.crypto.sample;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import kong.unirest.HttpResponse;
import kong.unirest.JsonNode;
import kong.unirest.Unirest;
import kong.unirest.json.JSONArray;
import kong.unirest.json.JSONObject;

public class TopCoinsPrices {

	public static void main(String[] args) {
        String apiUrl = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=100&page=1";
        
        try {
            HttpResponse<JsonNode> response = Unirest.get(apiUrl).asJson();
            JSONArray cryptos = response.getBody().getArray();
            writeToExcel(cryptos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void writeToExcel(JSONArray cryptos) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Top 100 Cryptos");

        // Creating header row
        Row headerRow = sheet.createRow(0);
        String[] columns = {"Name", "Symbol", "Current Price", "24h Low", "24h High", "All Time Low", "All Time High", "% increase"};
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }

        // Filling data
        int rowNum = 1;
        for (int i = 0; i < cryptos.length(); i++) {
            JSONObject crypto = cryptos.getJSONObject(i);
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(crypto.getString("name"));
            row.createCell(1).setCellValue(crypto.getString("symbol").toUpperCase());
            row.createCell(2).setCellValue(crypto.getDouble("current_price"));
            row.createCell(3).setCellValue(crypto.getDouble("low_24h"));
            row.createCell(4).setCellValue(crypto.getDouble("high_24h"));
            row.createCell(5).setCellValue(crypto.getDouble("atl"));
            row.createCell(6).setCellValue(crypto.getDouble("ath"));
            BigDecimal ath = BigDecimal.valueOf(crypto.getDouble("ath"));
            BigDecimal cp = BigDecimal.valueOf(crypto.getDouble("current_price"));
            
            BigDecimal per_inc = ((ath.subtract(cp)).divide(cp, 6, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)));
            row.createCell(7).setCellValue(per_inc.doubleValue());
        }

        // Auto-size columns
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("Top100_current_increase.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

        System.out.println("Excel file written successfully.");
    }

}
