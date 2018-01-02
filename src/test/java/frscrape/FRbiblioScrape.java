/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package frscrape;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;

/**
 *
 * @author prabakar
 */
class FRbiblioScrape extends PageObject {

    @FindBy(how = How.ID, using = "search-num")
    public WebElement searchkeyTextbox1;

    //driver.findElement(By.xpath("//input[@id='search-num']"));
    public FRbiblioScrape(WebDriver driver) {
        super(driver);
    }

    public void getFRdetails() throws FileNotFoundException, IOException, InterruptedException {
        FileInputStream file = new FileInputStream(new File(
                "D:\\FRscrape\\Bookfr.xlsx"));
        String excelFileName = "D:\\FRscrape\\froutput.xlsx";//name of excel file
        String sheetName = "fndetails";//name of sheetWrite
        XSSFWorkbook workbookWrite = new XSSFWorkbook();
        XSSFSheet sheetWrite = workbookWrite.createSheet(sheetName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell;
        XSSFRow rowWrite;
        Row row = null;

        int rowStart = Math.min(15, sheet.getFirstRowNum());
        int rowEnd = Math.max(1400, sheet.getLastRowNum());
        for (int rowNum = rowStart + 1; rowNum < rowEnd; rowNum++) {
            String lookupItem;
            rowWrite = sheetWrite.createRow(rowNum);
            row = sheet.getRow(rowNum);

            if (row == null) {
                continue;
            }
            int columnNumber = 0;
            cell = row.getCell(columnNumber, Row.RETURN_BLANK_AS_NULL);
            DataFormatter formatter = new DataFormatter();
            cell.setCellType(Cell.CELL_TYPE_STRING);
            lookupItem = formatter.formatCellValue(cell);
            System.out.println(lookupItem);
            enterApplicationSubmit(lookupItem);

            boolean resultSet = checkResultAvailableOrNot();
            System.out.println(resultSet);
            if (resultSet) {
                System.out.println("***** In progress" + "\n" + "Look up Item " + lookupItem);
                WebElement descTable = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[2]/div/table/tbody"));
                List<WebElement> descRow = descTable.findElements(By.tagName("tr"));
                String pubNo = "";
                String pubDate = "";
                String appNo = "";
                String filingDate = "";
                String priorityNo = "";
                String priorityDate = "";
                List<String> priorityNos = new ArrayList<>();
                List<String> priorityDates = new ArrayList<>();
                String Representative = "";
                String grantDate = "";
                String status = "";
                String title="";
                String epDate="";
                
                title = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/h1")).getText();

                for (int i = 1; i <= descRow.size(); i++) {
                    String label = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[2]/div/table/tbody/tr[" + i + "]/th")).getText();
                    if (label.equalsIgnoreCase("Publication No. and date")) {
                        String pubNoAndDate = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[2]/div/table/tbody/tr[" + i + "]/td/strong")).getText();

                        System.out.println(pubNoAndDate);
                        String[] parts = pubNoAndDate.split("(\\s+-\\s+|\\s+)");
                        pubNo = parts[0];
                        pubDate = parts[1];
                    }
                    if (label.equalsIgnoreCase("Application No. and date of filing")) {
                        String appNoAndDate = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[2]/div/table/tbody/tr[3]/td")).getText();
                        System.out.println(appNoAndDate);
                        String[] parts = appNoAndDate.split("(\\s+-\\s+|\\s+)");
                        appNo = parts[0];
                        filingDate = parts[1];
                    }

                    if (label.equalsIgnoreCase("Priority No. and date")) {
                        String priNoAndDate = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[2]/div/table/tbody/tr[4]/td")).getText();
                        System.out.println(priNoAndDate);
                        List<String> priInfo = null;
                        if (priNoAndDate.contains(";")) {
                            priInfo = Arrays.asList(priNoAndDate.split(";"));
                            System.out.println(priInfo.size());
                            for (int j = 0; j < priInfo.size(); j++) {
                                String dummay = priInfo.get(j).trim();
                                System.out.println("dummay value" + dummay);
                                String[] parts = dummay.split("(\\s+-\\s+|\\s+)");
                                System.out.println("first" + parts[0]);
                                System.out.println("second" + parts[1]);
                                priorityNos.add(parts[0]);
                                priorityDates.add(parts[1]);
                            }
                            epDate = priorityDates.get(priorityDates.size() - 1);
                            System.out.println("nos" + priorityNos);
                            System.out.println("nos" + priorityDates);
                        } else {
                            String[] parts = priNoAndDate.split("(\\s+-\\s+|\\s+)");
                            priorityNos.add(parts[0]);
                            priorityDates.add(parts[1]);
                            epDate = Arrays.toString(priorityDates.toArray()).replace("[", "").replace("]", "");
                        }

                    }
                    priorityNo = Arrays.toString(priorityNos.toArray()).replace("[", "").replace("]", "");
                    priorityDate = Arrays.toString(priorityDates.toArray()).replace("[", "").replace("]", "");
                }

               // String Representative = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[7]/p")).getText();
                WebElement ed = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div"));
                List<WebElement> df = ed.findElements(By.tagName("div"));
                System.out.println(df.size());
                if (df.size() > 8) {
                    for (int r = 4; r <= 8; r++) {
                        String heade = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[" + r + "]/h3")).getText();
                        System.out.println(heade);
                        if (heade.equalsIgnoreCase("Representative :")) {
                            Representative = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/div[" + r + "]/p")).getText();
                            break;
                        }
                    }
                }

                System.out.println(Representative);

                boolean resultSet1 = checkElement();
                System.out.println(resultSet1);
                if (resultSet1) {
                    WebElement grant = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/table/tbody"));
                    List<WebElement> rows = grant.findElements(By.tagName("tr"));
                    for (int v = 1; v <= rows.size(); v++) {
                        String abc = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/table/tbody/tr[" + v + "]/th")).getText();
                        if (abc.equalsIgnoreCase("Date of grant")) {
                            grantDate = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/table/tbody/tr[" + v + "]/td[1]")).getText();
                            break;
                        }
                    }
                }else{
                    System.out.println("grant detail missing");
                }

                status = driver.findElement(By.xpath("//*[@id='statutDoc']/div[1]/h2/strong")).getText();
                System.out.println(grantDate);
                rowWrite.createCell(0).setCellValue(lookupItem);
                rowWrite.createCell(1).setCellValue(pubNo); //app no
                rowWrite.createCell(2).setCellValue(pubDate); //file date
                rowWrite.createCell(3).setCellValue(appNo);
                rowWrite.createCell(4).setCellValue(filingDate);
                rowWrite.createCell(5).setCellValue(priorityNo);
                rowWrite.createCell(6).setCellValue(priorityDate);
                rowWrite.createCell(7).setCellValue(epDate);
                rowWrite.createCell(8).setCellValue(Representative);
                rowWrite.createCell(9).setCellValue(grantDate);
                rowWrite.createCell(10).setCellValue(status);
                rowWrite.createCell(11).setCellValue(title);
                

                driver.findElement(By.xpath("//a[contains(.,'Advanced Search')]")).click();
                Thread.sleep(3000);

            } else {
                System.out.println("Biblio detail not available");
                rowWrite.createCell(0).setCellValue(lookupItem);
                rowWrite.createCell(1).setCellValue("Details not available");
                driver.findElement(By.xpath("//a[contains(.,'Advanced Search')]")).click();
                Thread.sleep(3000);
            }

            //write this workbookWrite to an Outputstream.
            try (FileOutputStream fileOut = new FileOutputStream(excelFileName)) {
                rowWrite = sheetWrite.createRow(0);
                rowWrite.createCell(0).setCellValue("input");
                rowWrite.createCell(1).setCellValue("publication#");
                rowWrite.createCell(2).setCellValue("publication Date");
                rowWrite.createCell(3).setCellValue("application#");
                rowWrite.createCell(4).setCellValue("filing date");
                rowWrite.createCell(5).setCellValue("priority no");
                rowWrite.createCell(6).setCellValue("priority date");
                rowWrite.createCell(7).setCellValue("EP date");
                rowWrite.createCell(8).setCellValue("Representative");
                rowWrite.createCell(9).setCellValue("grantDate");
                rowWrite.createCell(10).setCellValue("status");
                rowWrite.createCell(11).setCellValue("title");

                //write this workbookWrite to an Outputstream.
                workbookWrite.write(fileOut);
                fileOut.flush();
            }
        }
    }

    public void enterApplicationSubmit(String searchValue) throws InterruptedException {
        this.searchkeyTextbox1.clear();
        this.searchkeyTextbox1.sendKeys(searchValue);
        this.searchkeyTextbox1.sendKeys(Keys.ENTER);
        Thread.sleep(3000);
    }

    public Boolean checkResultAvailableOrNot() {
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
        boolean checkresultGrid;
        try {
            WebElement tf = driver.findElement(By.xpath("//*[@id='c38']/div[3]"));
            tf.isDisplayed();
            checkresultGrid = true;
        } catch (NoSuchElementException e) {
            checkresultGrid = false;
        } finally {
            driver.manage().timeouts().implicitlyWait(90, TimeUnit.SECONDS);
        }
        return checkresultGrid;
    }

    public boolean checkElement() {
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
        boolean checkGrantTable;
        try {
            WebElement grantTable = driver.findElement(By.xpath("//*[@id='c38']/div[3]/div/table"));
            grantTable.isDisplayed();
            checkGrantTable = true;
        } catch (NoSuchElementException e) {
            checkGrantTable = false;
        } finally {
            driver.manage().timeouts().implicitlyWait(90, TimeUnit.SECONDS);
        }
        return checkGrantTable;
    }

}
