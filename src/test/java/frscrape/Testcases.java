/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package frscrape;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

/**
 *
 * @author prabakar
 */
public class Testcases extends DriverSetup {

    XSSFWorkbook workbook = new XSSFWorkbook();

    @Test
    public void fetchAssignment() throws InterruptedException, FileNotFoundException, IOException {
        navigateToSearchPage();
        FRbiblioScrape scrapeAssignmentDetails = new FRbiblioScrape(sfDriver);
        scrapeAssignmentDetails.getFRdetails();
        

    }
}
            
                    
                   
                
            
        
    



