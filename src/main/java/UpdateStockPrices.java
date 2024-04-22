import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.io.FileOutputStream;

public class UpdateStockPrices {

    public static WebDriver driver;

    public static void main(String[] args) {
        ArrayList<String> symbols = getStockSymbol();;
        ArrayList<Double> prices = new ArrayList<>();
        try {
            Thread.sleep(15000);
        } catch(Exception ex){}

        // Specify the path to the chromedriver executable
        System.setProperty("webdriver.chrome.driver", "chromedriver.exe");

        // Create a new instance of the Chrome driver
        driver = new ChromeDriver();

        // Navigate to a website
        driver.get("https://kite.zerodha.com");
        WebDriverWait wait = new WebDriverWait(driver, 1000); // 10 seconds timeout
        WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(text(), 'NIFTY 50')]")));


        for(String symbol : symbols) {
            double price = getPrice(symbol);

            // Get the text of the element
            prices.add(price);
        }

        updatePrices(prices);
        // Close the browser
        driver.quit();
    }

    public static ArrayList<String> getStockSymbol() {
        try {
            // Specify the path to the Excel file
            String filePath = "StockData.xlsx";
            ArrayList<String> symbols = new ArrayList<>();
            // Create a FileInputStream to read the Excel file
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));

            // Create a Workbook instance
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Get the desired cell (e.g., cell A1)
            Row row = sheet.getRow(1);
            Cell cell = row.getCell(0);
            // Get the cell value
            int i = 1;

            while (cell != null) {
                symbols.add(cell.getStringCellValue());
                row = sheet.getRow(++i);
                if(row == null)
                    break;
                cell = row.getCell(0);
            }

            workbook.close();
            fileInputStream.close();

            return symbols;
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        return null;
    }

    public static void updatePrices(ArrayList<Double> prices) {
        try {
            // Specify the path to the Excel file
            String filePath = "C:/Users/risha/OneDrive/Documents/StockData.xlsx";

            // Create a FileInputStream to read the Excel file
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));

            // Create a Workbook instance
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Get the desired cell (e.g., cell A1)
            int i = 0;

            for (Double price : prices) {
                Row row = sheet.getRow(++i);
                Cell cell = row.getCell(1);
                // Set the new value for the cell
                System.out.println("price: " + price);

                if (cell == null) {
                    // If the cell doesn't exist, create a new one
                    cell = row.createCell(1);
                }

                cell.setCellValue(price);
            }


            // Create a FileOutputStream to write changes to the Excel file
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);

            // Write changes to the Excel file
            workbook.write(fileOutputStream);

            // Close the workbook, fileInputStream, and fileOutputStream
            workbook.close();
            fileInputStream.close();
            fileOutputStream.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static double getPrice(String symbol) {
        WebElement searchElement = driver.findElement(By.xpath("//input[@placeholder='Search (infy bse, nifty fut, etc)']"));
        synchronized (searchElement){
            try {
                searchElement.wait(3000);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        }

        searchElement.click();
        searchElement.sendKeys(symbol);

        WebElement stockElemInSearch = driver.findElement(By.xpath("(//span[text()='" + symbol + "'])[1]"));
        Actions actions = new Actions(driver);
        actions.moveToElement(stockElemInSearch).build().perform();
        WebElement buyButton = driver.findElement(By.xpath("(//button[@class='button-blue'])[1]"));
        synchronized (buyButton) {
            try {
            buyButton.wait(1000);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        }
        buyButton.click();
        double price = 0;

        while(price == 0) {
            WebElement priceElement = driver.findElement(By.xpath("//label[contains(text(), 'BSE')]//span[@class='last-price']"));
            try {
                String unformatedPrice = priceElement.getText();
                unformatedPrice = unformatedPrice.substring(1);
                String[] unformatedPriceSplitted = unformatedPrice.split(",");
                for(String part : unformatedPriceSplitted)
                    price += Double.parseDouble(part);
            } catch (Exception ex) {
                price = 0;
            }
        }

        return price;
    }
}
