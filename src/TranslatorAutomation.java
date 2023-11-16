import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class TranslatorAutomation {

    public static void main(String[] args) {
        // Set the path of the ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", "chromedriver-mac-arm64/chromedriver");

        // Create a new instance of the ChromeDriver
        WebDriver driver = new ChromeDriver();

        // region Read Quotes ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        // Navigate to the quotes page and read the content
        driver.get("https://blog.empuls.io/famous-scientists-quotes/");
        try {
            WebDriverWait wait = new WebDriverWait(driver, 10);

            // Wait until the page loads
            String orderedList = "//*[@id=\"main-content\"]/div/section[2]/div/div[1]/article/section[1]/ol";
            waitForElementVisible(wait, driver, orderedList);

            // Read the text of the translation result
            String listText = readElementText(wait, driver, orderedList);

            // Call the function to write to Excel
            writeStringToExcel(listText, "output.xls");

        } catch (Exception e) {
            e.printStackTrace();
        }
        // endregion |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        // region Google Translator ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        // Navigate to Google Translator and translate the content
        driver.get("https://translate.google.com/?hl=en&tab=TT");

        try {
            WebDriverWait wait = new WebDriverWait(driver, 20);

            // Click on the first element
            String emptyTranslateBox = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[3]/c-wiz[1]";
            waitForElementVisible(wait, driver, emptyTranslateBox);
            clickElement(wait, driver, emptyTranslateBox);

            // Click on language dropdown
            String languageDropdown = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[1]/c-wiz/div[1]/c-wiz/div[5]/button/div[3]";
            waitForElementVisible(wait, driver, languageDropdown);
            clickElement(wait, driver, languageDropdown);

            // Click on the Bengali option
            String bengaliLanguageOption = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[1]/c-wiz/div[2]/c-wiz/div[2]/div/div[3]/div/div[2]/div[13]/div[2]";
            String frenchLanguageOption = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[1]/c-wiz/div[2]/c-wiz/div[2]/div/div[3]/div/div[2]/div[35]/div[2]";
            waitForElementVisible(wait, driver, frenchLanguageOption);
            clickElement(wait, driver, frenchLanguageOption);

            // click close language dropdown button
            String closeLanguageDropdownButton = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[1]/c-wiz/div[1]/c-wiz/div[5]/button/div[3]";
            waitForElementVisible(wait, driver, closeLanguageDropdownButton);
            clickElement(wait, driver, closeLanguageDropdownButton);
            clickElement(wait, driver, closeLanguageDropdownButton);

            // Reading "output.xls" file
            List<String> lines = readExcelFile("output.xls");

            List<String> translatedTextList = new ArrayList<>();

            // Loop through the list of lines
            for (String line : lines) {
                System.out.println(line);

                // Type into the input field
                String textAreaMain = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[3]/c-wiz[1]/span/span/div/textarea";
                typeIntoInputField(wait, driver, textAreaMain, line);

                // Wait for a specific duration (e.g., 2 seconds)
                try {
                    Thread.sleep(3000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                // Wait until the translation is done
                String starButton = "//*[@id=\"yDmH0d\"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[3]/c-wiz[2]/div[1]/div[7]/div/div[1]";
                waitForElementVisible(wait, driver, starButton);

                // Read the text of the translation result
                String translatedText = readElementText(wait, driver, starButton);
                System.out.println("Translated Text: " + translatedText);
                translatedTextList.add(translatedText); // Add translatedText to the list
            }
            writeListToExcel(translatedTextList, "output.xls", 1);

        } catch (Exception e) {
            e.printStackTrace();
        }
        // endregion |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        // region Reverso Translator ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        // Navigate to Reverso Translator and translate the content
        driver.get("https://www.reverso.net/text-translation");

        try {
            WebDriverWait wait = new WebDriverWait(driver, 20);

            // Click on the first element
            String emptyTranslateBox = "/html/body/app-root/app-translation/div/app-translation-box/div[1]/div[1]/div[2]/div[1]/div[1]/div/div[1]/textarea";
            waitForElementVisible(wait, driver, emptyTranslateBox);
            clickElement(wait, driver, emptyTranslateBox);

            // Click on language dropdown
            String languageDropdown = "/html/body/app-root/app-translation/div/app-translation-box/div[1]/div[1]/div[1]/app-language-switch/div/app-language-select[2]/div/div/app-icon/span";
            waitForElementVisible(wait, driver, languageDropdown);
            clickElement(wait, driver, languageDropdown);

            // Click on the French option
            String frenchLanguageOption = "/html/body/div[3]/app-language-select-options/div/ul/li[6]/span";
            waitForElementVisible(wait, driver, frenchLanguageOption);
            clickElement(wait, driver, frenchLanguageOption);

            // click close language dropdown button
            String closeLanguageDropdownButton = "/html/body/app-root/app-translation/div/app-translation-box/div[1]/div[1]/div[1]/app-language-switch/div/app-language-select[2]/div/div/app-icon/span";
            waitForElementVisible(wait, driver, closeLanguageDropdownButton);
            clickElement(wait, driver, closeLanguageDropdownButton);
            clickElement(wait, driver, closeLanguageDropdownButton);

            // Reading "output.xls" file
            List<String> lines = readExcelFile("output.xls");

            List<String> translatedTextList = new ArrayList<>();

            // Loop through the list of lines
            for (String line : lines) {
                System.out.println(line);

                // Type into the input field
                String textAreaMain = "/html/body/app-root/app-translation/div/app-translation-box/div[1]/div[1]/div[2]/div[1]/div[1]/div/div[1]/textarea";
                typeIntoInputField(wait, driver, textAreaMain, line);

                // Wait for a specific duration (e.g., 2 seconds)
                try {
                    Thread.sleep(4000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                // Wait until the translation is done
                String starButton = "/html/body/app-root/app-translation/div/app-translation-box/div[1]/div[1]/div[2]/div[2]";
                waitForElementVisible(wait, driver, starButton);

                // Define an array of keywords to remove
                String[] keywordsToRemove = {"Rephrase", "NEW"};

                // Build a regular expression to match any of the keywords
                String regex = String.join("|", keywordsToRemove);

                // Read the text of the translation result
                String translatedText = readElementText(wait, driver, starButton);
                translatedText = translatedText.replaceAll(regex, "");
                System.out.println("Translated Text: " + translatedText);
                translatedTextList.add(translatedText); // Add translatedText to the list
            }
            System.out.println(translatedTextList);
            writeListToExcel(translatedTextList, "output.xls", 2);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
        // endregion |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    }

    private static void clickElement(WebDriverWait wait, WebDriver driver, String xpath) {
        // Click on the element after it becomes clickable
        WebElement element = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));
        element.click();
    }

    private static void typeIntoInputField(WebDriverWait wait, WebDriver driver, String xpath, String text) {
        // Type the specified text into the input field
        WebElement inputField = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
        inputField.click();
        inputField.clear();
        inputField.sendKeys(text);
    }

    private static void waitForElementVisible(WebDriverWait wait, WebDriver driver, String xpath) {
        // Wait until the specified element becomes visible
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
    }

    private static String readElementText(WebDriverWait wait, WebDriver driver, String xpath) {
        // Read and return the text of the specified element
        WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
        return element.getText();
    }

    private static void writeStringToExcel(String multilineString, String fileName) {
        // Write a multiline string to an Excel file
        try (Workbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Add headers
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("Main Sentence");

            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("Google Translator");

            Cell headerCell3 = headerRow.createCell(2);
            headerCell3.setCellValue("Reverso");

            // Split the multiline string into lines
            String[] lines = multilineString.split("\n");

            // Create a new row for each line
            for (int i = 0; i < 5; i++) {
                Row row = sheet.createRow(i + 1); // Start from the second row after headers
                Cell cell = row.createCell(0);
                cell.setCellValue(lines[i]);
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> readExcelFile(String fileName) {
        List<String> lines = new ArrayList<>();

        try (Workbook workbook = new HSSFWorkbook(new FileInputStream(fileName))) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Iterate over rows
            Iterator<Row> rowIterator = sheet.iterator();
            // Skip the first row with headers
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0); // Assuming the data is in the first column

                if (cell != null) {
                    lines.add(cell.getStringCellValue());
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return lines;
    }

    private static void writeListToExcel(List<String> translatedTextList, String fileName, int columnNum) {
        try (FileInputStream fileIn = new FileInputStream(fileName);
             Workbook workbook = new HSSFWorkbook(fileIn)) {

            // Assuming the data is in the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Start from the second row after headers
            int rowIndex = 1;

            // Iterate through the translated text list
            for (String translatedText : translatedTextList) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }

                // Create or update the cell in the "Google Translator" column (index 1)
                Cell cell = row.createCell(columnNum);
                cell.setCellValue(translatedText);

                rowIndex++;
            }

            // Save the updated workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
