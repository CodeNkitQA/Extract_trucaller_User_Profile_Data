import com.google.common.collect.ImmutableList;
import io.appium.java_client.AppiumBy;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.nativekey.AndroidKey;
import io.appium.java_client.android.nativekey.KeyEvent;
import io.appium.java_client.android.options.UiAutomator2Options;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.PointerInput;
import org.openqa.selenium.interactions.Sequence;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class Data_Extract {

    static AndroidDriver driver;


    @BeforeClass
    public void setUp() throws MalformedURLException {
        UiAutomator2Options options = new UiAutomator2Options()
                .setPlatformName("Android")
//                .setDeviceName("emulator-5554") // Ensure this matches the emulator/device name
                .setDeviceName("Pixel7a")
                .setAutomationName("UiAutomator2")
                .setApp(System.getProperty("user.dir") + "/apps/truecaller-caller-id-and-block-13-63-7.apk")
//                .setApp("/Users/ankitkumar/IdeaProjects/Quaha-main/apps/app-release (35).apk")
//                .setAutoGrantPermissions(true)
//                .setAppPackage("com.truecaller")
//                .setAppActivity("com.truecaller.ui.LauncherActivity") // Updated activity name
                .setNoReset(true)
                .setFullReset(false);


        driver = new AndroidDriver(new URL("http://127.0.0.1:4723"), options);

        try {
            Thread.sleep(4000); // Wait for the app to load
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    @AfterClass
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }

    protected void dismissKeyboard() {
        if (driver != null) {
            driver.pressKey(new KeyEvent(AndroidKey.BACK));
        }
    }

    public enum ScrollDirection {
        UP, DOWN, LEFT, RIGHT
    }

    public static void scroll(ScrollDirection dir, double scrollRatio) {
        Duration SCROLL_DUR = Duration.ofMillis(300);
        if (scrollRatio < 0 || scrollRatio > 1) {
            throw new IllegalArgumentException("Scroll distance must be between 0 and 1");
        }
        Dimension size = driver.manage().window().getSize();
        Point midPoint = new Point((int) (size.width * 0.5), (int) (size.height * 0.5));
        int bottom = midPoint.y + (int) (midPoint.y * scrollRatio);
        int top = midPoint.y - (int) (midPoint.y * scrollRatio);
        int left = midPoint.x - (int) (midPoint.x * scrollRatio);
        int right = midPoint.x + (int) (midPoint.x * scrollRatio);

        if (dir == ScrollDirection.UP) {
            swipe(new Point(midPoint.x, top), new Point(midPoint.x, bottom), SCROLL_DUR);
        } else if (dir == ScrollDirection.DOWN) {
            swipe(new Point(midPoint.x, bottom), new Point(midPoint.x, top), SCROLL_DUR);
        } else if (dir == ScrollDirection.LEFT) {
            swipe(new Point(left, midPoint.y), new Point(right, midPoint.y), SCROLL_DUR);
        } else {
            swipe(new Point(right, midPoint.y), new Point(left, midPoint.y), SCROLL_DUR);
        }
    }

    protected static void swipe(Point start, Point end, Duration duration) {
        PointerInput input = new PointerInput(PointerInput.Kind.TOUCH, "finger1");
        Sequence swipe = new Sequence(input, 0);
        swipe.addAction(input.createPointerMove(Duration.ZERO, PointerInput.Origin.viewport(), start.x, start.y));
        swipe.addAction(input.createPointerDown(PointerInput.MouseButton.LEFT.asArg()));
        swipe.addAction(input.createPointerMove(duration, PointerInput.Origin.viewport(), end.x, end.y));
        swipe.addAction(input.createPointerUp(PointerInput.MouseButton.LEFT.asArg()));
        driver.perform(ImmutableList.of(swipe));
    }

    @Test

    public void log() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {
            // Fetch all numbers from Excel (First sheet)
            List<String> numbers = getAllDataFromExcel("/Users/ankitkumar/Downloads/Input_Sheet.xlsx", 0);
            var el1 = driver.findElement(AppiumBy.id("com.truecaller:id/searchBarLabel"));
            el1.click();
            Thread.sleep(2000);
            for (String number : numbers) {
                // Wait for and click the search bar element


                // Wait for the search input and send the number
                var el2 = driver.findElement(AppiumBy.xpath("//android.widget.AutoCompleteTextView[@resource-id=\"com.truecaller:id/search_field\"]"));
                el2.clear(); // Clear any previous number
                el2.sendKeys(number);
                Thread.sleep(2000);

                // Select the third instance of a view group (search result)
                var el3 = driver.findElement(AppiumBy.androidUIAutomator("new UiSelector().className(\"android.view.ViewGroup\").instance(3)"));
                el3.click();
                Thread.sleep(4000);

                // Get name, number, and other details
                String name = driver.findElement(AppiumBy.xpath("//android.widget.TextView[@resource-id=\"com.truecaller:id/nameOrNumber\"]")).getText();
                String phoneNumber = driver.findElement(AppiumBy.xpath("//android.widget.TextView[@resource-id=\"com.truecaller:id/number\"]")).getText();
//                String lastseen = String.valueOf(driver.findElement(AppiumBy.xpath("//android.widget.TextView[@resource-id=\"com.truecaller:id/title\" and @text=\"Local time 10:18 AM\"]")));

                // Save the fetched data into an Excel file
                writeDataToExcel("/Users/ankitkumar/Downloads/Outputsheet.xlsx", name, phoneNumber);

                // Navigate back to the previous screen
                driver.navigate().back();
                Thread.sleep(2000);
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }

        }
    }

    public List<String> getAllDataFromExcel(String filePath, int sheetIndex) {
        List<String> numbers = new ArrayList<>();
        FileInputStream fis = null;
        try {
            File excelFile = new File(filePath);
            if (!excelFile.exists()) {
                System.out.println("Excel file does not exist at the specified path: " + filePath);
                return numbers;
            }

            fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            for (Row row : sheet) {
                Cell cell = row.getCell(0); // Assuming numbers are in the first column
                if (cell != null) {
                    DataFormatter formatter = new DataFormatter();
                    String cellValue = formatter.formatCellValue(cell);
                    if (!cellValue.isEmpty()) {
                        numbers.add(cellValue); // Add the number to the list
                    }
                }
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return numbers;
    }

    public void writeDataToExcel(String filePath, String name, String number) {
        FileOutputStream fos = null;
        Workbook workbook = null;
        try {
            File excelFile = new File(filePath);
            if (excelFile.exists()) {
                FileInputStream fis = new FileInputStream(excelFile);
                workbook = new XSSFWorkbook(fis);
            } else {
                workbook = new XSSFWorkbook();
            }

            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                sheet = workbook.createSheet("Output");
            }

            int rowCount = sheet.getLastRowNum();
            Row row = sheet.createRow(rowCount + 1);

            // Fill in the values into the new row
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(number);
//            row.createCell(2).setCellValue(lastseen);

            fos = new FileOutputStream(excelFile);
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fos != null) {
                    fos.close();
                }
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }






}



