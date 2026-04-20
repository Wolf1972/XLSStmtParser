package ru.bis.javautil.xlsparse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

abstract public class AParser {
    // Abstract class for any parser
    String inFileName;
    String outFileName;
    BufferedWriter writer;

    HSSFWorkbook book;
    HSSFSheet sheet;

    XSSFWorkbook nBook;
    XSSFSheet nSheet;

    boolean check() { // Check statement for expected format
        return true;
    };
    int parse() { // Parse all statement and create CSV
        return 0;
    };

    boolean process(String inFileName, XLSType type, String outFileName, String charset) {
        boolean result = false;
        Charset chs;

        this.inFileName = inFileName;
        try (FileInputStream in = new FileInputStream(this.inFileName)) { // Try for input
            try {
                if (type == XLSType.XLSX) { // New XLSX format (XML ZIP)
                    nBook = new XSSFWorkbook(in);
                    nSheet = nBook.getSheetAt(0);
                } else if (type == XLSType.XLS) { // Old XLS format (binary horror)
                    POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(this.inFileName));
                    book = new HSSFWorkbook(fs);
                    sheet = book.getSheetAt(0);
                }
                this.outFileName = outFileName;
                try {
                    chs = Charset.forName(charset);
                }
                catch (Exception e) {
                    System.out.println("E021. Unknown output file charset: " + charset);
                    chs = StandardCharsets.UTF_8; // Error, use charset by default
                }
                try (FileOutputStream out = new FileOutputStream(outFileName);
                     OutputStreamWriter outwr = new OutputStreamWriter(out, chs);
                     BufferedWriter writer = new BufferedWriter(outwr)) { // Try for output
                    this.writer = writer;
                    if (check()) {
                        int lines = parse();
                        System.out.println("Done. " + lines + " operation(s) created.");
                    }
                    result = true;
                }
                catch (Exception e) {
                    System.out.println("E023. Error opening output file: " + this.outFileName + " : " + e.getMessage());
                }
                finally {
                    try {
                        if (book != null) {
                            book.close();
                        }
                        if (nBook != null) {
                            nBook.close();
                        }
                    }
                    catch (Exception e) {
                        System.out.println("E025. Error when closing XLS file: " + outFileName);
                    }
                }
            }
            catch (Exception e) {
                System.out.println("E022. Error opening XLS file: " + this.inFileName);
            }
        }
        catch (IOException e) {
            System.out.println("E021. Error opening XLS file: " + this.inFileName);
        }
        return result;
    }

    static String getStrNumber(HSSFCell cell) { // Returns string with decimal value 0.00 from String or Numeric cell
        String str = "";
        try {
            str = cell.getStringCellValue();
        }
        catch (Exception e) { // May be numeric cell
            double dec = cell.getNumericCellValue();
            str = new DecimalFormat("#0.00").format(dec);
        }
        return str;
    }

    static String getStrDate(HSSFCell cell) { // Returns string with date value "DD.MM.YYYY" from String or Date cell
        String str = "";
        try {
            str = cell.getStringCellValue();
        }
        catch (Exception e) { // May be date cell
            Date date = cell.getDateCellValue();
            LocalDate localDate = LocalDate.ofInstant(date.toInstant(), ZoneId.systemDefault());;
            str = localDate.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        }
        return str;
    }
}
