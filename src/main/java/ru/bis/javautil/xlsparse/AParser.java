package ru.bis.javautil.xlsparse;

import org.apache.poi.hssf.usermodel.HSSFCell;
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
    BufferedWriter writer;

    HSSFWorkbook book;
    HSSFSheet sheet;

    XSSFWorkbook nBook;
    XSSFSheet nSheet;

    boolean check() { // Check statement for expected format
        return true;
    }
    int parse() { // Parse all statement and create CSV
        return 0;
    }

    boolean process(String inFileName, XLSType type, String outFileName, String charset) {
        boolean result = false;
        Charset chs;

        String tmpInFileName = inFileName + ".process";
        String tmpOutFileName = outFileName + ".process";

        File input = new File(inFileName);
        File tmpIn = new File(tmpInFileName);
        if (!input.exists()) {
            System.out.println("E020. Input file is not found: " + inFileName);
            return false;
        }
        if (!input.renameTo(tmpIn)) {
            System.out.println("E021. Error renaming input file " + inFileName + " to " + tmpInFileName);
            return false;
        }

        try (FileInputStream in = new FileInputStream(tmpInFileName)) { // Try for input
            try {
                if (type == XLSType.XLSX) { // New XLSX format (XML ZIP)
                    nBook = new XSSFWorkbook(in);
                    nSheet = nBook.getSheetAt(0);
                } else if (type == XLSType.XLS) { // Old XLS format (binary horror)
                    POIFSFileSystem fs = new POIFSFileSystem(in);
                    book = new HSSFWorkbook(fs);
                    sheet = book.getSheetAt(0);
                }
                try {
                    chs = Charset.forName(charset);
                }
                catch (Exception e) {
                    System.out.println("E022. Unknown output file charset: " + charset);
                    chs = StandardCharsets.UTF_8; // Error, use charset by default
                }
                try (FileOutputStream out = new FileOutputStream(tmpOutFileName);
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
                    System.out.println("E023. Error opening output file: " + outFileName + " : " + e.getMessage());
                }
                finally {
                    File tmpOut = new File(tmpOutFileName);
                    if (tmpOut.exists()) {
                        File out = new File(outFileName);
                        if (out.exists()) { // Remove old output file
                            if (!out.delete()) {
                                System.out.println("E024. Can't delete old output file: " + outFileName);
                                result = false;
                            }
                            if (!tmpOut.renameTo(out)) {
                                System.out.println("E025. Error renaming input file " + tmpOutFileName + " to " + outFileName);
                                result = false;
                            }
                        }
                    }
                    try {
                        if (book != null) {
                            book.close();
                        }
                        if (nBook != null) {
                            nBook.close();
                        }
                    }
                    catch (Exception e) {
                        System.out.println("E026. Error when closing XLS file: " + outFileName);
                        result = false;
                    }
                }
            }
            catch (Exception e) {
                System.out.println("E027. Error opening XLS file: " + inFileName);
            }
        }
        catch (IOException e) {
            System.out.println("E028. Error opening XLS file: " + inFileName);
        }
        finally {
            if (tmpIn.exists()) {
                if (!tmpIn.delete()) {
                    System.out.println("E029. Error deleting temporary file: " + tmpInFileName);
                }
            }
        }
        return result;
    }

    static String getStrNumber(HSSFCell cell) { // Returns string with decimal value 0.00 from String or Numeric cell
        String str;
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
        String str;
        try {
            str = cell.getStringCellValue();
        }
        catch (Exception e) { // May be date cell
            Date date = cell.getDateCellValue();
            LocalDate localDate = LocalDate.ofInstant(date.toInstant(), ZoneId.systemDefault());
            str = localDate.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        }
        return str;
    }
}
