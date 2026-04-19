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
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

abstract public class AParser {
    // Abstract class for any parser
    String inFileName;
    String outFileName;
    OutputStream out;
    BufferedWriter writer;

    HSSFWorkbook book;
    HSSFSheet sheet;

    XSSFWorkbook nBook;
    XSSFSheet nSheet;

    abstract boolean check();
    abstract void parse();

    boolean open(String inFileName, XLSType type, String outFileName, String charset) {
        boolean result = false;
        this.inFileName = inFileName;
        try {
            if (type == XLSType.XLSX) { // New XLSX format (XML ZIP)
              nBook = new XSSFWorkbook(new FileInputStream(this.inFileName));
              nSheet = nBook.getSheetAt(0);
            }
            else if (type == XLSType.XLS) { // Old XLS format (binary horror)
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(this.inFileName));
                book = new HSSFWorkbook(fs);
                sheet = book.getSheetAt(0);
            }
        }
        catch (Exception e) {
            System.out.println("E001. Error opening input file: " + this.inFileName);
        }
        this.outFileName = outFileName;
        try {
            Charset chs = Charset.forName(charset);
            try {
                out = new FileOutputStream(outFileName);
                writer = new BufferedWriter(new OutputStreamWriter(out, chs));
                result = true;
            }
            catch (Exception e) {
                System.out.println("E002. Error opening output file: " + this.outFileName + " : " + e.getMessage());
            }
        }
        catch (Exception e) {
            System.out.println("E011. Unknown output file charset: " + charset);
        }
        return result;
    }

    void close() {
        try {
            if (book != null) {
                book.close();
            }
            if (nBook != null) {
                nBook.close();
            }
        }
        catch (Exception e) {
            System.out.println("E003. Error when closing input file: " + inFileName);
        }
        try {
            if (writer != null) {
                writer.flush();
                writer.close();
            }
        }
        catch (Exception e) {
            System.out.println("E004. Error when closing output file: " + outFileName);
        }
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
