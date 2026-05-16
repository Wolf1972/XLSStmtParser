package ru.bis.javautil.xlsparse;

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
    boolean parse(ErrHandleStrategy errHandleStrategy) { // Parse all statement and create CSV
        return true;
    }

    /**
     * Main method to parse statement
     * @param inFileName - input file name
     * @param type - XLS or XLSX
     * @param outFileName - output file name
     * @param charset - charset
     * @param errorHandling - which errors will fail the operation: ALL - all errors, FORMAT - only checked format errors, NONE - no errors checking
     * @param arcFileName - archive file name (when empty - processed statement file just will be deleted)
     * @return true when the operation is successful, false when the operation fails.
     */
    boolean process(String inFileName, XLSType type, String outFileName, String charset, ErrHandleStrategy errorHandling, String arcFileName) {
        boolean result = false;
        Charset chs;

        String tmpInFileName = inFileName + ".process";
        String tmpOutFileName = outFileName + ".process";

        File input = new File(inFileName);
        File tmpIn = new File(tmpInFileName);
        File tmpOut = new File(tmpOutFileName);
        File out = new File(outFileName);

        if (!input.exists()) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E020","E020. Input file is not found: {0}", inFileName);
            return false;
        }
        if (!input.renameTo(tmpIn)) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E021", "E021. Error renaming input file {0} to {1}", inFileName, tmpInFileName);
            return false;
        }

        if (type == XLSType.AUTO) { // Define Excel file format type (XLS or XLSX) - 2 first chars in file
            try (FileInputStream in = new FileInputStream(tmpInFileName)) {
                char ch0 = (char) in.read();
                char ch1 = (char) in.read();
                if (ch0 == 'P' && ch1 == 'K') { // Zip signature found - new Excel format type (XLSX)
                    type = XLSType.XLSX;
                }
                else {
                    type = XLSType.XLS;
                }
            }
            catch (IOException e) {
                Main.logr.log(System.Logger.Level.ERROR,"error.E022", "E022. Can't define Excel file type");
            }
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
                    Main.logr.log(System.Logger.Level.ERROR,"error.E023", "E023. Unknown output file charset: {0}", charset);
                    chs = StandardCharsets.UTF_8; // Error, use charset by default
                }
                try (FileOutputStream fileOut = new FileOutputStream(tmpOutFileName);
                     OutputStreamWriter outWr = new OutputStreamWriter(fileOut, chs);
                     BufferedWriter writer = new BufferedWriter(outWr)) { // Try for output
                    this.writer = writer;
                    if (check() || errorHandling == ErrHandleStrategy.NONE) {
                        result = parse(errorHandling);
                    }
                }
                catch (Exception e) {
                    Main.logr.log(System.Logger.Level.ERROR,"error.E024","E024. Error opening output file: {0}: {1}", outFileName, e.getMessage());
                }
                finally {

                    if (tmpOut.exists()) {
                        if (out.exists()) { // Remove old output file
                            if (!out.delete()) {
                                Main.logr.log(System.Logger.Level.ERROR,"error.E025", "E025. Can't delete previous output file: {0}", outFileName);
                                result = false;
                            }
                        }
                        if (!tmpOut.renameTo(out)) {
                            Main.logr.log(System.Logger.Level.ERROR,"error.E026", "E026. Error renaming output file {0} to {1}", tmpOutFileName, outFileName);
                            result = false;
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
                        Main.logr.log(System.Logger.Level.ERROR,"error.E027", "E027. Error when closing XLS file: {0}", outFileName);
                        result = false;
                    }
                }
            }
            catch (Exception e) {
                Main.logr.log(System.Logger.Level.ERROR,"error.E028", "E028. Error opening Excel file: {0}", inFileName);
            }
        }
        catch (IOException e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E029", "E029. Error opening input file: {0}", inFileName);
        }
        finally {
            if (tmpIn.exists()) {
                if (result) {
                    if (!arcFileName.isEmpty()) { // Is archive need?
                        File arc = new File(arcFileName);
                        if (arc.exists()) {
                            Main.logr.log(System.Logger.Level.ERROR,"error.E030", "E030. Archive file already exists: {0}", arcFileName);
                        }
                        else {
                            if (!tmpIn.renameTo(arc)) {
                                Main.logr.log(System.Logger.Level.ERROR,"error.E031", "E031. Can't create archive file: {0}", arcFileName);
                                // We don't need to delete the input file because we couldn't create archive
                            }
                        }
                    }
                    else if (!tmpIn.delete()) { // Delete statement file
                        Main.logr.log(System.Logger.Level.ERROR,"error.E032", "E032. Error deleting temporary file: {0}", tmpInFileName);
                    }
                }
                else { // When error: return the input file with its previous name
                    if (!tmpIn.renameTo(input)) {
                        Main.logr.log(System.Logger.Level.ERROR,"error.E033", "E033. Error rollback input file {0} to {1}", tmpInFileName, inFileName);
                    }
                    if (out.exists()) {
                        if (!out.delete()) {
                            Main.logr.log(System.Logger.Level.ERROR,"error.E034", "E034. Can't delete output file with error: {0}", outFileName);
                        }
                    }
                }
            }
        }
        return result;
    }


    String getCellString(int rowNo, int cellNo) { // Returns string from specified row and cell from any book - XLS or XLSX
        String result = "";
        try {
            if (nSheet != null) {
                result = nSheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            } else if (sheet != null) {
                result = sheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            }
        }
        catch (Exception e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E035", "E035. Can't get value for cell {0}:{1}", rowNo + 1, cellNo + 1);
        }
        return result;
    }

    String getCellNumber(int rowNo, int cellNo) { // Returns string with decimal value like "0.00" from String or Numeric cell
        String str = "";
        try {
            if (nSheet != null) {
                str = nSheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            }
            else if (sheet != null) {
                str = sheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            }
        }
        catch (Exception e) { // May be numeric cell
            double dec = 0;
            try {
                if (nSheet != null) {
                    dec = nSheet.getRow(rowNo).getCell(cellNo).getNumericCellValue();
                } else if (sheet != null) {
                    dec = sheet.getRow(rowNo).getCell(cellNo).getNumericCellValue();
                }
                str = new DecimalFormat("#0.00").format(dec);
            }
            catch (Exception x) {
                Main.logr.log(System.Logger.Level.ERROR,"error.E036", "E036. Can't get decimal value for cell {0}:{1}", rowNo + 1, cellNo + 1);
            }
        }
        return str;
    }

    String getCellDate(int rowNo, int cellNo) { // Returns string with date value "DD.MM.YYYY" from String or Date cell
        String str = "";
        try {
            if (nSheet != null) {
                str = nSheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            }
            else if (sheet != null) {
                str = sheet.getRow(rowNo).getCell(cellNo).getStringCellValue();
            }
        }
        catch (Exception e) { // May be date cell
            Date date = new Date();
            try {
                if (nSheet != null) {
                    date = nSheet.getRow(rowNo).getCell(cellNo).getDateCellValue();
                } else if (sheet != null) {
                    date = sheet.getRow(rowNo).getCell(cellNo).getDateCellValue();
                }
                LocalDate localDate = LocalDate.ofInstant(date.toInstant(), ZoneId.systemDefault());
                str = localDate.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
            }
            catch (Exception x) {
                Main.logr.log(System.Logger.Level.ERROR,"error.E037", "E037. Can't get date value for cell {0}:{1}", rowNo + 1, cellNo + 1);
            }
        }
        return str;
    }
}
