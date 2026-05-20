package ru.bis.javautil.xlsparse;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

import static java.nio.file.Files.isRegularFile;

public class Util {

    static String lSep = System.lineSeparator(); // line separator (must be initialized before all)
    static String dSep; // decimal separator
    static String fSep; // field separator
    static String fileSep; // file separator
    static String outDateFormat;

    static final String nbsp = "\u00A0";

    /**
     * System parameters set
     */
    static void init() {
        DecimalFormat format = (DecimalFormat) DecimalFormat.getInstance();
        DecimalFormatSymbols symbols = format.getDecimalFormatSymbols();
        dSep = String.valueOf(symbols.getDecimalSeparator());
        Main.logr.log(System.Logger.Level.TRACE, "trace.system_decimal_separator", "System decimal separator: {0}", dSep);

        String lineSepStr = "";
        for (int i = 0; i < lSep.length(); i++) {
            lineSepStr += " 0x" + String.format("%04x", (int) lSep.charAt(i));
        }
        Main.logr.log(System.Logger.Level.TRACE,"trace.system_line_separator", "System line separator: {0}",  lineSepStr);

        fSep = ";";

        fileSep = File.separator;
        Main.logr.log(System.Logger.Level.TRACE,"trace.system_path_separator", "System file path separator: {0}", fileSep);

        outDateFormat = "yyyy-MM-dd";
    }

    /**
     * Converts long value to string with 2 digital digits separated
     * @param amount - long value
     * @return string
     */
    static String long2str(long amount) {
        return String.format("%d" + Util.dSep + "%02d", amount / 100, amount % 100);
    }

    /**
     * // Converts string with 2 decimal digits separated to long, possible NBSP between digits (1 000 000.00)
     * @param str - unput string
     * @return long value
     * @throws NumberFormatException when parse fails
     */
    static long str2long(String str) throws NumberFormatException {
        if (str == null || str.isEmpty()) {
            return 0;
        }
        else {
            return Long.parseLong(str.replaceAll("[ .," + nbsp + "]", ""));
        }
    }

    /**
     * Clean from string values \n, \r, replaces nbsp to ordinary spaces
     * @param str - input string
     * @return converted string, not null
     */
    static String cleanStr(String str) { //
        if (str == null) {
            return "";
        } else {
            return str.replaceAll("\n", "").replaceAll("\r", "").replaceAll(nbsp, " ");
        }
    }

    /**
     * Converts string 2 CSV-clean string (double quota replaced by "")
     * @param str - input string
     * @return string with CSV formatting
     */
    static String str2CSV(String str) { //
       if (str == null) {
           return "";
       }
       else {
           return str.replace("\"", "\"\"");
       }
    }

    /**
     * Changes file extension (e.g. *.xls to *.csv)
     * @param fileName - file name with source extension
     * @param ext - target extension
     * @return file name with target extension
     */
    static String changeFileExtension(String fileName, String ext) {
        String result = null;
        if (fileName != null) {
            int pos = fileName.lastIndexOf(".");
            if (pos >= 0) {
                result = fileName.substring(0, pos) + "." + ext;
            }
            else { // There is no "." in file name
                result = fileName + "." + ext;
            }
        }
        return result;
    }

    /**
     * Restricts string with some length even if this length more than length of the real string (instead Apache Commons)
     * @param str - input string
     * @param max - max length
     * @return converted string
     */
    static String leftStr(String str, int max) {
        return str == null ? "" : str.length() > max ? str.substring(0, max) : str;
    }

    /**
     * Returns resource string or default string
     * @param code - resource name
     * @param defaultString - return string when resource not found
     * @param args - multiply arguments for substitute (MessageFormat syntax)
     * @return result string
     */
    static String resource(String code, String defaultString, Object... args) { //
        try {
            String result;
            if (Main.bundle != null) {
                result = Main.bundle.getString(code);
            }
            else {
                result = defaultString;
            }
            return MessageFormat.format(result, args);
        }
        catch (Exception e) {
            return MessageFormat.format(defaultString, args);
        }
    }

    /**
     * Returns list of files, in case inFile is a directory returns all files from it
     * @param fileName - directory or file name
     * @return list of files (when input parameter is file - list with only one file
     */
    static List<Path> getFileList(String fileName) {
        List<Path> aFiles = new ArrayList<>();
        File file = new File(fileName);
        if (file.exists()) {
          if (file.isDirectory()) {
              try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(Paths.get(fileName))) {
                  for (Path path : directoryStream) {
                      if (isRegularFile(path)) {
                          aFiles.add(path.getFileName());
                      }
                  }
              }
              catch (IOException e) {
                  Main.logr.log(System.Logger.Level.ERROR, "error.E014", "E014. Error reading directory {0}: {1}", fileName, e.getMessage());
              }
          }
          else {
              aFiles.add(file.toPath());
          }
        }
        return aFiles;
    }
}
