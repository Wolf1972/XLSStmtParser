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

    static String long2str(long amount) { // Converts long value to string with 2 digital digits separated
        return String.format("%d" + Util.dSep + "%02d", amount / 100, amount % 100);
    }

    static long str2long(String str) throws NumberFormatException { // Converts string with 2 decimal digits separated to long, possible NBSP between digits (1 000 000.00)
        if (str == null || str.isEmpty()) {
            return 0;
        }
        else {
            return Long.parseLong(str.replaceAll("[ .," + nbsp + "]", ""));
        }
    }

    static String cleanStr(String str) { // Clean string from \n, \r, replaces nbsp to ordinary spaces
        if (str == null) {
            return "";
        } else {
            return str.replace("\n", "").replace("\r", "").replaceAll(nbsp, " ");
        }
    }

    static String str2CSV(String str) { // Converts string 2 CSV-clean string (double quota replaced by "")
       if (str == null) {
           return "";
       }
       else {
           return str.replace("\"", "\"\"");
       }
    }

    static String changeFileExtension(String fileName, String ext) { // Changes file extension (e.g. *.xls to *.csv)
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

    static String leftStr(String str, int max) { // Restricts string with some length even if this length more than length of the real string (instead Apache Commons)
        return str == null ? "" : str.length() > max ? str.substring(0, max) : str;
    }

    static String resource(String code, String defaultString, Object... args) { // Returns resource string or default string
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

    static List<Path> getFileList(String fileName) { // Returns list of files, in case inFile is a directory returns all files from it
        List<Path> aFiles = new ArrayList<>();
        File file = new File(fileName);
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
        return aFiles;
    }
}
