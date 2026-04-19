package ru.bis.javautil.xlsparse;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.stream.Stream;

public class Util {
    static String lSep; // line separator
    static String dSep; // decimal separator
    static String fSep; // field separator

    static final String nbsp = "\u00A0";

    static void init() {
        DecimalFormat format = (DecimalFormat) DecimalFormat.getInstance();
        DecimalFormatSymbols symbols = format.getDecimalFormatSymbols();
        dSep = String.valueOf(symbols.getDecimalSeparator());
        System.out.println("System decimal separator: " + dSep);

        lSep = System.lineSeparator();
        System.out.print("System line separator: ");
        Stream<Character> sch = lSep.chars().mapToObj(i -> (char)i);
        sch.forEach(ch -> System.out.printf("#%d ", (int) ch));
        System.out.println();

        fSep = ";";
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
}
