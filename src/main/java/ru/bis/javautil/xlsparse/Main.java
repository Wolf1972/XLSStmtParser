package ru.bis.javautil.xlsparse;

import org.apache.commons.cli.*;

public class Main {

    // TODO: output date format parameter
    // TODO: "behaviour when error" switch
    // TODO: group of files processing

    public static void main(String[] args) {
        System.out.println("XLS Statement parser.");
        Util.init();

        String inFileName = "stmt.xls";
        String outFileName = "out.csv";
        String arcDirectory = "";
        String lineSeparatorCommand = "";
        String fieldSeparatorCommand = "";
        String decimalSeparatorCommand = "";
        String errorHandleCommand = "";
        String xlsTypeStr = "0";

        String stmtType = "1";     // Statement type
        String codePage = "UTF-8"; // Code page
        XLSType xlsType = XLSType.XLS; // Old XLS format by default
        String dateFormat = "yyyy-MM-dd"; // Output date format
        int errorHandle = 1; // Error handling (when error leads to fail operation): 0 - all errors, 1 - only format errors (before parsing), 2 - no handling

        CommandLineParser parser = new DefaultParser();
        Options options = makeCmdOptions();
        try {
            CommandLine command = parser.parse(options, args);

            if (command.hasOption('i')) inFileName = command.getOptionValue('i');
            if (command.hasOption('o')) outFileName = command.getOptionValue('o');
            if (command.hasOption('a')) arcDirectory = command.getOptionValue('a');
            if (command.hasOption('s')) stmtType = command.getOptionValue('s');
            if (command.hasOption('l')) lineSeparatorCommand = command.getOptionValue('l');
            if (command.hasOption('f')) fieldSeparatorCommand = command.getOptionValue('f');
            if (command.hasOption('n')) decimalSeparatorCommand = command.getOptionValue('n');
            if (command.hasOption('c')) codePage = command.getOptionValue('c');
            if (command.hasOption('x')) xlsTypeStr = command.getOptionValue('x');
            if (command.hasOption('d')) dateFormat = command.getOptionValue('d');
            if (command.hasOption('e')) errorHandleCommand = command.getOptionValue('e');

        }
        catch (ParseException e) {
            System.out.println("E000. Invalid command line.");
            HelpFormatter help = new HelpFormatter();
            help.printHelp(Main.class.getSimpleName(), options);
            return;
        }

        if (!lineSeparatorCommand.isEmpty()) {
            if (lineSeparatorCommand.charAt(0) == 'r') Util.lSep = "\r";
            else if (lineSeparatorCommand.charAt(0) == 'n') Util.lSep = "\n";
            if (lineSeparatorCommand.length() > 1) {
                if (lineSeparatorCommand.charAt(1) == 'r') Util.lSep += "\r";
                else if (lineSeparatorCommand.charAt(1) == 'n') Util.lSep += "\n";
            }
        }

        if (!"0".equals(xlsTypeStr)) {
            xlsType = XLSType.XLSX;
        }

        if (!fieldSeparatorCommand.isEmpty()) {
            Util.fSep = fieldSeparatorCommand;
        }

        if (!decimalSeparatorCommand.isEmpty()) {
            Util.dSep = decimalSeparatorCommand;
        }

        if (!dateFormat.isEmpty()) {
            Util.outDateFormat = dateFormat;
        }

        if (!errorHandleCommand.isEmpty()) {
            try {
                errorHandle = Integer.parseInt(errorHandleCommand);
            }
            catch (NumberFormatException e) {
                System.out.println("E011. Unknown error handling strategy.");
            }
        }

        if (!arcDirectory.isEmpty()) {
            if (!"/\\".contains(arcDirectory.substring(arcDirectory.length() - 1))) { // Add last "\" or "/" to arc directory name, if needed
                arcDirectory += Util.fileSep;
            }
        }

        if (Util.dSep.equals(Util.fSep)) {
            System.out.println("E010. Field separator is set equal with decimal separator (\"" + Util.dSep + "\"). Unable to create correct CSV file.");
            return;
        }

        if (stmtType.isEmpty() || stmtType.equals("1")) {
            AParser stmtParser = ParserFactory.getParser(StatementType.BTB);
            if (stmtParser != null) {
                if (stmtParser.process(inFileName, xlsType, outFileName, codePage, errorHandle, arcDirectory)) {
                    System.out.println("Input statement file " + inFileName + " was processed successful. Output file: " + outFileName);
                }
            }
        }
    }

    static Options makeCmdOptions() {
        Options options = new Options();
        options.addRequiredOption("i", "input", true, "Input PDF file, required");
        options.addRequiredOption("o", "output", true, "Output CSV file, required");
        options.addOption("a", "arc-dir", true, "Archival directory, no archive by default");
        options.addOption("s", "stmt-type", true, "Statement type (1 - BTB Bank), 1 by default");
        options.addOption("l", "line-separator", true, "Line separator (\"n\" or \"rn\"), system separator by default");
        options.addOption("f", "field-separator", true, "Field separator, \";\" by default");
        options.addOption("n", "numeric-separator", true, "Numeric separator, \".\" or \",\", system separator by default");
        options.addOption("c", "codepage", true, "Output file in specified code page, default UTF-8");
        options.addOption("x", "xls-type", true, "XLS file type (0 - XLS, 1 - XLSX), 0 by default");
        options.addOption("d", "date-format", true, "Output date format, YYYY-MM-DD by default");
        options.addOption("e", "error-handling", true, "Error handling (when error leads to fail operation): 0 - fail when any error, 1 - fail when only format errors (before parsing), 2 - try to no fail");
        return options;
    }
}
