package ru.bis.javautil.xlsparse;

import org.apache.commons.cli.*;

public class Main {

    // TODO: XLS file type switch
    // TODO: behaviour when error switch
    // TODO: group of files processing

    public static void main(String[] args) {
        System.out.println("XLS Statement parser.");
        Util.init();

        String inFileName = "stmt.xls";
        String outFileName = "out.csv";
        String lineSeparatorCommand = "";
        String fieldSeparatorCommand = "";
        String decimalSeparatorCommand = "";

        String stmtType = "1";     // Statement type
        String codePage = "UTF-8"; // Code page

        CommandLineParser parser = new DefaultParser();
        Options options = makeCmdOptions();
        try {
            CommandLine command = parser.parse(options, args);

            if (command.hasOption('i')) inFileName = command.getOptionValue('i');
            if (command.hasOption('o')) outFileName = command.getOptionValue('o');
            if (command.hasOption('s')) stmtType = command.getOptionValue('s');
            if (command.hasOption('l')) lineSeparatorCommand = command.getOptionValue('l');
            if (command.hasOption('f')) fieldSeparatorCommand = command.getOptionValue('f');
            if (command.hasOption('d')) decimalSeparatorCommand = command.getOptionValue('d');
            if (command.hasOption('c')) codePage = command.getOptionValue('c');
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

        if (!fieldSeparatorCommand.isEmpty()) {
            Util.fSep = fieldSeparatorCommand;
        }

        if (!decimalSeparatorCommand.isEmpty()) {
            Util.dSep = decimalSeparatorCommand;
        }

        if (Util.dSep.equals(Util.fSep)) {
            System.out.println("E009. Field separator is set equal with decimal separator (\"" + Util.dSep + "\"). Unable to create correct CSV file.");
            return;
        }

        if (stmtType.isEmpty() || stmtType.equals("1")) {
            AParser stmtParser = ParserFactory.getParser(StatementType.BTB);
            if (stmtParser != null) {
                if (stmtParser.process(inFileName, XLSType.XLS, outFileName, codePage)) {
                    System.out.println("Input statement file " + inFileName + " was processed successful. Output file: " + outFileName);
                }
            }
        }
    }

    static Options makeCmdOptions() {
        Options options = new Options();
        options.addRequiredOption("i", "input", true, "Input PDF file, required");
        options.addRequiredOption("o", "output", true, "Output CSV file, required");
        options.addOption("s", "stmt-type", true, "Statement type (1 - BTB Bank), 1 by default");
        options.addOption("l", "line-separator", true, "Line separator (\"n\" or \"rn\"), system separator by default");
        options.addOption("f", "field-separator", true, "Field separator, \";\" by default");
        options.addOption("d", "decimal-separator", true, "Decimal separator, \".\" or \",\", system separator by default");
        options.addOption("c", "codepage", true, "Output file in specified code page, default UTF-8");
        return options;
    }
}
