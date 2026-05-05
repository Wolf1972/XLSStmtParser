package ru.bis.javautil.xlsparse;

import org.apache.commons.cli.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import static java.nio.file.Files.isRegularFile;

public class Main {
    
    // TODO: automatic Excel file type defining

    public static void main(String[] args) {
        System.out.println("XLS Statement parser.");
        Util.init();

        String inFileName = "stmt.xls";
        String outFileName = "out.csv";
        String arcFileName = "";
        String lineSeparatorCommand = "";
        String fieldSeparatorCommand = "";
        String decimalSeparatorCommand = "";
        String errorHandleCommand = "";
        String xlsTypeStr = "0";

        String stmtType = "1";     // Statement type
        String codePage = "UTF-8"; // Code page
        XLSType xlsType = XLSType.AUTO; // Old XLS format by default
        String dateFormat = "yyyy-MM-dd"; // Output date format
        int errorHandle = 1; // Error handling (when error leads to fail operation): 0 - all errors, 1 - only format errors (before parsing), 2 - no handling

        CommandLineParser parser = new DefaultParser();
        Options options = makeCmdOptions();
        try {
            CommandLine command = parser.parse(options, args);

            if (command.hasOption('i')) inFileName = command.getOptionValue('i');
            if (command.hasOption('o')) outFileName = command.getOptionValue('o');
            if (command.hasOption('a')) arcFileName = command.getOptionValue('a');
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
            if ("1".equals(xlsTypeStr)) {
                xlsType = XLSType.XLS;
            }
            else if ("2".equals(xlsTypeStr)) {
                xlsType = XLSType.XLSX;
            }
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

        if (Util.dSep.equals(Util.fSep)) {
            System.out.println("E010. Field separator is set equal with decimal separator (\"" + Util.dSep + "\"). Unable to create correct CSV file.");
            return;
        }

        File inFile = new File(inFileName);
        File outFile = new File(outFileName);
        if ((!inFile.isDirectory() && outFile.isDirectory()) || (inFile.isDirectory() && !outFile.isDirectory())) {
            System.out.println(inFile.isDirectory());
            System.out.println(outFile.isDirectory());
            System.out.println("E011. Input and output parameters must be only directories or only files simultaneously.");
            return;
        }
        if (inFile.isDirectory()) {
            if (!"/\\".contains(inFileName.substring(inFileName.length() - 1))) { // Add last "\" or "/" to arc directory name, if needed
                inFileName += Util.fileSep;
            }
        }
        if (outFile.isDirectory()) {
            if (!"/\\".contains(outFileName.substring(outFileName.length() - 1))) { // Add last "\" or "/" to arc directory name, if needed
                outFileName += Util.fileSep;
            }
        }

        AParser stmtParser;
        if (stmtType.isEmpty() || stmtType.equals("1")) {
            stmtParser = ParserFactory.getParser(StatementType.BTB);
        }
        else {
            System.out.println("E012. Unknown statement type: " + stmtType);
            return;
        }

        List<Path> aFiles = getFileList(inFileName);

        int iProceessed = 0;
        int iSuccess = 0;

        for (Path file : aFiles) {
            File next = file.toFile();
            String nextInFileName = inFileName + next.getName();
            String nextOutFileName = outFileName;
            String nextArcFileName = arcFileName;

            if (outFile.isDirectory()) {
                nextOutFileName = outFileName + Util.changeFileExtension(next.getName(), "csv");
            }
            if (!arcFileName.isEmpty()) { // Is archive need?
                File arc = new File(arcFileName);
                if (arc.isDirectory()) {
                    LocalDateTime now = LocalDateTime.now();
                    String timeStamp = now.format(DateTimeFormatter.ofPattern("yyMMdd_HHmmss"));
                    if (!"/\\".contains(arcFileName.substring(arcFileName.length() - 1))) { // Add last "\" or "/" to arc directory name, if needed
                        arcFileName += Util.fileSep;
                    }
                    nextArcFileName = arcFileName + timeStamp + "_" + next.getName();
                }
            }

            iProceessed++;

            System.out.println();
            System.out.println("Input statement file: " + nextInFileName);
            if (stmtParser.process(nextInFileName, xlsType, nextOutFileName, codePage, errorHandle, nextArcFileName)) {
                System.out.println("Output file created: " + nextOutFileName +
                        (arcFileName.isEmpty() ? "" : " Archive file created: " + nextArcFileName));
                iSuccess++;
            }
        }
        System.out.println();
        System.out.println("Processed " + iProceessed + " file(s), successful " + iSuccess + " file(s).");
    }

    static Options makeCmdOptions() {
        Options options = new Options();
        options.addRequiredOption("i", "input", true, "* Input XLS file or directory, required");
        options.addRequiredOption("o", "output", true, "* Output CSV file or directory, required");
        options.addOption("a", "archive", true, "Archive file or directory, no archivation by default");
        options.addOption("s", "stmt-type", true, "Statement type (1 - BTB Bank), 1 by default");
        options.addOption("l", "line-separator", true, "Line separator (\"n\" or \"rn\"), system separator by default");
        options.addOption("f", "field-separator", true, "Field separator, \";\" by default");
        options.addOption("n", "numeric-separator", true, "Numeric separator, \".\" or \",\", system separator by default");
        options.addOption("c", "codepage", true, "Output file in specified code page, default UTF-8");
        options.addOption("x", "xls-type", true, "XLS file type (0 - auto, 1 - XLS, 2 - XLSX), 0 by default");
        options.addOption("d", "date-format", true, "Output date format, YYYY-MM-DD by default");
        options.addOption("e", "error-handling", true, "Error handling (when error leads to fail operation): 0 - fail when any error, 1 - fail when only format errors (before parsing), 2 - try to no fail");
        return options;
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
                System.out.println("E012. Error reading directory: " + fileName + ": " + e.getMessage());
            }
        }
        else {
            aFiles.add(file.toPath());
        }
        return aFiles;
    }
}
