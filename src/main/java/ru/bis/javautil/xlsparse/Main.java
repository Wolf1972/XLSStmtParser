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

    static final System.Logger logger = System.getLogger("ru.bis.javautil.xlsparse");

    public static void main(String[] args) {

        logger.log(System.Logger.Level.INFO, "XLS Statement parser.");

        String inFileName = "stmt.xls";
        String outFileName = "out.csv";
        String arcFileName = "";
        String stmtTypeCommand = "";
        String lineSeparatorCommand = "";
        String fieldSeparatorCommand = "";
        String decimalSeparatorCommand = "";
        String errorHandleCommand = "";
        String xlsTypeStr = "0";

        String codePage = "UTF-8"; // Code page
        StatementType statementType = StatementType.BTB;
        XLSType xlsType = XLSType.AUTO; // Old XLS format by default
        String dateFormat = "yyyy-MM-dd"; // Output date format
        ErrHandleStrategy errorHandle = ErrHandleStrategy.FORMAT; // Error handling (when error leads to fail operation): ALL, FORMAT, NONE

        CommandLineParser parser = new DefaultParser();
        Options options = makeCmdOptions();
        try {
            CommandLine command = parser.parse(options, args);

            if (command.hasOption('i')) inFileName = command.getOptionValue('i');
            if (command.hasOption('o')) outFileName = command.getOptionValue('o');
            if (command.hasOption('a')) arcFileName = command.getOptionValue('a');
            if (command.hasOption('s')) stmtTypeCommand = command.getOptionValue('s');
            if (command.hasOption('l')) lineSeparatorCommand = command.getOptionValue('l');
            if (command.hasOption('f')) fieldSeparatorCommand = command.getOptionValue('f');
            if (command.hasOption('n')) decimalSeparatorCommand = command.getOptionValue('n');
            if (command.hasOption('c')) codePage = command.getOptionValue('c');
            if (command.hasOption('x')) xlsTypeStr = command.getOptionValue('x');
            if (command.hasOption('d')) dateFormat = command.getOptionValue('d');
            if (command.hasOption('p')) errorHandleCommand = command.getOptionValue('p');

        }
        catch (ParseException e) {
            logger.log(System.Logger.Level.ERROR, "E000. Invalid command line.", e.getMessage());
            HelpFormatter help = new HelpFormatter();
            help.printHelp(Main.class.getSimpleName(), options);
            return;
        }

        Util.init();

        if (!lineSeparatorCommand.isEmpty()) {
            if (lineSeparatorCommand.charAt(0) == 'r') Util.lSep = "\r";
            else if (lineSeparatorCommand.charAt(0) == 'n') Util.lSep = "\n";
            if (lineSeparatorCommand.length() > 1) {
                if (lineSeparatorCommand.charAt(1) == 'r') Util.lSep += "\r";
                else if (lineSeparatorCommand.charAt(1) == 'n') Util.lSep += "\n";
            }
            logger.log(System.Logger.Level.INFO, "System line separator was override: " + lineSeparatorCommand);
        }

        if (!xlsTypeStr.isEmpty()) {
            for (XLSType type : XLSType.values()) {
                if (type.name().equals(xlsTypeStr)) {
                    xlsType = type;
                    break;
                }
            }
            logger.log(System.Logger.Level.INFO,"Excel file format is set: " + xlsType.name());
        }

        if (!fieldSeparatorCommand.isEmpty()) {
            Util.fSep = fieldSeparatorCommand;
            logger.log(System.Logger.Level.INFO, "Field separator is set: " + fieldSeparatorCommand);
        }

        if (!decimalSeparatorCommand.isEmpty()) {
            Util.dSep = decimalSeparatorCommand;
            logger.log(System.Logger.Level.INFO, "Decimal separator is set: " + decimalSeparatorCommand);
        }

        if (!dateFormat.isEmpty()) {
            Util.outDateFormat = dateFormat;
            logger.log(System.Logger.Level.INFO, "Date format is set: " + dateFormat);
        }

        if (!errorHandleCommand.isEmpty()) {
            for (ErrHandleStrategy strategy : ErrHandleStrategy.values()) {
                if (strategy.name().equals(errorHandleCommand)) {
                    errorHandle = strategy;
                    break;
                }
            }
            logger.log(System.Logger.Level.INFO, "Process termination strategy: " + errorHandle.name());
        }

        if (Util.dSep.equals(Util.fSep)) {
            logger.log(System.Logger.Level.ERROR, "E010. Field separator is set equal with decimal separator (\"" + Util.dSep + "\"). Unable to create correct CSV file.");
            return;
        }

        File inFile = new File(inFileName);
        File outFile = new File(outFileName);
        if (inFileName.equals(outFileName)) {
            logger.log(System.Logger.Level.ERROR,"E011. Input and output parameters must be different.");
            return;
        }
        if ((!inFile.isDirectory() && outFile.isDirectory()) || (inFile.isDirectory() && !outFile.isDirectory())) {
            logger.log(System.Logger.Level.ERROR,"E012. Input and output parameters must be only directories or must be only files simultaneously.");
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

        if (!stmtTypeCommand.isEmpty()) {
            for (StatementType stmtType : StatementType.values()) {
                if (stmtType.name().equals(stmtTypeCommand)) {
                    statementType = stmtType;
                    break;
                }
            }
            logger.log(System.Logger.Level.INFO,"Statement type is set: " + statementType.name());
        }

        AParser stmtParser = null;
        if (statementType == StatementType.BTB) {
            stmtParser = ParserFactory.getParser(StatementType.BTB);
        }

        if (stmtParser != null) {
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

                if (iProceessed > 0) {
                    logger.log(System.Logger.Level.INFO, "");
                    logger.log(System.Logger.Level.INFO, "Input statement file: " + nextInFileName);
                }
                if (stmtParser.process(nextInFileName, xlsType, nextOutFileName, codePage, errorHandle, nextArcFileName)) {
                    logger.log(System.Logger.Level.INFO, "Output file created: " + nextOutFileName +
                            (arcFileName.isEmpty() ? "" : " Archive file created: " + nextArcFileName));
                    iSuccess++;
                }
            }
            if (iSuccess > 0) {
                logger.log(System.Logger.Level.INFO, "");
            }
            logger.log(System.Logger.Level.INFO, "Processed " + iProceessed + " file(s), successful " + iSuccess + " file(s).");
        }
        else {
            logger.log(System.Logger.Level.ERROR, "E013. There is no parser available for statement type: " + statementType.name());
        }
    }

    static Options makeCmdOptions() {
        Options options = new Options();
        options.addRequiredOption("i", "input", true, "* Input XLS file or directory, required");
        options.addRequiredOption("o", "output", true, "* Output CSV file or directory, required");
        options.addOption("a", "archive", true, "Archive file or directory, no archivation by default");
        options.addOption("s", "stmt-type", true, "Statement type (BTB - BTB Bank), BTB by default");
        options.addOption("l", "line-separator", true, "Line separator (\"n\" or \"rn\"), system separator by default");
        options.addOption("f", "field-separator", true, "Field separator, \";\" by default");
        options.addOption("n", "numeric-separator", true, "Numeric separator, \".\" or \",\", system separator by default");
        options.addOption("c", "codepage", true, "Output file in specified code page, default UTF-8");
        options.addOption("x", "xls-type", true, "XLS file type (AUTO - auto defining, XLS only, XLSX only), AUTO by default");
        options.addOption("d", "date-format", true, "Output date format, YYYY-MM-DD by default");
        options.addOption("p", "process-termination", true, "Process termination (when error leads to fail operation): ALL - fail when any error, FORMAT - fail when only format errors (before parsing), NONE - try not to fail when any error");
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
                logger.log(System.Logger.Level.ERROR, "E014. Error reading directory: " + fileName + ": ", e);
            }
        }
        else {
            aFiles.add(file.toPath());
        }
        return aFiles;
    }
}
