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
import java.util.*;

import static java.nio.file.Files.isRegularFile;

public class Main {

    static final Logger logr = new Logger(System.Logger.Level.INFO);
    static Locale locale = Locale.of("ru");
    static ResourceBundle bundle = null;

    public static void main(String[] args) {try {
            bundle = ResourceBundle.getBundle("strings", locale);
        }
        catch (Exception e) {
            // One message without resources
            logr.log(System.Logger.Level.WARNING, "E000. Can't load resource file." + e.getMessage());
        }

        logr.log(System.Logger.Level.INFO, Util.resource("info.title", "XLS Statement parser"));
        final Properties properties = new Properties();
        try {
            properties.load(Main.class.getClassLoader().getResourceAsStream("project.properties"));
            final String version = properties.getProperty("version");
            logr.log(System.Logger.Level.INFO, "info.ver", "Version {0}", version);
        }
        catch (final Exception e) {
            logr.log(System.Logger.Level.ERROR, "error.E001", "E001. Can't read properties file: {0}", e.getMessage());
        }

        String inFileName = "stmt.xls";
        String outFileName = "out.csv";
        String arcFileName = "";
        String stmtTypeCommand = "";
        String xlsTypeCommand = "0";
        String lineSeparatorCommand = "";
        String fieldSeparatorCommand = "";
        String decimalSeparatorCommand = "";
        String errorHandleCommand = "";
        String logLevelCommand = "";

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
            if (command.hasOption('x')) xlsTypeCommand = command.getOptionValue('x');
            if (command.hasOption('d')) dateFormat = command.getOptionValue('d');
            if (command.hasOption('p')) errorHandleCommand = command.getOptionValue('p');
            if (command.hasOption('e')) logLevelCommand = command.getOptionValue('e');

        }
        catch (ParseException e) {
            logr.log(System.Logger.Level.ERROR, "error.E002","E002. Invalid command line: {0}", e.getMessage());
            HelpFormatter help = new HelpFormatter();
            help.printHelp(Main.class.getSimpleName(), options);
            return;
        }

        Util.init();

        if (!logLevelCommand.isEmpty()) {
            for (System.Logger.Level type : System.Logger.Level.values()) {
                if (type.name().equals(logLevelCommand)) {
                    logr.setMinLogLevel(type);
                    break;
                }
            }
            logr.log(System.Logger.Level.TRACE,"trace.log_level", "Log level is set: {0}", logr.getMinLogLevel().name());
        }

        if (!lineSeparatorCommand.isEmpty()) {
            if (lineSeparatorCommand.charAt(0) == 'r') Util.lSep = "\r";
            else if (lineSeparatorCommand.charAt(0) == 'n') Util.lSep = "\n";
            if (lineSeparatorCommand.length() > 1) {
                if (lineSeparatorCommand.charAt(1) == 'r') Util.lSep += "\r";
                else if (lineSeparatorCommand.charAt(1) == 'n') Util.lSep += "\n";
            }
            logr.log(System.Logger.Level.TRACE, "trace.line_separator", "System line separator was override: {0}", lineSeparatorCommand);
        }

        if (!xlsTypeCommand.isEmpty()) {
            for (XLSType type : XLSType.values()) {
                if (type.name().equals(xlsTypeCommand)) {
                    xlsType = type;
                    break;
                }
            }
            logr.log(System.Logger.Level.TRACE,"trace.excel_format", "Excel file format is set: {0}", xlsType.name());
        }

        if (!fieldSeparatorCommand.isEmpty()) {
            Util.fSep = fieldSeparatorCommand;
            logr.log(System.Logger.Level.TRACE, "trace.field_separator", "Field separator is set: {0}", fieldSeparatorCommand);
        }

        if (!decimalSeparatorCommand.isEmpty()) {
            Util.dSep = decimalSeparatorCommand;
            logr.log(System.Logger.Level.TRACE, "trace.decimal_separator", "Decimal separator is set: {0}", decimalSeparatorCommand);
        }

        if (!dateFormat.isEmpty()) {
            Util.outDateFormat = dateFormat;
            logr.log(System.Logger.Level.TRACE, "trace.date_format","Date format is set: {0}", dateFormat);
        }

        if (!errorHandleCommand.isEmpty()) {
            for (ErrHandleStrategy strategy : ErrHandleStrategy.values()) {
                if (strategy.name().equals(errorHandleCommand)) {
                    errorHandle = strategy;
                    break;
                }
            }
            logr.log(System.Logger.Level.TRACE, "trace.error_handle","Process termination strategy: {0}", errorHandle.name());
        }

        if (Util.dSep.equals(Util.fSep)) {
            logr.log(System.Logger.Level.ERROR, "error.E010","E010. Field separator is set equal with decimal separator ({0}}). Unable to create correct CSV file.", Util.dSep);
            return;
        }

        File inFile = new File(inFileName);
        File outFile = new File(outFileName);
        if (inFileName.equals(outFileName)) {
            logr.log(System.Logger.Level.ERROR,"error.E011", "E011. Input and output parameters must be different.");
            return;
        }
        if ((!inFile.isDirectory() && outFile.isDirectory()) || (inFile.isDirectory() && !outFile.isDirectory())) {
            logr.log(System.Logger.Level.ERROR,"error.E012", "E012. Input and output parameters must be only directories or must be only files simultaneously.");
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
            logr.log(System.Logger.Level.TRACE,"trace.statement_type", "Statement type is set: {0}", statementType.name());
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
                    logr.log(System.Logger.Level.INFO, "info.space", "");
                    logr.log(System.Logger.Level.INFO, "info.input_file","Input statement file: {0}", nextInFileName);
                }
                if (stmtParser.process(nextInFileName, xlsType, nextOutFileName, codePage, errorHandle, nextArcFileName)) {
                    logr.log(System.Logger.Level.INFO,"info.output_file", "Output file created: {0}", nextOutFileName);
                    if (!arcFileName.isEmpty()) {
                        logr.log(System.Logger.Level.INFO,"info.arc_file", "Archive file created: {0}", nextArcFileName);
                    }
                    iSuccess++;
                }
            }
            if (iSuccess > 0) {
                logr.log(System.Logger.Level.INFO, "info.space","");
            }
            logr.log(System.Logger.Level.INFO, "info.total_files", "Processed {0} file(s), successful {1} file(s).", iProceessed, iSuccess);
        }
        else {
            logr.log(System.Logger.Level.ERROR, "error.E013", "E013. There is no parser available for statement type: {0}", statementType.name());
        }
    }

    static Options makeCmdOptions() {
        Options options = new Options();
        options.addRequiredOption("i", "input", true, Util.resource("cmd.i", "* Input XLS file or directory, required"));
        options.addRequiredOption("o", "output", true, Util.resource("cmd.o", "* Output CSV file or directory, required"));
        options.addOption("a", "archive", true, Util.resource("cmd.a", "Archive file or directory, no archivation by default"));
        options.addOption("s", "stmt-type", true, Util.resource("cmd.s", "Statement type (BTB - BTB Bank), BTB by default"));
        options.addOption("l", "line-separator", true, Util.resource("cmd.l", "Line separator (\"n\" or \"rn\"), system separator by default"));
        options.addOption("f", "field-separator", true, Util.resource("cmd.f", "Field separator, \";\" by default"));
        options.addOption("n", "numeric-separator", true, Util.resource("cmd.n", "Numeric separator, \".\" or \",\", system separator by default"));
        options.addOption("c", "codepage", true, Util.resource("cmd.c", "Output file in specified code page, UTF-8 by default"));
        options.addOption("x", "xls-type", true, Util.resource("cmd.x", "XLS file type (AUTO - auto defining, XLS only, XLSX only), AUTO by default"));
        options.addOption("d", "date-format", true, Util.resource("cmd.d", "Output date format, YYYY-MM-DD by default"));
        options.addOption("p", "process-termination", true, Util.resource("cmd.p", "Process termination (when error leads to fail operation): ALL - fail when any error, FORMAT - fail when only format errors (before parsing), NONE - try not to fail when any error"));
        options.addOption("e", "log-level", true, Util.resource("cmd.e", "Log level (DEBUG, TRACE, INFO, WARNING, ERROR), INFO by default"));
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
                logr.log(System.Logger.Level.ERROR, "error.E014", "E014. Error reading directory {0}: {1}", fileName, e.getMessage());
            }
        }
        else {
            aFiles.add(file.toPath());
        }
        return aFiles;
    }
}
