package ru.bis.javautil.xlsparse;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class Logger {
    static private int minLogLevelWeight; // Weight of current min log level
    static final private System.Logger.Level[] levelWeights = { // Log level weights
            System.Logger.Level.DEBUG,
            System.Logger.Level.TRACE,
            System.Logger.Level.INFO,
            System.Logger.Level.WARNING,
            System.Logger.Level.ERROR
    };
    static DateTimeFormatter logDtFormatter = DateTimeFormatter.ofPattern("dd.MM.yy HH:mm:ss");
    static String logTmpFileName = "log.tmp";
    static String logFileName = "current.log";
    static BufferedWriter logWriter = null;

    /**
     * Constructor
     * @param minLevel - min log level
     */
    Logger(System.Logger.Level minLevel) {
        setMinLogLevel(minLevel);
        try {
            Charset chs = StandardCharsets.UTF_8;
            OutputStream os = new FileOutputStream("current.log");
            logWriter = new BufferedWriter(new OutputStreamWriter(os, chs));
        }
        catch (Exception e) {
            log(System.Logger.Level.ERROR, Util.resource("error.E016", "E016. Error creating log file: {0}: {1}", logTmpFileName, e.getMessage()));
        }
    }

    int getLogLevelWeight(System.Logger.Level level) {
        int weight = -1;
        for (int i = 0; i < levelWeights.length; i++) {
            if (levelWeights[i] == level) {
                weight = i;
                break;
            }
        }
        return weight;
    }

    void setMinLogLevel(System.Logger.Level level) {
        Logger.minLogLevelWeight = getLogLevelWeight(level);
    }

    System.Logger.Level getMinLogLevel() {
        return Logger.levelWeights[minLogLevelWeight];
    }

    /**
     * Raw log output
     * @param level - log level
     * @param msg - message
     */
    void log(System.Logger.Level level, String msg) {
        System.out.println(LocalDateTime.now().format(logDtFormatter) + " " + level.getName() + ": " + msg);
        if (logWriter != null) {
            try {
                logWriter.write(LocalDateTime.now().format(logDtFormatter) + " " + level.getName() + ": " + msg + Util.lSep);
            }
            catch (Exception e) {
                System.out.println(Util.resource("error.E017", "E017. Error writing log file: {0}", e.getMessage()));
            }
        }
    }

    /**
     * Outputs log message from resource with arguments
     * @param level - ERROR, WARNING, INFO
     * @param resource - resource bundle id
     * @param msg - using when resource not found
     * @param args - multiple arguments
     */
    void log(System.Logger.Level level, String resource, String msg, Object... args) {
        if (minLogLevelWeight <= getLogLevelWeight(level)) {
            String message;
            try {
                message = Main.bundle.getString(resource);
            } catch (Exception e) {
                message = msg;
            }
            log(level, MessageFormat.format(message, args));
        }
    }

    /**
     * Must be call when application close to close log file
     */
    void close() {
        if (logWriter != null) {
            try {
                logWriter.close();
            }
            catch (Exception e) {
                log(System.Logger.Level.ERROR, Util.resource("error.E018", "E018. Error closing log file: {0}: {1}", logTmpFileName, e.getMessage()));
            }
            logWriter = null;
            File tmpLog = new File(logTmpFileName);
            if (tmpLog.exists()) {
                if (!tmpLog.renameTo(new File(logFileName))) {
                    log(System.Logger.Level.ERROR, Util.resource("error.E019", "E019. Error renaming log file: {0} to {1}", logTmpFileName, logFileName));
                }
            }
        }
    }

}
