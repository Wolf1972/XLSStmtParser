package ru.bis.javautil.xlsparse;

import java.text.MessageFormat;

public class Logger {
    static private int minLogLevelWeight; // Weight of current min log level
    static final private System.Logger.Level[] levelWeights = { // Log level weights
            System.Logger.Level.DEBUG,
            System.Logger.Level.TRACE,
            System.Logger.Level.INFO,
            System.Logger.Level.WARNING,
            System.Logger.Level.ERROR
    };

    Logger(System.Logger.Level minLevel) {
        setMinLogLevel(minLevel);
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

    void log(System.Logger.Level level, String msg) { // Raw log output
        System.out.println(level.getName() + " : " + msg);
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
}
