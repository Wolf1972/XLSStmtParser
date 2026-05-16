package ru.bis.javautil.xlsparse;

public class ParserFactory {
    // Static factory for parser
    public static AParser getParser(StatementType type) {
        AParser parser = null;
        Main.logr.log(System.Logger.Level.INFO, "info.parser","Statement parser: {0}", type.getName());
        if (type == StatementType.BTB) {
            parser = new BTBParser();
        }
        return parser;
    }
}
