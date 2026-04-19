package ru.bis.javautil.xlsparse;

public class ParserFactory {
    // Static factory for parser
    public static AParser getParser(StatementType type) {
        AParser parser = null;
        System.out.println("Statement type: " + type.getName());
        if (type == StatementType.BTB) {
            parser = new BTBParser();
        }
        return parser;
    }
}
