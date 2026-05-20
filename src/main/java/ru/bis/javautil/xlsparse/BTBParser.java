package ru.bis.javautil.xlsparse;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Properties;

public class BTBParser extends AParser {

    private String dateFormat = "dd.MM.yyyy";
    private int lastHeaderRow = 9; // Last table header row, minimum rows in statement
    private int trailerRows = 6; // Trailer rows count
    private int firstColumn = 2; // First column number
    private String firstColumnName = "ДатаСоздания"; // Name is used for format check
    private String incomingRestName = "ВходящийОстаток"; // Name is used for incoming rest
    private String outgoingRestName = "ИсходящийОстаток:"; // Name is used for outgoing rest
    private String opNumColumnName = "№операции"; // Name is used for operation number (columns positioning)
    private int accountRow = 4; // Row with account number

    private int headerNameColumn = 1; // Column with header names
    private int headerValueColumn = 6; // Column with header values

    private int trailerNameColumn = 2; // Column with trailer names
    private int trailerValueColumn = 6; // Column with trailer values

    private int dtTurnoverRowDistance = 4; // Distance from last row for row with debit turnover
    private int crTurnoverRowDistance = 3;
    private int inRestRow = 8; // Row for incoming rest (may be missed)
    private int outRestRowDistance = 2; // Distance from last row for row with outgoing rest (may be missed)

    private int opNumColumn = 7; // Operation number column
    private int opDateColumn = 2; // Operation date column
    private int opValueColumn = 3; // Operation value column
    private int ctrPartAccountColumn = 9; // Counterparty account column
    private int ctrPartNameColumn = 11; // Counterparty name column
    private int dtAmountColumn = 13; // Debit amount column
    private int crAmountColumn = 14; // Credit amount column
    private int purposeColumn = 16; // Purpose column

    DateTimeFormatter formatterDt = DateTimeFormatter.ofPattern(dateFormat);

    /**
     * This method for overriding markup parameters
     */
    void init() {
        Properties props = new Properties();
        String iniFileName = "btb.ini";
        try (Reader reader = new InputStreamReader(new FileInputStream(iniFileName), StandardCharsets.UTF_8)) {
            props.load(reader);

            lastHeaderRow = Integer.parseInt(props.getProperty("lastHeaderRow", String.valueOf(lastHeaderRow))); // Last table header row, minimum rows in statement
            trailerRows = Integer.parseInt(props.getProperty("trailerRows", String.valueOf(trailerRows))); // Trailer rows count
            firstColumn = Integer.parseInt(props.getProperty("firstColumn", String.valueOf(firstColumn))); // First column number
            firstColumnName = props.getProperty("firstColumnName", firstColumnName); // Name is used for format check
            incomingRestName = props.getProperty("incomingRestName", incomingRestName); // Name is used for incoming rest
            outgoingRestName = props.getProperty("outgoingRestName", outgoingRestName); // Name is used for outgoing rest
            opNumColumnName = props.getProperty("opNumColumnName", opNumColumnName); // Operation number column name (for columns positioning)
            accountRow = Integer.parseInt(props.getProperty("accountRow", String.valueOf(accountRow))); // Row with account number

            headerValueColumn = Integer.parseInt(props.getProperty("headerValueColumn", String.valueOf(headerValueColumn))); // Column with header names
            headerNameColumn = Integer.parseInt(props.getProperty("headerNameColumn", String.valueOf(headerNameColumn))); // Column with header values

            trailerValueColumn = Integer.parseInt(props.getProperty("trailerValueColumn", String.valueOf(trailerValueColumn))); // Column with trailer names
            trailerNameColumn = Integer.parseInt(props.getProperty("trailerNameColumn", String.valueOf(trailerNameColumn))); // Column with trailer values

            dtTurnoverRowDistance = Integer.parseInt(props.getProperty("dtTurnoverRowDistance", String.valueOf(dtTurnoverRowDistance))); // Distance from last row for row with debit turnover
            crTurnoverRowDistance = Integer.parseInt(props.getProperty("crTurnoverRowDistance", String.valueOf(crTurnoverRowDistance)));
            inRestRow = Integer.parseInt(props.getProperty("inRestRow", String.valueOf(inRestRow))); // Distance from last row for row with outgoing rest
            outRestRowDistance = Integer.parseInt(props.getProperty("outRestRowDistance", String.valueOf(outRestRowDistance))); // Distance from last row for row with outgoing rest

            opDateColumn = Integer.parseInt(props.getProperty("opDateColumn", String.valueOf(opDateColumn))); // Operation date column
            opValueColumn = Integer.parseInt(props.getProperty("opValueColumn", String.valueOf(opValueColumn))); // Operation value column
            opNumColumn = Integer.parseInt(props.getProperty("opNumColumn", String.valueOf(opNumColumn))); // Operation number column
            ctrPartAccountColumn = Integer.parseInt(props.getProperty("ctrPartAccountColumn", String.valueOf(ctrPartAccountColumn))); // Counterparty account column
            ctrPartNameColumn = Integer.parseInt(props.getProperty("ctrPartNameColumn", String.valueOf(ctrPartNameColumn))); // Counterparty name column
            dtAmountColumn = Integer.parseInt(props.getProperty("dtAmountColumn", String.valueOf(dtAmountColumn))); // Debit amount column
            crAmountColumn = Integer.parseInt(props.getProperty("crAmountColumn", String.valueOf(crAmountColumn))); // Credit amount column
            purposeColumn = Integer.parseInt(props.getProperty("purposeColumn", String.valueOf(purposeColumn))); // Purpose column

            dateFormat = props.getProperty("dateFormat", dateFormat);
            formatterDt = DateTimeFormatter.ofPattern(dateFormat);
        }
        catch (Exception e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E100", "E100. Error reading parse parameters from file: {0}: {1}", iniFileName, e.getMessage());
        }
    }

    /**
     * This method checks if statement format correct
     * Can change lastHeaderRow and column indexes values
     * @return corrext format or not
     */
    @Override
    boolean check() {
        boolean isValid = true;
        init();
        try {
            int maxRow = sheet == null ? nSheet.getLastRowNum() : sheet.getLastRowNum();
            int maxColumn = sheet == null ? nSheet.getRow(maxRow).getPhysicalNumberOfCells() : sheet.getRow(maxRow).getPhysicalNumberOfCells();

            if (maxRow >= lastHeaderRow) {
                String firstRow = "?";
                while (lastHeaderRow <= maxRow) {
                   firstRow = Util.cleanStr(getCellString(lastHeaderRow, firstColumn));
                   if (firstColumnName.equalsIgnoreCase(firstRow)) {
                       break;
                   }
                   else {
                       lastHeaderRow++; // We slide down searching header line
                   }
                }
                if (lastHeaderRow >= maxRow) {
                    throw new StatementFormatError(Util.resource("fmt.unknown_row", "Unknown header row: {0} != {1}", firstRow, firstColumnName));
                }
                // We found header line, but we need to check column indexes - may be inserted fu**ing hidden columns
                while (opNumColumn <= maxColumn) {
                    String opNumColumnHeader = Util.cleanStr(getCellString(lastHeaderRow, opNumColumn).replaceAll(" ", ""));
                    if (opNumColumnName.equalsIgnoreCase(opNumColumnHeader)) {
                        break;
                    }
                    else {
                        headerValueColumn++;
                        trailerValueColumn++;

                        opNumColumn++;
                        ctrPartAccountColumn++;
                        ctrPartNameColumn++;
                        dtAmountColumn++;
                        crAmountColumn++;
                        purposeColumn++;
                    }
                }
            }
            else { // We didn't find even all header rows that we expected
                throw new StatementFormatError(Util.resource("fmt.too_few_rows","Too few rows: {0}", maxRow));
            }
        }
        catch (StatementFormatError e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E101", "E101. Statement format error: {0}", e.getMessage());
            isValid = false;
        }
        return isValid;
    }

    /**
     * This method parses statement
     * @param errHandleStrategy - error handle strategy (ALL, FORMAT, NONE)
     * @return result of parsing
     */
    @Override
    boolean parse(ErrHandleStrategy errHandleStrategy) {

        boolean result = true;
        String acctNumber;

        long dtCalcTurnover = 0;
        long crCalcTurnover = 0;

        long dtStmtTurnover = 0;
        long crStmtTurnover = 0;

        long inRest = 0;
        long outRest = 0;

        int line = 0;

        try {

            int maxRow = sheet == null ? nSheet.getLastRowNum() : sheet.getLastRowNum();

            acctNumber = getCellString(accountRow, headerValueColumn);
            Main.logr.log(System.Logger.Level.INFO, "info.our_account", "Statement for account: {0}", acctNumber);

            String inRestName = Util.cleanStr(getCellString(inRestRow, headerNameColumn).replaceAll(" ", ""));
            if (incomingRestName.equalsIgnoreCase(inRestName)) { // Incoming rest may be missed if it is = 0
                String inRestStr = getCellNumber(inRestRow, headerValueColumn);
                try {
                    inRest = Util.str2long(inRestStr);
                } catch (NumberFormatException e) {
                    String error = Util.resource("error.E112", "E112. Incoming rest format error: {0}", inRestStr);
                    if (errHandleStrategy != ErrHandleStrategy.NONE) {
                        throw new StatementFormatError(error);
                    } else {
                        Main.logr.log(System.Logger.Level.ERROR, error);
                    }
                }
            }

            String outRestName = Util.cleanStr(getCellString(maxRow - outRestRowDistance, trailerNameColumn).replaceAll(" ",""));
            if (outgoingRestName.equalsIgnoreCase(outRestName)) { // Outgoing rest may be missed if it is = 0
                String outRestStr = getCellNumber(maxRow - outRestRowDistance, trailerValueColumn);
                try {
                    outRest = Util.str2long(outRestStr);
                } catch (NumberFormatException e) {
                    String error = Util.resource("error.E113", "E113. Outgoing rest format error: {0}", outRestStr);
                    if (errHandleStrategy != ErrHandleStrategy.NONE) {
                        throw new StatementFormatError(error);
                    } else {
                        Main.logr.log(System.Logger.Level.ERROR, error);
                    }
                }
            }

            String turnoverStr = "?";
            try {
                turnoverStr = getCellNumber(maxRow - dtTurnoverRowDistance, headerValueColumn);
                dtStmtTurnover = Util.str2long(turnoverStr);
                turnoverStr = getCellNumber(maxRow - crTurnoverRowDistance, headerValueColumn);
                crStmtTurnover = Util.str2long(turnoverStr);
            }
            catch (NumberFormatException e) {
                String error = Util.resource("error.E111", "E111. Debit or credit turnover format error: {0}", turnoverStr);
                if (errHandleStrategy != ErrHandleStrategy.NONE) {
                    throw new StatementFormatError(error);
                }
                else {
                    Main.logr.log(System.Logger.Level.ERROR,error);
                }
            }

            for (int rowNum = lastHeaderRow + 1; rowNum < maxRow - trailerRows; rowNum++) {

                line++;

                try {
                    Operation op = new Operation();

                    op.id = Integer.toString(line);

                    op.opNum = getCellString(rowNum, opNumColumn);

                    String opDateStr = "?";
                    try {
                        opDateStr = getCellDate(rowNum, opDateColumn);
                        op.opDate = LocalDate.parse(opDateStr, formatterDt);
                        opDateStr = getCellDate(rowNum, opValueColumn);
                        op.opValue = LocalDate.parse(opDateStr, formatterDt);
                    } catch (DateTimeException e) {
                        String error = Util.resource("error.E102", "E102. Date format error: {0}, line {1}", opDateStr, rowNum);
                        result = false;
                        if (errHandleStrategy != ErrHandleStrategy.NONE) {
                            throw new StatementFormatError(error);
                        }
                        else {
                            Main.logr.log(System.Logger.Level.ERROR, error); // We will continue
                        }
                    }

                    op.ctrPartAccount = getCellString(rowNum, ctrPartAccountColumn);
                    op.ctrPartName = Util.cleanStr(getCellString(rowNum, ctrPartNameColumn));

                    String dtAmountStr = getCellNumber(rowNum, dtAmountColumn);
                    try {
                        op.dtAmount = Util.str2long(dtAmountStr);
                    } catch (NumberFormatException e) {
                        String error = Util.resource("error.E103", "E103. Debit amount format error: {0}, line {1}", dtAmountStr, rowNum);
                        if (errHandleStrategy != ErrHandleStrategy.NONE) {
                            throw new StatementFormatError(error);
                        }
                        else {
                            Main.logr.log(System.Logger.Level.ERROR,error);
                        }
                    }

                    String crAmountStr = getCellNumber(rowNum, crAmountColumn);
                    try {
                        op.crAmount = Util.str2long(crAmountStr);
                    } catch (NumberFormatException e) {
                        String error = Util.resource("error.E104", "E104. Credit amount format error: {0}, line {1}", crAmountStr, rowNum);
                        if (errHandleStrategy != ErrHandleStrategy.NONE) {
                            throw new StatementFormatError(error);
                        }
                        else {
                            Main.logr.log(System.Logger.Level.ERROR,error);
                        }
                    }

                    op.purpose = Util.cleanStr(getCellString(rowNum, purposeColumn));

                    // Constant part of CSV statement
                    op.ourAccount = Util.cleanStr(acctNumber);
                    op.dtTurnover = dtStmtTurnover;
                    op.crTurnover = crStmtTurnover;
                    op.incomingRest = inRest;
                    op.outgoingRest = outRest;

                    dtCalcTurnover += op.dtAmount;
                    crCalcTurnover += op.crAmount;

                    String str = op.toString();
                    try {
                        writer.write(str);
                        writer.write(Util.lSep);
                    }
                    catch (IOException e) {
                        String error = Util.resource("error.E105", "E105. CSV file output error: {0}", e.getMessage());
                        Main.logr.log(System.Logger.Level.ERROR, error);
                        if (errHandleStrategy != ErrHandleStrategy.NONE) {
                            result = false;
                            break;
                        }
                    }
                }
                catch (Exception e) {
                    String error = Util.resource("error.E106", "E106. Line {0} parsing error: {1}",  rowNum, e.getMessage());
                    Main.logr.log(System.Logger.Level.ERROR, error);
                    if (errHandleStrategy != ErrHandleStrategy.NONE) {
                        result = false;
                        break;
                    }
                }
            }

            Main.logr.log(System.Logger.Level.INFO,"info.operations_total","Done. {0} operation(s) parsed.", line);

            if (dtStmtTurnover != dtCalcTurnover) {
                Main.logr.log(System.Logger.Level.ERROR, "error.E107",
                        "E107. Debit turnover mismatch. Specified: {0}, calculated: {1}",
                        Util.long2str(dtStmtTurnover), Util.long2str(dtCalcTurnover));
                if (errHandleStrategy != ErrHandleStrategy.NONE) {
                    result = false;
                }
            }
            else {
                Main.logr.log(System.Logger.Level.INFO, "info.debit_turnover", "Debit turnover: {0}", Util.long2str(dtStmtTurnover));
            }

            if (crStmtTurnover != crCalcTurnover) {
                Main.logr.log(System.Logger.Level.ERROR, "error.E108",
                        "E108. Credit turnover mismatch. Specified: {0}, calculated: {1}",
                         Util.long2str(crStmtTurnover), Util.long2str(crCalcTurnover));
                if (errHandleStrategy != ErrHandleStrategy.NONE) {
                    result = false;
                }
            } else {
                Main.logr.log(System.Logger.Level.INFO,"info.credit_turnover", "Credit turnover: {0}", Util.long2str(crStmtTurnover));
            }


            Main.logr.log(System.Logger.Level.INFO,"info.incoming_rest", "Incoming rest: {0}", Util.long2str(inRest));
            Main.logr.log(System.Logger.Level.INFO,"info.outgoing_rest", "Outgoing rest: {0}", Util.long2str(outRest));
        }
        catch (StatementFormatError e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E109", "E109. Statement parse error: {0}", e.getMessage());
            if (errHandleStrategy != ErrHandleStrategy.NONE) {
                result = false;
            }
        }
        catch (Exception e) {
            Main.logr.log(System.Logger.Level.ERROR,"error.E110", "E110. General error: {0}", e.getMessage());
            if (errHandleStrategy != ErrHandleStrategy.NONE) {
                result = false;
            }
        }
        return result;
    }
}
