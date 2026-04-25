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
    private String firstColumnName = "Датасоздания"; // Name is used for format check
    private int accountRow = 4; // Row with account number
    private int accountColumn = 6; // Column with account number

    private int turnoverColumn = 6; // Column with turnover
    private int dtTurnoverRowDistance = 4; // Distance from last row for row with debit turnover
    private int crTurnoverRowDistance = 3;
    private int restRowDistance = 2; // Distance from last row for row with outgoing rest

    private int opNumColumn = 7; // Operation number column
    private int opDateColumn = 3; // Operation date column
    private int ctrPartAccountColumn = 9; // Counterparty account column
    private int ctrPartNameColumn = 11; // Counterparty name column
    private int dtAmountColumn = 13; // Debit amount column
    private int crAmountColumn = 14; // Credit amount column
    private int purposeColumn = 16; // Purpose column

    DateTimeFormatter formatterDt = DateTimeFormatter.ofPattern(dateFormat);

    void init() { // Override markup parameters
        Properties props = new Properties();
        String iniFileName = "btb.ini";
        try (Reader reader = new InputStreamReader(new FileInputStream(iniFileName), StandardCharsets.UTF_8)) {
            props.load(reader);

            lastHeaderRow = Integer.parseInt(props.getProperty("lastHeaderRow", String.valueOf(lastHeaderRow))); // Last table header row, minimum rows in statement
            trailerRows = Integer.parseInt(props.getProperty("trailerRows", String.valueOf(trailerRows))); // Trailer rows count
            firstColumn = Integer.parseInt(props.getProperty("firstColumn", String.valueOf(firstColumn))); // First column number
            firstColumnName = props.getProperty("firstColumnName", firstColumnName); // Name is used for format check
            accountRow = Integer.parseInt(props.getProperty("accountRow", String.valueOf(accountRow))); // Row with account number
            accountColumn = Integer.parseInt(props.getProperty("accountColumn", String.valueOf(accountColumn))); // Column with account number

            turnoverColumn = Integer.parseInt(props.getProperty("turnoverColumn", String.valueOf(turnoverColumn))); // Column with turnover
            dtTurnoverRowDistance = Integer.parseInt(props.getProperty("dtTurnoverRowDistance", String.valueOf(dtTurnoverRowDistance))); // Distance from last row for row with debit turnover
            crTurnoverRowDistance = Integer.parseInt(props.getProperty("crTurnoverRowDistance", String.valueOf(crTurnoverRowDistance)));
            restRowDistance = Integer.parseInt(props.getProperty("restRowDistance", String.valueOf(restRowDistance))); // Distance from last row for row with outgoing rest

            opDateColumn = Integer.parseInt(props.getProperty("opDateColumn", String.valueOf(opDateColumn))); // Operation date column
            ctrPartAccountColumn = Integer.parseInt(props.getProperty("ctrPartAccountColumn", String.valueOf(ctrPartAccountColumn))); // Counterparty account column
            ctrPartNameColumn = Integer.parseInt(props.getProperty("ctrPartNameColumn", String.valueOf(ctrPartNameColumn))); // Counterparty name column
            dtAmountColumn = Integer.parseInt(props.getProperty("dtAmountColumn", String.valueOf(dtAmountColumn))); // Debit amount column
            crAmountColumn = Integer.parseInt(props.getProperty("crAmountColumn", String.valueOf(crAmountColumn))); // Credit amount column
            purposeColumn = Integer.parseInt(props.getProperty("purposeColumn", String.valueOf(purposeColumn))); // Purpose column
            opNumColumn = Integer.parseInt(props.getProperty("opNumColumn", String.valueOf(opNumColumn))); // Operation number column

            dateFormat = props.getProperty("dateFormat", dateFormat);
            formatterDt = DateTimeFormatter.ofPattern(dateFormat);
        }
        catch (Exception e) {
            System.out.println("E013. Error reading parse parameters from file: " + iniFileName + " : " + e.getMessage());
        }
    }

    boolean check() {
        boolean isValid = true;
        init();
        try {
            int maxRow = sheet == null ? nSheet.getLastRowNum() : sheet.getLastRowNum();

            if (maxRow >= lastHeaderRow) {
                String firstRow = sheet == null ?
                        nSheet.getRow(lastHeaderRow).getCell(firstColumn).getStringCellValue() :
                        sheet.getRow(lastHeaderRow).getCell(firstColumn).getStringCellValue();
                if (!firstColumnName.equals(firstRow.replace("\n","").replace("\r", ""))) {
                    throw new StatementFormatError("Unknown header row: " + firstRow + " <> " + firstColumnName);
                }
            }
            else {
                throw new StatementFormatError("Too few rows: " + maxRow);
            }
        }
        catch (StatementFormatError e) {
            System.out.println("E101. Statement format error. " + e.getMessage());
            isValid = false;
        }
        return isValid;
    }

    @Override
    int parse() {

        String acctNumber;

        long dtCalcTurnover = 0;
        long  crCalcTurnover = 0;

        int line = 0;

        try {

            int maxRow = sheet == null? nSheet.getLastRowNum() : sheet.getLastRowNum();

            acctNumber = sheet == null?
                    nSheet.getRow(accountRow).getCell(accountColumn).getStringCellValue() :
                    sheet.getRow(accountRow).getCell(accountColumn).getStringCellValue();
            System.out.println("Statement for account: " + acctNumber);

            for (int rowNum = lastHeaderRow + 1; rowNum < maxRow - trailerRows; rowNum++) {

                line++;

                try {
                    Operation op = new Operation();

                    op.id = Integer.toString(line);

                    op.opNum = getCellString(rowNum, opNumColumn);
                    String opDateStr = getCellDate(rowNum, opDateColumn);
                    try {
                        op.opDate = LocalDate.parse(opDateStr, formatterDt);
                    } catch (DateTimeException e) {
                        System.out.println("E105. Date format error: " + opDateStr + ", line:" + line);
                    }

                    op.ctrPartAccount = getCellString(rowNum, ctrPartAccountColumn);
                    op.ctrPartName = Util.cleanStr(getCellString(rowNum, ctrPartNameColumn));

                    String dtAmountStr = getCellNumber(rowNum, dtAmountColumn);
                    try {
                        op.dtAmount = Util.str2long(dtAmountStr);
                    } catch (NumberFormatException e) {
                        System.out.println("E106. Debit amount format error: " + dtAmountStr + ", line:" + line);
                    }

                    String crAmountStr = getCellNumber(rowNum, crAmountColumn);
                    try {
                        op.crAmount = Util.str2long(crAmountStr);
                    } catch (NumberFormatException e) {
                        System.out.println("E107. Credit amount format error: " + crAmountStr + ", line:" + line);
                    }

                    op.purpose = Util.cleanStr(getCellString(rowNum, purposeColumn));

                    dtCalcTurnover += op.dtAmount;
                    crCalcTurnover += op.crAmount;

                    String str = op.getCSVString();
                    try {
                        writer.write(str);
                        writer.write(Util.lSep);
                    }
                    catch (IOException e) {
                        System.out.println("E108. CSV file output error: " + e.getMessage());
                    }

                }
                catch (Exception e) {
                    System.out.println("E110. Line " + line + " parsing error: " + e.getMessage());
                }
            }

            String dtTurnoverStr = getCellNumber(maxRow - dtTurnoverRowDistance, turnoverColumn);
            String crTurnoverStr = getCellNumber(maxRow - crTurnoverRowDistance, turnoverColumn);
            String outRestStr = getCellNumber(maxRow - restRowDistance, turnoverColumn);

            long dtStmtTurnover = Util.str2long(dtTurnoverStr);
            if (dtStmtTurnover != dtCalcTurnover) {
                System.out.println("E102. Debit turnover mismatch. Specified: " + Util.long2str(dtStmtTurnover) +
                                                               ", calculated: " + Util.long2str(dtCalcTurnover));
            }
            else {
                System.out.println("Debit turnover: " + Util.long2str(dtStmtTurnover));
            }

            long crStmtTurnover = Util.str2long(crTurnoverStr);
            if (crStmtTurnover != crCalcTurnover) {
                System.out.println("E103. Credit turnover mismatch. Specified: " + Util.long2str(crStmtTurnover) +
                                                                ", calculated: " + Util.long2str(crCalcTurnover));
            } else {
                System.out.println("Credit turnover: " + Util.long2str(crStmtTurnover));
            }
            long outRest = Util.str2long(outRestStr);
            System.out.println("Outgoing rest: " + Util.long2str(outRest));
        }
        catch (StatementFormatError e) {
            System.out.println("E101. Statement format error. " + e.getMessage());
        }
        catch (Exception e) {
            System.out.println("E100. Error: " + e.getMessage());
        }
        return line;
    }
}
