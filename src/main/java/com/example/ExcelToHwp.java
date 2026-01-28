package com.example;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.control.table.DivideAtPageBoundary;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelToHwp extends JFrame {

    private JTextField txtExcelPath;
    private JTextField txtSheetName;
    private JTextField txtStartRow;
    private JTextField txtStartCol;
    private JTextField txtEndRow;
    private JTextField txtEndCol;
    private JTextArea logArea;

    public ExcelToHwp() {
        setTitle("Excel to HWP Converter (Single Paragraph Safe)");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(500, 450);
        setLocationRelativeTo(null);
        setLayout(new BorderLayout());

        JPanel panelInput = new JPanel(new GridLayout(7, 2, 5, 5));
        panelInput.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 1. Excel File
        panelInput.add(new JLabel("Excel File:"));
        JPanel filePanel = new JPanel(new BorderLayout(5, 0));
        txtExcelPath = new JTextField();
        txtExcelPath.setEditable(false);
        JButton btnBrowse = new JButton("Browse");
        btnBrowse.addActionListener(e -> browseFile());
        filePanel.add(txtExcelPath, BorderLayout.CENTER);
        filePanel.add(btnBrowse, BorderLayout.EAST);
        panelInput.add(filePanel);

        // 2. Sheet Name
        panelInput.add(new JLabel("Sheet Name (Empty=1st):"));
        txtSheetName = new JTextField();
        panelInput.add(txtSheetName);

        // 3. Range
        panelInput.add(new JLabel("Start Row (1-based):"));
        txtStartRow = new JTextField("1");
        panelInput.add(txtStartRow);

        panelInput.add(new JLabel("Start Col (A, B... or 1):"));
        txtStartCol = new JTextField("A");
        panelInput.add(txtStartCol);

        panelInput.add(new JLabel("End Row (Empty=Auto):"));
        txtEndRow = new JTextField("");
        panelInput.add(txtEndRow);

        panelInput.add(new JLabel("End Col (Empty=Auto):"));
        txtEndCol = new JTextField("");
        panelInput.add(txtEndCol);

        // Convert Button
        JButton btnConvert = new JButton("Convert to HWP");
        btnConvert.addActionListener(e -> startConversion());
        panelInput.add(new JLabel(""));
        panelInput.add(btnConvert);

        add(panelInput, BorderLayout.NORTH);

        logArea = new JTextArea();
        logArea.setEditable(false);
        add(new JScrollPane(logArea), BorderLayout.CENTER);
    }

    private void log(String msg) {
        logArea.append(msg + "\n");
        logArea.setCaretPosition(logArea.getDocument().getLength());
    }

    private void browseFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        chooser.setCurrentDirectory(new File("."));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            txtExcelPath.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private int colStrToInt(String col) {
        if (col == null || col.trim().isEmpty())
            return -1;
        if (col.matches("\\d+"))
            return Integer.parseInt(col) - 1;

        int result = 0;
        col = col.toUpperCase().trim();
        for (int i = 0; i < col.length(); i++) {
            result *= 26;
            result += col.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    private void startConversion() {
        new Thread(() -> {
            try {
                String excelPath = txtExcelPath.getText();
                if (excelPath.isEmpty()) {
                    log("Error: Select Excel file first.");
                    return;
                }

                log("Reading Excel...");
                List<List<String>> data = readExcel(excelPath);

                if (data == null || data.isEmpty()) {
                    log("No data found.");
                    return;
                }

                log("Read " + data.size() + " rows. Creating HWP...");

                String hwpPath = excelPath.substring(0, excelPath.lastIndexOf('.')) + "_Result.hwp";

                createHwpSafe(hwpPath, data);

                log("Success! Saved to: " + hwpPath);
                JOptionPane.showMessageDialog(this, "Success!\nSaved to: " + hwpPath);

            } catch (Exception e) {
                log("Error: " + e.getMessage());
                e.printStackTrace();
            }
        }).start();
    }

    private List<List<String>> readExcel(String path) throws Exception {
        FileInputStream fis = new FileInputStream(new File(path));
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet sheet;
        String sName = txtSheetName.getText().trim();
        if (sName.isEmpty()) {
            sheet = workbook.getSheetAt(0);
        } else {
            sheet = workbook.getSheet(sName);
            if (sheet == null) {
                throw new Exception("Sheet not found: " + sName);
            }
        }

        int startR = -1;
        try {
            startR = Integer.parseInt(txtStartRow.getText().trim()) - 1;
        } catch (NumberFormatException e) {
            startR = 0;
        }

        int startC = colStrToInt(txtStartCol.getText());
        if (startC < 0)
            startC = 0;

        int endR = -1;
        if (!txtEndRow.getText().trim().isEmpty())
            endR = Integer.parseInt(txtEndRow.getText().trim()) - 1;
        else
            endR = sheet.getLastRowNum();

        int endC = -1;
        if (!txtEndCol.getText().trim().isEmpty())
            endC = colStrToInt(txtEndCol.getText());

        List<List<String>> data = new ArrayList<>();

        for (int r = startR; r <= endR; r++) {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
            List<String> rowData = new ArrayList<>();
            int localEndC = (endC == -1) ? (row == null ? -1 : row.getLastCellNum() - 1) : endC;
            if (row != null) {
                if (localEndC == -1)
                    localEndC = row.getLastCellNum() - 1;
                for (int c = startC; c <= localEndC; c++) {
                    String val = getMergedRegionValue(sheet, r, c);
                    if (val == null) {
                        org.apache.poi.ss.usermodel.Cell cell = row.getCell(c,
                                org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        val = getCellValueAsString(cell);
                    }
                    rowData.add(val);
                }
            } else {
                for (int c = startC; c <= localEndC; c++)
                    rowData.add("");
            }
            data.add(rowData);
        }

        workbook.close();
        fis.close();
        return data;
    }

    private String getMergedRegionValue(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(row, col)) {
                int firstRow = region.getFirstRow();
                int firstCol = region.getFirstColumn();
                org.apache.poi.ss.usermodel.Row fRow = sheet.getRow(firstRow);
                if (fRow == null)
                    return "";
                org.apache.poi.ss.usermodel.Cell fCell = fRow.getCell(firstCol);
                return getCellValueAsString(fCell);
            }
        }
        return null;
    }

    private String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null)
            return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell))
                    return cell.getDateCellValue().toString();
                double startVal = cell.getNumericCellValue();
                if (startVal == (long) startVal)
                    return String.format("%d", (long) startVal);
                return String.valueOf(startVal);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default:
                return "";
        }
    }

    private void createHwpSafe(String path, List<List<String>> data) throws Exception {
        HWPFile hwpFile = BlankFileMaker.make();
        if (hwpFile == null)
            throw new Exception("Failed to make blank hwp file");

        // Use the first paragraph of the first section
        Section section = hwpFile.getBodyText().getSectionList().get(0);
        Paragraph paragraph = section.getParagraph(0);

        if (data == null || data.isEmpty())
            return;

        int rowCount = data.size();
        int colCount = data.get(0).size();

        // 1. Create Table Control
        ControlTable table = (ControlTable) paragraph.addNewControl(ControlType.Table);

        // 2. Initialize Table Properties (Rows, Cols)
        // Note: hwplib requires careful initialization of ListHeaders and Cells
        initTable(table, rowCount, colCount);

        // 3. Fill Data
        for (int r = 0; r < rowCount; r++) {
            Row row = table.getRowList().get(r);
            List<String> rowData = data.get(r);
            for (int c = 0; c < colCount; c++) {
                if (c >= row.getCellList().size())
                    break;

                Cell cell = row.getCellList().get(c);
                Paragraph cellParagraph = cell.getParagraphList().getParagraph(0);

                // Set Text
                String cellValue = (c < rowData.size()) ? rowData.get(c) : "";
                setParagraphText(cellParagraph, cellValue);
            }
        }

        HWPWriter.toFile(hwpFile, path);
    }

    private void initTable(ControlTable table, int rowCount, int colCount) {
        // Basic Table Property Setup
        // Simplified: Default properties should suffice for now.

        // Row Count / Col Count
        table.getTable().setRowCount(rowCount);
        table.getTable().setColumnCount(colCount);

        // Add Rows and Cells
        table.getRowList().clear();

        for (int i = 0; i < rowCount; i++) {
            Row row = table.addNewRow();
            for (int j = 0; j < colCount; j++) {
                // Add Cell
                Cell cell = row.addNewCell();
                // Each Cell must have at least one Paragraph
                // Correct API: Cell -> ParagraphList -> addNewParagraph()
                Paragraph p = cell.getParagraphList().addNewParagraph();

                // Initialize Paragraph Header
                p.getHeader().setParaShapeId(1);
                p.getHeader().setStyleId((short) 1);
                p.createText();

                // Set Cell dimensions (Optional but good for visual)
                cell.getListHeader().setWidth(2000); // approx width
                cell.getListHeader().setHeight(1000); // approx height
            }
        }
    }

    private void setParagraphText(Paragraph p, String text) {
        if (text == null)
            text = "";

        // Ensure text object exists
        if (p.getText() == null) {
            p.createText();
        }

        // We do not clear text as 'clear()' is not easily available on ParaText.
        // Assuming new paragraphs are empty.

        // HWP Text needs to be added via Char objects.
        try {
            p.getText().addString(text);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            new ExcelToHwp().setVisible(true);
        });
    }
}
