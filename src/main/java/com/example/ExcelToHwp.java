package com.example;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.ListHeaderForCell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.control.table.Table;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.control.table.*;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.ParaText;
import kr.dogfoot.hwplib.object.docinfo.BorderFill;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BackSlashDiagonalShape;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BorderThickness;
import kr.dogfoot.hwplib.object.docinfo.borderfill.BorderType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.SlashDiagonalShape;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PatternFill;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PatternType;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        setTitle("Excel to HWP Converter (Refactored)");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(600, 500);
        setLocationRelativeTo(null);
        setLayout(new BorderLayout());

        JPanel panelInput = new JPanel(new GridLayout(7, 2, 5, 5));
        panelInput.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 1. Excel File
        panelInput.add(new JLabel("엑셀 파일 경로:"));
        JPanel filePanel = new JPanel(new BorderLayout(5, 0));
        txtExcelPath = new JTextField();
        txtExcelPath.setEditable(false);
        JButton btnBrowse = new JButton("찾아보기");
        btnBrowse.addActionListener(e -> browseFile());
        filePanel.add(txtExcelPath, BorderLayout.CENTER);
        filePanel.add(btnBrowse, BorderLayout.EAST);
        panelInput.add(filePanel);

        // 2. Sheet Name
        panelInput.add(new JLabel("시트 이름 (비워두면 첫번째 시트):"));
        txtSheetName = new JTextField();
        panelInput.add(txtSheetName);

        // 3. Range
        panelInput.add(new JLabel("시작 행 (1부터 시작):"));
        txtStartRow = new JTextField("1");
        panelInput.add(txtStartRow);

        panelInput.add(new JLabel("시작 열 (A, B... 또는 1):"));
        txtStartCol = new JTextField("A");
        panelInput.add(txtStartCol);

        panelInput.add(new JLabel("종료 행 (비워두면 자동):"));
        txtEndRow = new JTextField("");
        panelInput.add(txtEndRow);

        panelInput.add(new JLabel("종료 열 (비워두면 자동):"));
        txtEndCol = new JTextField("");
        panelInput.add(txtEndCol);

        // Convert Button
        JButton btnConvert = new JButton("HWP로 변환");
        btnConvert.setBackground(new java.awt.Color(70, 130, 180));
        btnConvert.setForeground(java.awt.Color.WHITE);
        btnConvert.setFont(new java.awt.Font("Malgun Gothic", java.awt.Font.BOLD, 14));
        btnConvert.addActionListener(e -> startConversion());
        panelInput.add(new JLabel("")); // Placeholder
        panelInput.add(btnConvert);

        add(panelInput, BorderLayout.NORTH);

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 12));
        add(new JScrollPane(logArea), BorderLayout.CENTER);
    }

    private void log(String msg) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(msg + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private void browseFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        chooser.setCurrentDirectory(new File("."));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            txtExcelPath.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void startConversion() {
        new Thread(() -> {
            try {
                String excelPath = txtExcelPath.getText();
                if (excelPath.isEmpty()) {
                    log("오류: 엑셀 파일을 선택해주세요.");
                    JOptionPane.showMessageDialog(this, "엑셀 파일을 선택해주세요.", "오류", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                log("엑셀 파일 읽는 중...");
                List<List<String>> data = readExcel(excelPath);

                if (data == null || data.isEmpty()) {
                    log("데이터가 없습니다.");
                    return;
                }

                log("총 " + data.size() + "행의 데이터를 읽었습니다. HWP 파일 생성 중...");

                String hwpPath = excelPath.substring(0, excelPath.lastIndexOf('.')) + "_Result.hwp";
                createHwpFile(hwpPath, data);

                log("완료! 저장된 경로: " + hwpPath);
                JOptionPane.showMessageDialog(this, "변환 완료!\n저장 경로: " + hwpPath, "성공", JOptionPane.INFORMATION_MESSAGE);

            } catch (Exception e) {
                log("오류 발생: " + e.toString());
                for (StackTraceElement ste : e.getStackTrace()) {
                    log("\tat " + ste.toString());
                }
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "오류 발생: " + e.getMessage(), "오류", JOptionPane.ERROR_MESSAGE);
            }
        }).start();
    }

    // --- Excel Processing ---

    private List<List<String>> readExcel(String path) throws Exception {
        try (FileInputStream fis = new FileInputStream(new File(path));
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet;
            String sName = txtSheetName.getText().trim();
            if (sName.isEmpty()) {
                sheet = workbook.getSheetAt(0);
            } else {
                sheet = workbook.getSheet(sName);
                if (sheet == null) {
                    throw new Exception("시트를 찾을 수 없습니다: " + sName);
                }
            }

            int startR = parseRowIndex(txtStartRow.getText(), 0);
            int startC = parseColIndex(txtStartCol.getText(), 0);
            int endR = parseRowIndex(txtEndRow.getText(), sheet.getLastRowNum());
            int endC = parseColIndex(txtEndCol.getText(), -1); // -1 means auto detection per row

            List<List<String>> data = new ArrayList<>();

            for (int r = startR; r <= endR; r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                List<String> rowData = new ArrayList<>();

                int localEndC = endC;
                if (localEndC == -1) {
                    localEndC = (row == null) ? startC : Math.max(startC, row.getLastCellNum() - 1);
                }

                for (int c = startC; c <= localEndC; c++) {
                    String val = getMergedRegionValue(sheet, r, c);
                    if (val == null) {
                        if (row != null) {
                            org.apache.poi.ss.usermodel.Cell cell = row.getCell(c,
                                    org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            val = getCellValueAsString(cell);
                        } else {
                            val = "";
                        }
                    }
                    rowData.add(val);
                }
                data.add(rowData);
                log("Row " + (r + 1) + ": " + rowData.toString());
            }
            return data;
        }
    }

    private int parseRowIndex(String text, int defaultValue) {
        try {
            String t = text.trim();
            if (t.isEmpty())
                return defaultValue;
            return Integer.parseInt(t) - 1;
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    private int parseColIndex(String text, int defaultValue) {
        String t = text.trim();
        if (t.isEmpty())
            return defaultValue;
        if (t.matches("\\d+"))
            return Integer.parseInt(t) - 1;
        return colStrToInt(t);
    }

    private int colStrToInt(String col) {
        int result = 0;
        col = col.toUpperCase().trim();
        for (int i = 0; i < col.length(); i++) {
            result *= 26;
            result += col.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    private String getMergedRegionValue(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(row, col)) {
                org.apache.poi.ss.usermodel.Row fRow = sheet.getRow(region.getFirstRow());
                if (fRow == null)
                    return "";
                org.apache.poi.ss.usermodel.Cell fCell = fRow.getCell(region.getFirstColumn());
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
                double val = cell.getNumericCellValue();
                if (val == (long) val)
                    return String.format("%d", (long) val);
                return String.valueOf(val);
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

    // --- HWP Processing ---

    private void createHwpFile(String path, List<List<String>> data) throws Exception {
        HWPFile hwpFile = BlankFileMaker.make();
        if (hwpFile == null)
            throw new Exception("HWP 파일 생성 실패");

        Section section = hwpFile.getBodyText().getSectionList().get(0);
        Paragraph paragraph = section.getParagraph(0);

        // 테이블 생성
        int rowCount = data.size();
        int colCount = data.isEmpty() ? 0 : data.get(0).size();
        if (rowCount == 0 || colCount == 0)
            return;

        ControlTable table = (ControlTable) paragraph.addNewControl(ControlType.Table);
        paragraph.getText().addExtendCharForTable();

        // [설정] 테두리 ID 생성
        int borderFillID = getBorderFillIDForCell(hwpFile);

        // [중요] 테이블 헤더(GSO) 속성 설정
        long cellWidth = mmToHwp(30.0);
        long cellHeight = mmToHwp(10.0);
        long totalWidth = cellWidth * colCount;
        long totalHeight = cellHeight * rowCount;

        setCtrlHeaderRecord(table, totalWidth, totalHeight);
        setTableRecord(table, rowCount, colCount, borderFillID);

        // 데이터 채우기
        for (int r = 0; r < rowCount; r++) {
            Row row = table.addNewRow();
            for (int c = 0; c < colCount; c++) {
                Cell cell = row.addNewCell();

                setListHeaderForCell(cell, c, r, cellWidth, cellHeight, borderFillID);

                String text = "";
                if (data.size() > r && data.get(r).size() > c) {
                    text = data.get(r).get(c);
                }
                if (text == null || text.isEmpty())
                    text = " ";

                setParagraphForCell(cell, text);
            }
        }

        HWPWriter.toFile(hwpFile, path);
    }

    private void setCtrlHeaderRecord(ControlTable table, long width, long height) {
        CtrlHeaderGso ctrlHeader = (CtrlHeaderGso) table.getHeader();
        ctrlHeader.getProperty().setLikeWord(false);
        ctrlHeader.getProperty().setApplyLineSpace(false);
        ctrlHeader.getProperty().setVertRelTo(VertRelTo.Para);
        ctrlHeader.getProperty().setVertRelativeArrange(RelativeArrange.TopOrLeft);
        ctrlHeader.getProperty().setHorzRelTo(HorzRelTo.Para);
        ctrlHeader.getProperty().setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        ctrlHeader.getProperty().setVertRelToParaLimit(false);
        ctrlHeader.getProperty().setAllowOverlap(false);
        ctrlHeader.getProperty().setWidthCriterion(WidthCriterion.Absolute);
        ctrlHeader.getProperty().setHeightCriterion(HeightCriterion.Absolute);
        ctrlHeader.getProperty().setProtectSize(false);
        ctrlHeader.getProperty().setTextFlowMethod(TextFlowMethod.FitWithText);
        ctrlHeader.getProperty().setTextHorzArrange(TextHorzArrange.BothSides);
        ctrlHeader.getProperty().setObjectNumberSort(ObjectNumberSort.Table);
        ctrlHeader.setxOffset(mmToHwp(20.0));
        ctrlHeader.setyOffset(mmToHwp(20.0));
        ctrlHeader.setWidth(width);
        ctrlHeader.setHeight(height);
        ctrlHeader.setzOrder(0);
        ctrlHeader.setOutterMarginLeft(0);
        ctrlHeader.setOutterMarginRight(0);
        ctrlHeader.setOutterMarginTop(0);
        ctrlHeader.setOutterMarginBottom(0);
    }

    private void setTableRecord(ControlTable table, int rowCount, int colCount, int borderFillId) {
        Table tableRecord = table.getTable();
        tableRecord.getProperty().setDivideAtPageBoundary(DivideAtPageBoundary.DivideByCell);
        tableRecord.getProperty().setAutoRepeatTitleRow(false);
        tableRecord.setRowCount(rowCount);
        tableRecord.setColumnCount(colCount);
        tableRecord.setCellSpacing(0);
        tableRecord.setLeftInnerMargin(0);
        tableRecord.setRightInnerMargin(0);
        tableRecord.setTopInnerMargin(0);
        tableRecord.setBottomInnerMargin(0);
        tableRecord.setBorderFillId(borderFillId);

        tableRecord.getCellCountOfRowList().clear();
        for (int i = 0; i < rowCount; i++) {
            tableRecord.getCellCountOfRowList().add(colCount);
        }
    }

    private void setListHeaderForCell(Cell cell, int colIndex, int rowIndex, long width, long height,
            int borderFillId) {
        ListHeaderForCell lh = cell.getListHeader();
        lh.setParaCount(1);
        lh.getProperty().setTextDirection(TextDirection.Horizontal);
        lh.getProperty().setLineChange(LineChange.Normal);
        lh.getProperty().setTextVerticalAlignment(TextVerticalAlignment.Center);
        lh.getProperty().setProtectCell(false);
        lh.getProperty().setEditableAtFormMode(false);
        lh.setColIndex(colIndex);
        lh.setRowIndex(rowIndex);
        lh.setColSpan(1);
        lh.setRowSpan(1);
        lh.setWidth(width);
        lh.setHeight(height);
        lh.setLeftMargin(0);
        lh.setRightMargin(0);
        lh.setTopMargin(0);
        lh.setBottomMargin(0);
        lh.setBorderFillId(borderFillId);
        lh.setTextWidth(width);
        lh.setFieldName("");
    }

    private void setParagraphForCell(Cell cell, String text) {
        Paragraph p = cell.getParagraphList().addNewParagraph();

        // Header
        ParaHeader ph = p.getHeader();
        ph.setLastInList(true);
        ph.setParaShapeId(0); // Basic
        ph.setStyleId((short) 0); // Basic
        ph.getDivideSort().setDivideSection(false);
        ph.getDivideSort().setDivideMultiColumn(false);
        ph.getDivideSort().setDividePage(false);
        ph.getDivideSort().setDivideColumn(false);
        ph.setCharShapeCount(1);
        ph.setRangeTagCount(0);
        ph.setLineAlignCount(1);
        ph.setInstanceID(0);
        ph.setIsMergedByTrack(0);

        // Text
        p.createText();
        ParaText pt = p.getText();
        try {
            pt.addString(text);
        } catch (Exception e) {
            e.printStackTrace();
        }

        // CharShape
        p.createCharShape();
        ParaCharShape pcs = p.getCharShape();
        pcs.addParaCharShape(0, 0); // StartPos 0, ShapeId 0
    }

    private int getBorderFillIDForCell(HWPFile hwpFile) {
        BorderFill bf = hwpFile.getDocInfo().addNewBorderFill();
        bf.getProperty().set3DEffect(false);
        bf.getProperty().setShadowEffect(false);
        bf.getProperty().setSlashDiagonalShape(SlashDiagonalShape.None);
        bf.getProperty().setBackSlashDiagonalShape(BackSlashDiagonalShape.None);

        bf.getLeftBorder().setType(BorderType.Solid);
        bf.getLeftBorder().setThickness(BorderThickness.MM0_5);
        bf.getLeftBorder().getColor().setValue(0x0);

        bf.getRightBorder().setType(BorderType.Solid);
        bf.getRightBorder().setThickness(BorderThickness.MM0_5);
        bf.getRightBorder().getColor().setValue(0x0);

        bf.getTopBorder().setType(BorderType.Solid);
        bf.getTopBorder().setThickness(BorderThickness.MM0_5);
        bf.getTopBorder().getColor().setValue(0x0);

        bf.getBottomBorder().setType(BorderType.Solid);
        bf.getBottomBorder().setThickness(BorderThickness.MM0_5);
        bf.getBottomBorder().getColor().setValue(0x0);

        bf.getDiagonalBorder().setType(BorderType.None);
        bf.getDiagonalBorder().setThickness(BorderThickness.MM0_5);
        bf.getDiagonalBorder().getColor().setValue(0x0);

        bf.getFillInfo().getType().setPatternFill(true);
        bf.getFillInfo().createPatternFill();
        PatternFill pf = bf.getFillInfo().getPatternFill();
        pf.setPatternType(PatternType.None);
        pf.getBackColor().setValue(-1);
        pf.getPatternColor().setValue(0);

        return hwpFile.getDocInfo().getBorderFillList().size();
    }

    private long mmToHwp(double mm) {
        return (long) (mm * 72000.0f / 254.0f + 0.5f);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelToHwp().setVisible(true));
    }
}
