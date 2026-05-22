import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.*;
import org.apache.pdfbox.pdmodel.PDPageContentStream;

import java.io.*;
import java.util.*;

public class MachineLable {

    public static void main(String[] args) throws Exception {

        String inputFolder = "2026PrimaryCounties";
        String outputRootFolder = "output_labels";

        File outputRoot = new File(outputRootFolder);
        outputRoot.mkdirs();

        File folder = new File(inputFolder);

        File[] files = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx")
                        && !name.startsWith("~$"));

        if (files == null || files.length == 0) {
            System.out.println("No county Excel files found.");
            return;
        }

        Arrays.sort(files, (a, b) ->
                a.getName().compareToIgnoreCase(b.getName()));

        for (File file : files) {

            String countyName = getCountyName(file);

            List<Label> labels =
                    readLabelsFromCountyFile(file, countyName);

            File countyFolder =
                    new File(outputRoot, countyName);

            countyFolder.mkdirs();

            String outputPdf =
                    countyFolder.getPath()
                            + File.separator
                            + countyName + "_Labels.pdf";

            createPdf(labels, outputPdf);

            System.out.println(
                    "Created: " + outputPdf
                            + " | Labels: " + labels.size());
        }

        System.out.println();
        System.out.println("DONE!");
    }

    private static List<Label> readLabelsFromCountyFile(
            File file,
            String countyName
    ) throws Exception {

        List<Label> labels = new ArrayList<>();

        FileInputStream fis =
                new FileInputStream(file);

        Workbook workbook =
                new XSSFWorkbook(fis);

        for (Sheet sheet : workbook) {

            for (Row row : sheet) {

                String pollingPlace =
                        getCellText(row.getCell(0)).trim();

                if (pollingPlace.isBlank()) {
                    continue;
                }

                String lower =
                        pollingPlace.toLowerCase();

                // Ignore headers/totals
                if (lower.contains("polling place")
                        || lower.contains("machines used")
                        || lower.contains("total")) {
                    continue;
                }

                // Skip Machines Owned row
                if (lower.contains("machine owned")
                        || lower.contains("machines owned")) {
                    continue;
                }

                // Rename Spares row
                if (lower.contains("spares")) {
                    pollingPlace = "SPARE";
                }

                int scans =
                        getNumericValue(row.getCell(2));

                int ada =
                        getNumericValue(row.getCell(3));

                int prints =
                        getNumericValue(row.getCell(4));

                boolean isSpare =
                        pollingPlace.equalsIgnoreCase("SPARE");

                // SCAN labels (printed TWICE)
                for (int i = 1; i <= scans; i++) {

                    String labelText =
                            isSpare
                                    ? "SCAN"
                                    : "SCAN " + i;

                    labels.add(
                            new Label(
                                    countyName,
                                    pollingPlace,
                                    labelText
                            )
                    );

                    labels.add(
                            new Label(
                                    countyName,
                                    pollingPlace,
                                    labelText
                            )
                    );
                }

                // ADA labels
                for (int i = 1; i <= ada; i++) {

                    String labelText =
                            isSpare
                                    ? "ADA"
                                    : "ADA " + i;

                    labels.add(
                            new Label(
                                    countyName,
                                    pollingPlace,
                                    labelText
                            )
                    );
                }

                // PRINT labels
                for (int i = 1; i <= prints; i++) {

                    String labelText =
                            isSpare
                                    ? "PRINT"
                                    : "PRINT " + i;

                    labels.add(
                            new Label(
                                    countyName,
                                    pollingPlace,
                                    labelText
                            )
                    );
                }
            }
        }

        workbook.close();
        fis.close();

        return labels;
    }

    private static void createPdf(
            List<Label> labels,
            String outputFile
    ) throws Exception {

        float pageWidth = 11 * 72;
        float pageHeight = 8.5f * 72;

        // 4x2 labels
        float labelWidth = 4 * 72;
        float labelHeight = 2 * 72;

        float leftMargin = 1.25f * 72;
        float topMargin = 0.25f * 72;

        int labelsAcross = 2;
        int labelsDown = 4;

        int labelsPerPage =
                labelsAcross * labelsDown;

        PDDocument document =
                new PDDocument();

        PDType1Font boldFont =
                new PDType1Font(
                        Standard14Fonts.FontName.HELVETICA_BOLD);

        for (int i = 0; i < labels.size(); i++) {

            if (i % labelsPerPage == 0) {

                document.addPage(
                        new PDPage(
                                new PDRectangle(
                                        pageWidth,
                                        pageHeight
                                )
                        )
                );
            }

            PDPage page =
                    document.getPage(
                            document.getNumberOfPages() - 1);

            PDPageContentStream content =
                    new PDPageContentStream(
                            document,
                            page,
                            PDPageContentStream.AppendMode.APPEND,
                            true
                    );

            int positionOnPage =
                    i % labelsPerPage;

            int column =
                    positionOnPage % labelsAcross;

            int row =
                    positionOnPage / labelsAcross;

            float x =
                    leftMargin + (column * labelWidth);

            float y =
                    pageHeight
                            - topMargin
                            - labelHeight
                            - (row * labelHeight);

            drawLabel(
                    content,
                    boldFont,
                    labels.get(i),
                    x,
                    y,
                    labelWidth,
                    labelHeight
            );

            content.close();
        }

        document.save(outputFile);
        document.close();
    }

    private static void drawLabel(
            PDPageContentStream content,
            PDType1Font font,
            Label label,
            float x,
            float y,
            float width,
            float height
    ) throws Exception {

        // Border
        content.setLineWidth(1);

        content.addRect(
                x,
                y,
                width,
                height
        );

        content.stroke();

        // County
        drawCenteredText(
                content,
                font,
                18,
                label.county,
                x,
                y + 105,
                width
        );

        // Polling place
        drawCenteredText(
                content,
                font,
                12,
                label.pollingPlace,
                x,
                y + 78,
                width
        );

        // Machine type
        drawCenteredText(
                content,
                font,
                22,
                label.machineType,
                x,
                y + 28,
                width
        );
    }

    private static void drawCenteredText(
            PDPageContentStream content,
            PDType1Font font,
            int fontSize,
            String text,
            float labelX,
            float startY,
            float labelWidth
    ) throws Exception {

        if (text == null) {
            text = "";
        }

        text = text.trim();

        float maxTextWidth =
                labelWidth - 20;

        List<String> lines =
                new ArrayList<>();

        String[] words =
                text.split(" ");

        StringBuilder currentLine =
                new StringBuilder();

        for (String word : words) {

            String testLine;

            if (currentLine.isEmpty()) {
                testLine = word;
            } else {
                testLine =
                        currentLine + " " + word;
            }

            float textWidth =
                    font.getStringWidth(testLine)
                            / 1000 * fontSize;

            if (textWidth > maxTextWidth) {

                if (!currentLine.isEmpty()) {

                    lines.add(
                            currentLine.toString()
                    );
                }

                currentLine =
                        new StringBuilder(word);

            } else {

                if (!currentLine.isEmpty()) {
                    currentLine.append(" ");
                }

                currentLine.append(word);
            }
        }

        if (!currentLine.isEmpty()) {
            lines.add(currentLine.toString());
        }

        float lineHeight =
                fontSize + 4;

        float y = startY;

        for (String line : lines) {

            float textWidth =
                    font.getStringWidth(line)
                            / 1000 * fontSize;

            float x =
                    labelX
                            + ((labelWidth - textWidth) / 2);

            content.beginText();

            content.setFont(
                    font,
                    fontSize
            );

            content.newLineAtOffset(
                    x,
                    y
            );

            content.showText(line);

            content.endText();

            y -= lineHeight;
        }
    }

    private static int getNumericValue(Cell cell) {

        if (cell == null) {
            return 0;
        }

        try {

            if (cell.getCellType()
                    == CellType.NUMERIC) {

                return (int)
                        cell.getNumericCellValue();
            }

            if (cell.getCellType()
                    == CellType.FORMULA) {

                if (cell.getCachedFormulaResultType()
                        == CellType.NUMERIC) {

                    return (int)
                            cell.getNumericCellValue();
                }
            }

            if (cell.getCellType()
                    == CellType.STRING) {

                String text =
                        cell.getStringCellValue()
                                .trim()
                                .replace(",", "");

                return Integer.parseInt(text);
            }

        } catch (Exception e) {
            return 0;
        }

        return 0;
    }

    private static String getCellText(Cell cell) {

        if (cell == null) {
            return "";
        }

        try {

            if (cell.getCellType()
                    == CellType.STRING) {

                return cell.getStringCellValue();
            }

            if (cell.getCellType()
                    == CellType.NUMERIC) {

                return String.valueOf(
                        (int)
                                cell.getNumericCellValue()
                );
            }

            if (cell.getCellType()
                    == CellType.FORMULA) {

                if (cell.getCachedFormulaResultType()
                        == CellType.STRING) {

                    return cell.getStringCellValue();
                }

                if (cell.getCachedFormulaResultType()
                        == CellType.NUMERIC) {

                    return String.valueOf(
                            (int)
                                    cell.getNumericCellValue()
                    );
                }
            }

        } catch (Exception e) {
            return "";
        }

        return "";
    }

    private static String getCountyName(File file) {

        return file.getName()
                .replace(".xlsx", "")
                .replace(".xls", "")
                .replace(" Election Plan", "");
    }

    static class Label {

        String county;
        String pollingPlace;
        String machineType;

        Label(
                String county,
                String pollingPlace,
                String machineType
        ) {

            this.county = county;
            this.pollingPlace = pollingPlace;
            this.machineType = machineType;
        }
    }
}