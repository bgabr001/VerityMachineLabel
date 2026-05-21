import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class MachineLable {

    public static void main(String[] args) throws Exception {

        String inputFolder = "counties";
        String outputFile = "Machine_Labels.pdf";

        List<Label> labels = new ArrayList<>();

        File folder = new File(inputFolder);

        File[] files = folder.listFiles((dir, name) -> {
            String lowerName = name.toLowerCase();

            return lowerName.endsWith(".xlsx")
                    && !lowerName.startsWith("~$");
        });

        if (files == null || files.length == 0) {
            System.out.println("No county Excel files found.");
            return;
        }

        Arrays.sort(files, (a, b) -> a.getName().compareToIgnoreCase(b.getName()));

        for (File file : files) {

            String countyName = getCountyName(file);

            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);

            for (Sheet sheet : workbook) {
                for (Row row : sheet) {

                    String pollingPlace = getCellText(row.getCell(0)).trim();

                    if (pollingPlace.isBlank()) {
                        continue;
                    }

                    String lower = pollingPlace.toLowerCase();

                    if (lower.contains("polling place")
                            || lower.contains("machines used")
                            || lower.contains("total")) {
                        continue;
                    }

                    int scans = getNumericValue(row.getCell(2));
                    int ada = getNumericValue(row.getCell(3));
                    int prints = getNumericValue(row.getCell(4));

                    for (int i = 0; i < scans; i++) {
                        labels.add(new Label(countyName, pollingPlace, "SCAN"));
                    }

                    for (int i = 0; i < ada; i++) {
                        labels.add(new Label(countyName, pollingPlace, "ADA"));
                    }

                    for (int i = 0; i < prints; i++) {
                        labels.add(new Label(countyName, pollingPlace, "PRINT"));
                    }
                }
            }

            workbook.close();
            fis.close();

            System.out.println("Processed: " + file.getName());
        }

        createPdf(labels, outputFile);

        System.out.println();
        System.out.println("DONE!");
        System.out.println("Created: " + outputFile);
        System.out.println("Total labels: " + labels.size());
    }

    private static void createPdf(List<Label> labels, String outputFile) throws Exception {

        float cardWidth = 3 * 72;
        float cardHeight = 2 * 72;

        PDDocument document = new PDDocument();

        PDType1Font boldFont =
                new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD);

        for (Label label : labels) {

            PDPage page = new PDPage(new PDRectangle(cardWidth, cardHeight));
            document.addPage(page);

            PDPageContentStream content =
                    new PDPageContentStream(document, page);

            drawCenteredText(content, boldFont, 16,
                    label.county, cardWidth, 100);

            drawCenteredText(content, boldFont, 13,
                    label.pollingPlace, cardWidth, 72);

            drawCenteredText(content, boldFont, 20,
                    label.machineType, cardWidth, 38);

            content.close();
        }

        document.save(outputFile);
        document.close();
    }

    private static void drawCenteredText(
            PDPageContentStream content,
            PDType1Font font,
            int fontSize,
            String text,
            float pageWidth,
            float y
    ) throws Exception {

        if (text == null) {
            text = "";
        }

        text = text.trim();

        if (text.length() > 32) {
            text = text.substring(0, 32);
        }

        float textWidth = font.getStringWidth(text) / 1000 * fontSize;
        float x = (pageWidth - textWidth) / 2;

        content.beginText();
        content.setFont(font, fontSize);
        content.newLineAtOffset(x, y);
        content.showText(text);
        content.endText();
    }

    private static int getNumericValue(Cell cell) {

        if (cell == null) {
            return 0;
        }

        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return (int) cell.getNumericCellValue();
            }

            if (cell.getCellType() == CellType.FORMULA) {
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    return (int) cell.getNumericCellValue();
                }

                if (cell.getCachedFormulaResultType() == CellType.STRING) {
                    String text = cell.getStringCellValue().trim();
                    return Integer.parseInt(text);
                }
            }

            if (cell.getCellType() == CellType.STRING) {
                String text = cell.getStringCellValue().trim();
                text = text.replace(",", "");
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
            if (cell.getCellType() == CellType.STRING) {
                return cell.getStringCellValue();
            }

            if (cell.getCellType() == CellType.NUMERIC) {
                return String.valueOf((int) cell.getNumericCellValue());
            }

            if (cell.getCellType() == CellType.FORMULA) {
                if (cell.getCachedFormulaResultType() == CellType.STRING) {
                    return cell.getStringCellValue();
                }

                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    return String.valueOf((int) cell.getNumericCellValue());
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

        Label(String county, String pollingPlace, String machineType) {
            this.county = county;
            this.pollingPlace = pollingPlace;
            this.machineType = machineType;
        }
    }
}