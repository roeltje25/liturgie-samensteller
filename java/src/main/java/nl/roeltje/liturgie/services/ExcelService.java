package nl.roeltje.liturgie.services;

import nl.roeltje.liturgie.models.Liturgy;
import nl.roeltje.liturgie.models.LiturgySection;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Manages the LiederenRegister.xlsx song-usage tracking spreadsheet.
 *
 * Two-pass pattern (mirrors the Python implementation):
 *   Pass 1: open with evaluateAllFormulaCells() to read formula results
 *   Pass 2: reopen for editing and write new data
 */
public class ExcelService {

    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);

    private static final String SHEET_LITURGY  = "Liturgie";
    private static final String SHEET_LIEDEREN = "Liederen";

    private static final DateTimeFormatter DATE_ISO = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    // ── Export ─────────────────────────────────────────────────────────────────

    public void exportLiturgy(Liturgy liturgy, Path excelPath) throws IOException {
        if (excelPath == null || !Files.exists(excelPath)) {
            log.warn("Excel register not found: {}", excelPath);
            return;
        }

        String serviceDate = liturgy.getService_date();
        String leider = liturgy.getDienstleider();
        List<String> songTitles = collectSongTitles(liturgy);

        // Write to temp file, then replace
        Path tmp = excelPath.resolveSibling(excelPath.getFileName() + ".tmp");
        try {
            Files.copy(excelPath, tmp, StandardCopyOption.REPLACE_EXISTING);

            try (InputStream in = Files.newInputStream(tmp);
                 XSSFWorkbook wb = new XSSFWorkbook(in)) {

                Sheet liturgieSheet = findSheet(wb, SHEET_LITURGY);
                Sheet liederenSheet = findSheet(wb, SHEET_LIEDEREN);

                if (liturgieSheet != null) {
                    writeToLiturgySheet(liturgieSheet, serviceDate, leider, songTitles);
                }
                if (liederenSheet != null) {
                    ensureSongsExist(liederenSheet, songTitles);
                }

                try (OutputStream out = Files.newOutputStream(excelPath,
                        StandardOpenOption.TRUNCATE_EXISTING)) {
                    wb.write(out);
                }
            }
        } finally {
            Files.deleteIfExists(tmp);
        }
        log.info("Excel register updated: {}", excelPath);
    }

    // ── Dienstleiders autocomplete ─────────────────────────────────────────────

    public List<String> getDienstleiders(Path excelPath) {
        if (excelPath == null || !Files.exists(excelPath)) return List.of();
        Set<String> result = new LinkedHashSet<>();
        try (InputStream in = Files.newInputStream(excelPath);
             XSSFWorkbook wb = new XSSFWorkbook(in)) {
            Sheet sheet = findSheet(wb, SHEET_LITURGY);
            if (sheet == null) return List.of();
            Map<String, Integer> cols = buildColumnMap(sheet);
            Integer col = cols.get("dienstleider");
            if (col == null) return List.of();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                Cell cell = row.getCell(col);
                if (cell != null) {
                    String val = getCellString(cell);
                    if (!val.isBlank()) result.add(val.trim());
                }
            }
        } catch (IOException e) {
            log.warn("Cannot read dienstleiders from {}", excelPath, e);
        }
        return new ArrayList<>(result);
    }

    // ── Registered songs ───────────────────────────────────────────────────────

    public Set<String> getRegisteredSongs(Path excelPath) {
        if (excelPath == null || !Files.exists(excelPath)) return Set.of();
        Set<String> result = new LinkedHashSet<>();
        try (InputStream in = Files.newInputStream(excelPath);
             XSSFWorkbook wb = new XSSFWorkbook(in)) {
            Sheet sheet = findSheet(wb, SHEET_LIEDEREN);
            if (sheet == null) return Set.of();
            Map<String, Integer> cols = buildColumnMap(sheet);
            Integer col = cols.get("titel");
            if (col == null) col = cols.get("naam");
            if (col == null) col = 0;
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                Cell cell = row.getCell(col);
                if (cell != null) {
                    String val = getCellString(cell);
                    if (!val.isBlank()) result.add(val.trim().toLowerCase());
                }
            }
        } catch (IOException e) {
            log.warn("Cannot read registered songs from {}", excelPath, e);
        }
        return result;
    }

    // ── Private helpers ────────────────────────────────────────────────────────

    private List<String> collectSongTitles(Liturgy liturgy) {
        List<String> titles = new ArrayList<>();
        for (LiturgySection section : liturgy.getSections()) {
            if (section.isSong()) {
                titles.add(section.getName());
            }
        }
        return titles;
    }

    private void writeToLiturgySheet(Sheet sheet, String serviceDate, String leider,
                                     List<String> songs) {
        Map<String, Integer> cols = buildColumnMap(sheet);
        // Find or create the row for this service date
        int targetRow = findOrCreateRow(sheet, serviceDate, cols);
        Row row = sheet.getRow(targetRow);
        if (row == null) row = sheet.createRow(targetRow);

        setCell(row, cols.getOrDefault("zondag", 0), serviceDate);
        setCell(row, cols.getOrDefault("dienstleider", 1), leider != null ? leider : "");
        setCell(row, cols.getOrDefault("liturgie_bekend", 2), "ja");

        for (int i = 0; i < songs.size() && i < 12; i++) {
            String key = "lied" + (i + 1);
            int col = cols.getOrDefault(key, 3 + i);
            setCell(row, col, songs.get(i));
        }
    }

    private void ensureSongsExist(Sheet sheet, List<String> songs) {
        Set<String> existing = new LinkedHashSet<>();
        Map<String, Integer> cols = buildColumnMap(sheet);
        int titleCol = cols.getOrDefault("titel", cols.getOrDefault("naam", 0));

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;
            Cell c = row.getCell(titleCol);
            if (c != null) existing.add(getCellString(c).toLowerCase().trim());
        }

        for (String song : songs) {
            if (!existing.contains(song.toLowerCase().trim())) {
                int nextRow = sheet.getLastRowNum() + 1;
                Row newRow = sheet.createRow(nextRow);
                setCell(newRow, titleCol, song);
                existing.add(song.toLowerCase().trim());
            }
        }
    }

    private int findOrCreateRow(Sheet sheet, String serviceDate, Map<String, Integer> cols) {
        int dateCol = cols.getOrDefault("zondag", 0);
        if (serviceDate != null) {
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                Cell c = row.getCell(dateCol);
                if (c != null && serviceDate.equals(getCellString(c).trim())) {
                    return row.getRowNum();
                }
            }
        }
        return sheet.getLastRowNum() + 1;
    }

    private Map<String, Integer> buildColumnMap(Sheet sheet) {
        Map<String, Integer> map = new HashMap<>();
        Row header = sheet.getRow(0);
        if (header == null) return map;
        for (Cell cell : header) {
            String key = getCellString(cell).toLowerCase().trim()
                    .replace(" ", "_").replace("-", "_");
            map.put(key, cell.getColumnIndex());
        }
        return map;
    }

    private Sheet findSheet(Workbook wb, String name) {
        Sheet s = wb.getSheet(name);
        if (s != null) return s;
        // Case-insensitive fallback
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (wb.getSheetName(i).equalsIgnoreCase(name)) return wb.getSheetAt(i);
        }
        return null;
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield DATE_ISO.format(cell.getLocalDateTimeCellValue().toLocalDate());
                }
                double d = cell.getNumericCellValue();
                yield d == Math.floor(d) ? String.valueOf((long) d) : String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA  -> {
                try { yield cell.getStringCellValue(); }
                catch (Exception e) { yield String.valueOf(cell.getNumericCellValue()); }
            }
            default -> "";
        };
    }

    private void setCell(Row row, int col, String value) {
        Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
    }
}
