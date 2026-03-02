package com.ruoyi.web.controller.tool;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import com.ruoyi.common.config.RuoYiConfig;
import com.ruoyi.common.core.domain.AjaxResult;
import com.ruoyi.common.utils.StringUtils;
import com.ruoyi.common.utils.file.FileUploadUtils;
import com.ruoyi.common.utils.file.FileUtils;

@RestController
@RequestMapping("/tool/excel-editor")
public class ExcelEditorController
{
    private static final Logger log = LoggerFactory.getLogger(ExcelEditorController.class);

    private static final String STORAGE_DIR = "excel-editor";

    private static final String CURRENT_FILE_MARKER = ".current";

    private static final String[] ALLOWED_EXTENSIONS = { "xls", "xlsx" };

    private static final DateTimeFormatter[] DATE_TIME_FORMATTERS = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm")
    };

    private static final DateTimeFormatter[] DATE_FORMATTERS = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("yyyy.MM.dd")
    };

    @GetMapping("/info")
    public AjaxResult info()
    {
        try
        {
            Map<String, Object> data = new LinkedHashMap<String, Object>();
            List<Map<String, Object>> files = listStoredFiles();
            Optional<Path> currentFile = resolveCurrentFile();
            data.put("exists", currentFile.isPresent());
            data.put("currentFileName", currentFile.map(path -> path.getFileName().toString()).orElse(""));
            data.put("files", files);
            if (currentFile.isPresent())
            {
                data.putAll(buildFileInfo(currentFile.get()));
            }
            return AjaxResult.success(data);
        }
        catch (IOException e)
        {
            log.error("Failed to read Excel storage info", e);
            return AjaxResult.error("Failed to read Excel storage info");
        }
    }

    @GetMapping("/content")
    public void content(@RequestParam(value = "fileName", required = false) String fileName, HttpServletResponse response)
            throws IOException
    {
        Optional<Path> targetFile = StringUtils.isNotEmpty(fileName) ? resolveFileByName(fileName) : resolveCurrentFile();
        if (!targetFile.isPresent())
        {
            response.sendError(HttpServletResponse.SC_NOT_FOUND, "No Excel file is available");
            return;
        }
        writeCurrentMarker(targetFile.get().getFileName().toString());
        writeBinaryFile(response, targetFile.get());
    }

    @GetMapping("/view")
    public AjaxResult view(@RequestParam(value = "fileName", required = false) String fileName)
    {
        try
        {
            Optional<Path> targetFile = StringUtils.isNotEmpty(fileName) ? resolveFileByName(fileName) : resolveCurrentFile();
            if (!targetFile.isPresent())
            {
                return AjaxResult.error("No Excel file is available");
            }
            writeCurrentMarker(targetFile.get().getFileName().toString());
            return AjaxResult.success(buildWorkbookView(targetFile.get()));
        }
        catch (Exception e)
        {
            log.error("Failed to build workbook view", e);
            return AjaxResult.error("Failed to build workbook view");
        }
    }

    @PostMapping("/upload")
    public AjaxResult upload(@RequestParam("file") MultipartFile file)
    {
        try
        {
            Path storedFile = storeUploadedWorkbook(file);
            return AjaxResult.success("Excel uploaded to server storage", buildUploadResponse(storedFile));
        }
        catch (Exception e)
        {
            log.error("Failed to upload Excel file", e);
            return AjaxResult.error(e.getMessage());
        }
    }

    @PostMapping("/save")
    public AjaxResult save(@RequestParam("file") MultipartFile file,
            @RequestParam(value = "fileName", required = false) String fileName)
    {
        try
        {
            Path storedFile = saveWorkbook(file, fileName);
            return AjaxResult.success("Excel saved to server storage", buildUploadResponse(storedFile));
        }
        catch (Exception e)
        {
            log.error("Failed to save Excel file", e);
            return AjaxResult.error(e.getMessage());
        }
    }

    @PostMapping("/apply")
    public AjaxResult apply(@RequestBody WorkbookPatchRequest request)
    {
        try
        {
            if (request == null || StringUtils.isEmpty(request.getFileName()))
            {
                return AjaxResult.error("File name is required");
            }

            Optional<Path> targetFile = resolveFileByName(request.getFileName());
            if (!targetFile.isPresent())
            {
                return AjaxResult.error("The selected Excel file does not exist");
            }

            applyWorkbookChanges(targetFile.get(), request.getChanges());
            return AjaxResult.success("Excel saved and original formatting was preserved", buildUploadResponse(targetFile.get()));
        }
        catch (Exception e)
        {
            log.error("Failed to apply Excel changes", e);
            return AjaxResult.error("Failed to apply Excel changes");
        }
    }

    private Map<String, Object> buildUploadResponse(Path storedFile) throws IOException
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.putAll(buildFileInfo(storedFile));
        data.put("currentFileName", storedFile.getFileName().toString());
        data.put("files", listStoredFiles());
        return data;
    }

    private Map<String, Object> buildWorkbookView(Path workbookPath) throws Exception
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.putAll(buildFileInfo(workbookPath));
        data.put("fileName", workbookPath.getFileName().toString());
        data.put("currentFileName", workbookPath.getFileName().toString());

        try (InputStream inputStream = Files.newInputStream(workbookPath);
                Workbook workbook = WorkbookFactory.create(inputStream))
        {
            DataFormatter formatter = new DataFormatter(Locale.getDefault());
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            List<Map<String, Object>> sheets = new ArrayList<Map<String, Object>>();
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheets.add(buildSheetView(workbook, workbook.getSheetAt(sheetIndex), formatter, evaluator));
            }
            data.put("sheetCount", workbook.getNumberOfSheets());
            data.put("sheets", sheets);
        }
        return data;
    }

    private Map<String, Object> buildSheetView(Workbook workbook, Sheet sheet, DataFormatter formatter,
            FormulaEvaluator evaluator)
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.put("name", sheet.getSheetName());

        int lastRowIndex = resolveLastActiveRowIndex(sheet);
        int maxColumnCount = resolveMaxColumnCount(sheet);
        Map<String, CellRangeAddress> mergedStartMap = new HashMap<String, CellRangeAddress>();
        Set<String> mergedCoveredCells = new HashSet<String>();

        for (int mergeIndex = 0; mergeIndex < sheet.getNumMergedRegions(); mergeIndex++)
        {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(mergeIndex);
            mergedStartMap.put(buildCellKey(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn()), mergedRegion);
            for (int rowIndex = mergedRegion.getFirstRow(); rowIndex <= mergedRegion.getLastRow(); rowIndex++)
            {
                for (int columnIndex = mergedRegion.getFirstColumn(); columnIndex <= mergedRegion.getLastColumn(); columnIndex++)
                {
                    if (rowIndex != mergedRegion.getFirstRow() || columnIndex != mergedRegion.getFirstColumn())
                    {
                        mergedCoveredCells.add(buildCellKey(rowIndex, columnIndex));
                    }
                }
            }
        }

        List<Map<String, Object>> rows = new ArrayList<Map<String, Object>>();
        for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);
            Map<String, Object> rowData = new LinkedHashMap<String, Object>();
            rowData.put("rowIndex", rowIndex);
            rowData.put("heightPx", resolveRowHeightPx(sheet, rowIndex));

            List<Map<String, Object>> cells = new ArrayList<Map<String, Object>>();
            for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
            {
                if (mergedCoveredCells.contains(buildCellKey(rowIndex, columnIndex)))
                {
                    continue;
                }

                CellRangeAddress mergedRegion = mergedStartMap.get(buildCellKey(rowIndex, columnIndex));
                int rowSpan = mergedRegion == null ? 1 : mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                int colSpan = mergedRegion == null ? 1 : mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                Cell cell = row == null ? null : row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                cells.add(buildCellView(workbook, sheet, row, cell, rowIndex, columnIndex, rowSpan, colSpan, formatter,
                        evaluator));
            }

            rowData.put("cells", cells);
            rows.add(rowData);
        }

        data.put("rowCount", rows.size());
        data.put("maxColumnCount", maxColumnCount);
        data.put("rows", rows);
        return data;
    }

    private Map<String, Object> buildCellView(Workbook workbook, Sheet sheet, Row row, Cell cell, int rowIndex,
            int columnIndex, int rowSpan, int colSpan, DataFormatter formatter, FormulaEvaluator evaluator)
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.put("rowIndex", rowIndex);
        data.put("colIndex", columnIndex);
        data.put("rowSpan", rowSpan);
        data.put("colSpan", colSpan);
        data.put("value", resolveEditableValue(cell, formatter, evaluator));
        data.put("displayValue", resolveDisplayValue(cell, formatter, evaluator));
        data.put("formula", isFormulaCell(cell));
        data.put("widthPx", resolveWidthPx(sheet, columnIndex, colSpan));
        data.put("heightPx", resolveHeightPx(sheet, rowIndex, rowSpan));
        data.put("style", buildStyleView(workbook, row, cell));
        return data;
    }

    private Map<String, Object> buildStyleView(Workbook workbook, Row row, Cell cell)
    {
        Map<String, Object> style = new LinkedHashMap<String, Object>();
        CellStyle cellStyle = resolveCellStyle(row, cell);
        if (cellStyle == null)
        {
            return style;
        }

        String backgroundColor = resolveFillColor(cellStyle);
        if (StringUtils.isNotEmpty(backgroundColor))
        {
            style.put("backgroundColor", backgroundColor);
        }

        Font font = workbook.getFontAt(cellStyle.getFontIndexAsInt());
        if (font != null)
        {
            String fontColor = resolveFontColor(workbook, font);
            if (StringUtils.isNotEmpty(fontColor))
            {
                style.put("color", fontColor);
            }
            if (font.getBold())
            {
                style.put("fontWeight", "700");
            }
            if (font.getItalic())
            {
                style.put("fontStyle", "italic");
            }
            if (font.getFontHeightInPoints() > 0)
            {
                style.put("fontSize", font.getFontHeightInPoints() + "pt");
            }
            if (StringUtils.isNotEmpty(font.getFontName()))
            {
                style.put("fontFamily", font.getFontName());
            }
            if (font.getUnderline() != Font.U_NONE)
            {
                style.put("textDecoration", "underline");
            }
        }

        if (cellStyle.getAlignment() != null)
        {
            switch (cellStyle.getAlignment())
            {
                case CENTER:
                case CENTER_SELECTION:
                    style.put("textAlign", "center");
                    break;
                case RIGHT:
                    style.put("textAlign", "right");
                    break;
                case FILL:
                case JUSTIFY:
                case DISTRIBUTED:
                    style.put("textAlign", "justify");
                    break;
                default:
                    style.put("textAlign", "left");
                    break;
            }
        }

        if (cellStyle.getVerticalAlignment() != null)
        {
            switch (cellStyle.getVerticalAlignment())
            {
                case TOP:
                    style.put("verticalAlign", "top");
                    break;
                case CENTER:
                    style.put("verticalAlign", "middle");
                    break;
                case JUSTIFY:
                case DISTRIBUTED:
                    style.put("verticalAlign", "middle");
                    break;
                default:
                    style.put("verticalAlign", "bottom");
                    break;
            }
        }

        if (cellStyle.getWrapText())
        {
            style.put("whiteSpace", "pre-wrap");
        }
        return style;
    }

    private void applyWorkbookChanges(Path workbookPath, List<CellPatch> changes) throws Exception
    {
        Path tempFile = null;
        try (InputStream inputStream = Files.newInputStream(workbookPath);
                Workbook workbook = WorkbookFactory.create(inputStream))
        {
            if (changes != null)
            {
                for (CellPatch change : changes)
                {
                    applySingleChange(workbook, change);
                }
            }

            workbook.setForceFormulaRecalculation(true);
            tempFile = Files.createTempFile(getStorageDirectory(), "excel-editor-", ".tmp");
            try (OutputStream outputStream = Files.newOutputStream(tempFile))
            {
                workbook.write(outputStream);
                outputStream.flush();
            }
        }
        catch (Exception e)
        {
            if (tempFile != null)
            {
                Files.deleteIfExists(tempFile);
            }
            throw e;
        }

        Files.move(tempFile, workbookPath, StandardCopyOption.REPLACE_EXISTING);
        writeCurrentMarker(workbookPath.getFileName().toString());
    }

    private void applySingleChange(Workbook workbook, CellPatch change)
    {
        if (change == null || StringUtils.isEmpty(change.getSheetName()) || change.getRowIndex() == null
                || change.getColIndex() == null)
        {
            return;
        }

        Sheet sheet = workbook.getSheet(change.getSheetName());
        if (sheet == null)
        {
            return;
        }

        Row row = sheet.getRow(change.getRowIndex());
        if (row == null)
        {
            row = sheet.createRow(change.getRowIndex());
        }

        Cell cell = row.getCell(change.getColIndex(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        applyCellValue(cell, change.getValue());
    }

    private void applyCellValue(Cell cell, String inputValue)
    {
        String safeValue = inputValue == null ? "" : inputValue;
        if (safeValue.startsWith("=") && safeValue.length() > 1)
        {
            cell.setCellFormula(safeValue.substring(1));
            return;
        }

        CellType originalType = cell.getCellType();
        if (originalType == CellType.FORMULA)
        {
            originalType = cell.getCachedFormulaResultType();
            cell.removeFormula();
        }

        if (StringUtils.isEmpty(safeValue))
        {
            cell.setBlank();
            return;
        }

        if (originalType == CellType.BOOLEAN)
        {
            if ("true".equalsIgnoreCase(safeValue) || "false".equalsIgnoreCase(safeValue))
            {
                cell.setCellValue(Boolean.parseBoolean(safeValue));
                return;
            }
        }

        if (originalType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell))
        {
            Date parsedDate = tryParseDateValue(safeValue);
            if (parsedDate != null)
            {
                cell.setCellValue(parsedDate);
                return;
            }
        }

        if (originalType == CellType.NUMERIC || originalType == CellType.BLANK)
        {
            Double numericValue = tryParseNumericValue(safeValue);
            if (numericValue != null && !looksLikePlainTextNumber(safeValue))
            {
                cell.setCellValue(numericValue.doubleValue());
                return;
            }
        }

        cell.setCellValue(safeValue);
    }

    private Date tryParseDateValue(String inputValue)
    {
        for (DateTimeFormatter formatter : DATE_TIME_FORMATTERS)
        {
            try
            {
                LocalDateTime dateTime = LocalDateTime.parse(inputValue, formatter);
                return Date.from(dateTime.atZone(ZoneId.systemDefault()).toInstant());
            }
            catch (DateTimeParseException e)
            {
                // Ignore and continue.
            }
        }

        for (DateTimeFormatter formatter : DATE_FORMATTERS)
        {
            try
            {
                LocalDate date = LocalDate.parse(inputValue, formatter);
                return Date.from(LocalDateTime.of(date, LocalTime.MIN).atZone(ZoneId.systemDefault()).toInstant());
            }
            catch (DateTimeParseException e)
            {
                // Ignore and continue.
            }
        }
        return null;
    }

    private Double tryParseNumericValue(String inputValue)
    {
        try
        {
            return new BigDecimal(inputValue.trim()).doubleValue();
        }
        catch (Exception e)
        {
            return null;
        }
    }

    private boolean looksLikePlainTextNumber(String inputValue)
    {
        String value = inputValue == null ? "" : inputValue.trim();
        if (value.length() <= 1 || !value.matches("[-+]?0\\d+"))
        {
            return false;
        }
        return !value.startsWith("0.");
    }

    private String resolveEditableValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator)
    {
        if (cell == null)
        {
            return "";
        }
        if (cell.getCellType() == CellType.FORMULA)
        {
            return "=" + cell.getCellFormula();
        }
        return formatter.formatCellValue(cell, evaluator);
    }

    private String resolveDisplayValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator)
    {
        if (cell == null)
        {
            return "";
        }
        return formatter.formatCellValue(cell, evaluator);
    }

    private boolean isFormulaCell(Cell cell)
    {
        return cell != null && cell.getCellType() == CellType.FORMULA;
    }

    private CellStyle resolveCellStyle(Row row, Cell cell)
    {
        if (cell != null)
        {
            return cell.getCellStyle();
        }
        if (row != null && row.isFormatted())
        {
            return row.getRowStyle();
        }
        return null;
    }

    private String resolveFillColor(CellStyle cellStyle)
    {
        if (cellStyle == null || cellStyle.getFillPattern() == FillPatternType.NO_FILL)
        {
            return "";
        }

        String foreground = toCssColor(cellStyle.getFillForegroundColorColor());
        if (StringUtils.isNotEmpty(foreground))
        {
            return foreground;
        }
        return toCssColor(cellStyle.getFillBackgroundColorColor());
    }

    private String resolveFontColor(Workbook workbook, Font font)
    {
        if (font instanceof XSSFFont)
        {
            return toCssColor(((XSSFFont) font).getXSSFColor());
        }
        if (font instanceof HSSFFont && workbook instanceof HSSFWorkbook)
        {
            HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            HSSFColor color = palette.getColor(((HSSFFont) font).getColor());
            return toCssColor(color);
        }
        return "";
    }

    private String toCssColor(Color color)
    {
        if (color == null)
        {
            return "";
        }
        if (color instanceof XSSFColor)
        {
            byte[] rgb = ((XSSFColor) color).getRGB();
            return rgb == null ? "" : toCssColor(rgb[0], rgb[1], rgb[2]);
        }
        if (color instanceof HSSFColor)
        {
            short[] triplet = ((HSSFColor) color).getTriplet();
            return triplet == null ? "" : toCssColor((byte) triplet[0], (byte) triplet[1], (byte) triplet[2]);
        }
        return "";
    }

    private String toCssColor(byte red, byte green, byte blue)
    {
        return String.format("#%02X%02X%02X", red & 0xFF, green & 0xFF, blue & 0xFF);
    }

    private int resolveLastActiveRowIndex(Sheet sheet)
    {
        int lastRowIndex = -1;
        for (Row row : sheet)
        {
            lastRowIndex = Math.max(lastRowIndex, row.getRowNum());
        }
        for (int mergeIndex = 0; mergeIndex < sheet.getNumMergedRegions(); mergeIndex++)
        {
            lastRowIndex = Math.max(lastRowIndex, sheet.getMergedRegion(mergeIndex).getLastRow());
        }
        return Math.max(lastRowIndex, 0);
    }

    private int resolveMaxColumnCount(Sheet sheet)
    {
        int maxColumnCount = 1;
        for (Row row : sheet)
        {
            if (row.getLastCellNum() > 0)
            {
                maxColumnCount = Math.max(maxColumnCount, row.getLastCellNum());
            }
        }
        for (int mergeIndex = 0; mergeIndex < sheet.getNumMergedRegions(); mergeIndex++)
        {
            maxColumnCount = Math.max(maxColumnCount, sheet.getMergedRegion(mergeIndex).getLastColumn() + 1);
        }
        return maxColumnCount;
    }

    private int resolveRowHeightPx(Sheet sheet, int rowIndex)
    {
        Row row = sheet.getRow(rowIndex);
        float heightInPoints = row == null ? sheet.getDefaultRowHeightInPoints() : row.getHeightInPoints();
        return Math.max(24, Math.round(heightInPoints * 96F / 72F));
    }

    private int resolveHeightPx(Sheet sheet, int startRowIndex, int rowSpan)
    {
        int total = 0;
        for (int offset = 0; offset < rowSpan; offset++)
        {
            total += resolveRowHeightPx(sheet, startRowIndex + offset);
        }
        return Math.max(total, 24);
    }

    private int resolveWidthPx(Sheet sheet, int startColumnIndex, int colSpan)
    {
        int total = 0;
        for (int offset = 0; offset < colSpan; offset++)
        {
            total += Math.max(72, approximateColumnWidthPx(sheet, startColumnIndex + offset));
        }
        return total;
    }

    private int approximateColumnWidthPx(Sheet sheet, int columnIndex)
    {
        double widthUnits = sheet.getColumnWidth(columnIndex);
        return (int) Math.round((widthUnits / 256D) * 7D + 12D);
    }

    private String buildCellKey(int rowIndex, int columnIndex)
    {
        return rowIndex + ":" + columnIndex;
    }

    private Path storeUploadedWorkbook(MultipartFile file) throws IOException
    {
        validateExcelFile(file);
        Path storageDir = getStorageDirectory();
        String targetFileName = buildUniqueFileName(storageDir, file.getOriginalFilename());
        Path targetFile = storageDir.resolve(targetFileName);
        copyMultipartFile(file, targetFile);
        writeCurrentMarker(targetFileName);
        return targetFile;
    }

    private Path saveWorkbook(MultipartFile file, String requestedFileName) throws IOException
    {
        validateExcelFile(file);
        Path storageDir = getStorageDirectory();
        Optional<Path> currentFile = StringUtils.isNotEmpty(requestedFileName) ? resolveFileByName(requestedFileName)
                : resolveCurrentFile();
        String targetFileName = currentFile.map(path -> path.getFileName().toString())
                .orElse(buildSafeFileName(file.getOriginalFilename()));
        Path targetFile = storageDir.resolve(targetFileName);
        copyMultipartFile(file, targetFile);
        writeCurrentMarker(targetFileName);
        return targetFile;
    }

    private void copyMultipartFile(MultipartFile file, Path targetFile) throws IOException
    {
        try (InputStream inputStream = file.getInputStream())
        {
            Files.copy(inputStream, targetFile, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private void validateExcelFile(MultipartFile file) throws IOException
    {
        if (file == null || file.isEmpty())
        {
            throw new IOException("Please select an Excel file");
        }
        try
        {
            FileUploadUtils.assertAllowed(file, ALLOWED_EXTENSIONS);
        }
        catch (Exception e)
        {
            throw new IOException("Only .xls and .xlsx files are supported", e);
        }
    }

    private Path getStorageDirectory() throws IOException
    {
        Path storageDir = Paths.get(RuoYiConfig.getProfile(), STORAGE_DIR);
        if (!Files.exists(storageDir))
        {
            Files.createDirectories(storageDir);
        }
        return storageDir;
    }

    private Path getMarkerPath() throws IOException
    {
        return getStorageDirectory().resolve(CURRENT_FILE_MARKER);
    }

    private void writeCurrentMarker(String fileName) throws IOException
    {
        Files.write(getMarkerPath(), fileName.getBytes(StandardCharsets.UTF_8));
    }

    private Optional<String> readCurrentMarker() throws IOException
    {
        Path markerPath = getMarkerPath();
        if (!Files.exists(markerPath))
        {
            return Optional.empty();
        }
        String fileName = new String(Files.readAllBytes(markerPath), StandardCharsets.UTF_8).trim();
        return StringUtils.isEmpty(fileName) ? Optional.empty() : Optional.of(fileName);
    }

    private Optional<Path> resolveCurrentFile() throws IOException
    {
        Optional<String> markerFileName = readCurrentMarker();
        if (markerFileName.isPresent())
        {
            Optional<Path> target = resolveFileByName(markerFileName.get());
            if (target.isPresent())
            {
                return target;
            }
        }
        List<Path> files = listStoredPaths();
        if (files.isEmpty())
        {
            return Optional.empty();
        }
        Path latestFile = files.get(0);
        writeCurrentMarker(latestFile.getFileName().toString());
        return Optional.of(latestFile);
    }

    private Optional<Path> resolveFileByName(String fileName) throws IOException
    {
        if (StringUtils.isEmpty(fileName))
        {
            return Optional.empty();
        }
        Path storageDir = getStorageDirectory();
        String safeFileName = FilenameUtils.getName(fileName);
        Path targetFile = storageDir.resolve(safeFileName);
        if (!Files.exists(targetFile) || !Files.isRegularFile(targetFile) || !isExcelFile(targetFile))
        {
            return Optional.empty();
        }
        return Optional.of(targetFile);
    }

    private List<Path> listStoredPaths() throws IOException
    {
        Path storageDir = getStorageDirectory();
        try (Stream<Path> stream = Files.list(storageDir))
        {
            return stream.filter(Files::isRegularFile).filter(this::isExcelFile)
                    .sorted(Comparator.comparingLong(this::lastModified).reversed()).collect(Collectors.toList());
        }
    }

    private List<Map<String, Object>> listStoredFiles() throws IOException
    {
        List<Map<String, Object>> files = new ArrayList<Map<String, Object>>();
        for (Path path : listStoredPaths())
        {
            files.add(buildFileInfo(path));
        }
        return files;
    }

    private boolean isExcelFile(Path path)
    {
        String extension = FilenameUtils.getExtension(path.getFileName().toString());
        return "xls".equalsIgnoreCase(extension) || "xlsx".equalsIgnoreCase(extension);
    }

    private long lastModified(Path path)
    {
        try
        {
            return Files.getLastModifiedTime(path).toMillis();
        }
        catch (IOException e)
        {
            return Long.MIN_VALUE;
        }
    }

    private String buildUniqueFileName(Path storageDir, String originalFileName)
    {
        String safeFileName = buildSafeFileName(originalFileName);
        Path candidate = storageDir.resolve(safeFileName);
        if (!Files.exists(candidate))
        {
            return safeFileName;
        }

        String baseName = FilenameUtils.getBaseName(safeFileName);
        String extension = FilenameUtils.getExtension(safeFileName);
        return baseName + "_" + System.currentTimeMillis() + "." + extension;
    }

    private String buildSafeFileName(String originalFileName)
    {
        String sourceName = StringUtils.isNotEmpty(originalFileName) ? FilenameUtils.getName(originalFileName)
                : "excel-editor.xlsx";
        String baseName = FilenameUtils.getBaseName(sourceName);
        String extension = FilenameUtils.getExtension(sourceName);

        if (StringUtils.isEmpty(baseName))
        {
            baseName = "excel-editor";
        }
        if (!"xls".equalsIgnoreCase(extension) && !"xlsx".equalsIgnoreCase(extension))
        {
            extension = "xlsx";
        }

        String normalizedBaseName = baseName.replaceAll("[\\\\/:*?\"<>|\\s]+", "_").replaceAll("_{2,}", "_");
        if (StringUtils.isEmpty(normalizedBaseName))
        {
            normalizedBaseName = "excel-editor";
        }
        if (normalizedBaseName.length() > 80)
        {
            normalizedBaseName = normalizedBaseName.substring(0, 80);
        }
        return normalizedBaseName + "." + extension.toLowerCase();
    }

    private Map<String, Object> buildFileInfo(Path file) throws IOException
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.put("fileName", file.getFileName().toString());
        data.put("size", Files.size(file));
        data.put("lastModified", Files.getLastModifiedTime(file).toMillis());
        return data;
    }

    private void writeBinaryFile(HttpServletResponse response, Path file) throws IOException
    {
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setDateHeader("Expires", 0);
        response.setHeader("X-Excel-File-Name", file.getFileName().toString());
        try (OutputStream outputStream = response.getOutputStream())
        {
            FileUtils.setAttachmentResponseHeader(response, file.getFileName().toString());
            Files.copy(file, outputStream);
            outputStream.flush();
        }
    }

    public static class WorkbookPatchRequest
    {
        private String fileName;

        private List<CellPatch> changes;

        public String getFileName()
        {
            return fileName;
        }

        public void setFileName(String fileName)
        {
            this.fileName = fileName;
        }

        public List<CellPatch> getChanges()
        {
            return changes;
        }

        public void setChanges(List<CellPatch> changes)
        {
            this.changes = changes;
        }
    }

    public static class CellPatch
    {
        private String sheetName;

        private Integer rowIndex;

        private Integer colIndex;

        private String value;

        public String getSheetName()
        {
            return sheetName;
        }

        public void setSheetName(String sheetName)
        {
            this.sheetName = sheetName;
        }

        public Integer getRowIndex()
        {
            return rowIndex;
        }

        public void setRowIndex(Integer rowIndex)
        {
            this.rowIndex = rowIndex;
        }

        public Integer getColIndex()
        {
            return colIndex;
        }

        public void setColIndex(Integer colIndex)
        {
            this.colIndex = colIndex;
        }

        public String getValue()
        {
            return value;
        }

        public void setValue(String value)
        {
            this.value = value;
        }
    }
}
