package com.ruoyi.web.controller.tool;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.ByteArrayInputStream;
import java.awt.image.BufferedImage;
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
import java.util.Base64;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
import com.ruoyi.common.utils.SecurityUtils;
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

    /**
     * Excel single-cell text hard limit (characters), shared by xls/xlsx.
     */
    private static final int MAX_CELL_TEXT_LENGTH = 32767;

    /**
     * Defensive upper bound to avoid pathological payloads.
     */
    private static final int MAX_PATCH_CHANGES = 100000;

    private static final int MIN_ROW_HEIGHT_PX = 24;

    private static final int MAX_ROW_HEIGHT_PX = 1200;

    private static final int MIN_COLUMN_WIDTH_PX = 40;

    private static final int MAX_COLUMN_WIDTH_PX = 1600;

    private static final long LOCK_TIMEOUT_MILLIS = 5 * 60 * 1000L;

    private static final Map<String, FileEditLock> FILE_EDIT_LOCKS = new ConcurrentHashMap<String, FileEditLock>();

    private static final Object FILE_EDIT_LOCK_MONITOR = new Object();

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

    private static final Pattern BACKGROUND_PICTURE_RELATION_PATTERN = Pattern
            .compile("<(?:\\w+:)?picture[^>]*\\br:id\\s*=\\s*\"([^\"]+)\"", Pattern.CASE_INSENSITIVE);

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

    @GetMapping("/lock")
    public AjaxResult lockStatus(@RequestParam("fileName") String fileName)
    {
        try
        {
            Optional<Path> targetFile = resolveFileByName(fileName);
            if (!targetFile.isPresent())
            {
                return AjaxResult.error("The selected Excel file does not exist");
            }
            EditorIdentity editor = resolveEditorIdentity();
            String safeFileName = targetFile.get().getFileName().toString();
            FileEditLock currentLock = getActiveLock(safeFileName);
            return AjaxResult.success("Lock status loaded",
                    buildLockResponseData(safeFileName, currentLock, editor, false));
        }
        catch (Exception e)
        {
            log.error("Failed to read lock status", e);
            return AjaxResult.error("Failed to read lock status");
        }
    }

    @PostMapping("/lock")
    public AjaxResult lock(@RequestBody FileLockRequest request)
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

            EditorIdentity editor = resolveEditorIdentity();
            String safeFileName = targetFile.get().getFileName().toString();
            LockOperationResult lockResult = acquireFileLock(safeFileName, editor, request.isForce());
            if (!lockResult.isSuccess())
            {
                return AjaxResult.error(lockResult.getMessage(),
                        buildLockResponseData(safeFileName, lockResult.getLock(), editor, false));
            }
            return AjaxResult.success("Lock acquired", buildLockResponseData(safeFileName, lockResult.getLock(), editor, true));
        }
        catch (Exception e)
        {
            log.error("Failed to acquire file lock", e);
            return AjaxResult.error("Failed to acquire file lock");
        }
    }

    @PostMapping("/lock/heartbeat")
    public AjaxResult lockHeartbeat(@RequestBody FileLockRequest request)
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

            EditorIdentity editor = resolveEditorIdentity();
            String safeFileName = targetFile.get().getFileName().toString();
            LockOperationResult lockResult = refreshFileLock(safeFileName, editor);
            if (!lockResult.isSuccess())
            {
                return AjaxResult.error(lockResult.getMessage(),
                        buildLockResponseData(safeFileName, lockResult.getLock(), editor, false));
            }
            return AjaxResult.success("Lock heartbeat accepted",
                    buildLockResponseData(safeFileName, lockResult.getLock(), editor, true));
        }
        catch (Exception e)
        {
            log.error("Failed to refresh file lock", e);
            return AjaxResult.error("Failed to refresh file lock");
        }
    }

    @PostMapping("/lock/release")
    public AjaxResult lockRelease(@RequestBody FileLockRequest request)
    {
        try
        {
            if (request == null || StringUtils.isEmpty(request.getFileName()))
            {
                return AjaxResult.error("File name is required");
            }

            String safeFileName = FilenameUtils.getName(request.getFileName());
            EditorIdentity editor = resolveEditorIdentity();
            LockOperationResult releaseResult = releaseFileLock(safeFileName, editor, request.isForce());
            if (!releaseResult.isSuccess())
            {
                return AjaxResult.error(releaseResult.getMessage(),
                        buildLockResponseData(safeFileName, releaseResult.getLock(), editor, false));
            }
            return AjaxResult.success("Lock released", buildLockResponseData(safeFileName, null, editor, true));
        }
        catch (Exception e)
        {
            log.error("Failed to release file lock", e);
            return AjaxResult.error("Failed to release file lock");
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

            EditorIdentity editor = resolveEditorIdentity();
            String safeFileName = targetFile.get().getFileName().toString();
            LockValidationResult lockValidation = validateLockForWrite(safeFileName, editor);
            if (!lockValidation.isAllowed())
            {
                return AjaxResult.error(lockValidation.getMessage(),
                        buildLockResponseData(safeFileName, lockValidation.getLock(), editor, false));
            }

            if (StringUtils.isEmpty(request.getVersion()))
            {
                Map<String, Object> data = buildUploadResponse(targetFile.get());
                data.put("versionConflict", true);
                return AjaxResult.error("Workbook version is required. Please reload and retry.", data);
            }
            String currentVersion = buildFileVersion(targetFile.get());
            if (!StringUtils.equals(request.getVersion(), currentVersion))
            {
                Map<String, Object> data = buildUploadResponse(targetFile.get());
                data.put("versionConflict", true);
                return AjaxResult.error("Workbook has changed on server. Please reload before saving.", data);
            }

            applyWorkbookChanges(targetFile.get(), request.getChanges(), request.getMergeRegions(), request.getRowHeights(),
                    request.getColumnWidths());
            return AjaxResult.success("Excel saved and original formatting was preserved", buildUploadResponse(targetFile.get()));
        }
        catch (IllegalArgumentException e)
        {
            return AjaxResult.error(e.getMessage());
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
        data.put("version", buildFileVersion(storedFile));
        data.put("files", listStoredFiles());
        return data;
    }

    private Map<String, Object> buildWorkbookView(Path workbookPath) throws Exception
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.putAll(buildFileInfo(workbookPath));
        data.put("fileName", workbookPath.getFileName().toString());
        data.put("currentFileName", workbookPath.getFileName().toString());
        data.put("version", buildFileVersion(workbookPath));

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

    private String buildFileVersion(Path filePath) throws IOException
    {
        long lastModified = Files.getLastModifiedTime(filePath).toMillis();
        long fileSize = Files.size(filePath);
        return lastModified + ":" + fileSize;
    }

    private EditorIdentity resolveEditorIdentity()
    {
        Long userId = SecurityUtils.getUserId();
        String username = SecurityUtils.getUsername();
        if (StringUtils.isEmpty(username))
        {
            username = userId == null ? "Anonymous" : "User-" + userId;
        }
        return new EditorIdentity(userId, username);
    }

    private LockOperationResult acquireFileLock(String fileName, EditorIdentity editor, boolean forceTakeover)
    {
        long now = System.currentTimeMillis();
        synchronized (FILE_EDIT_LOCK_MONITOR)
        {
            cleanupExpiredLocks(now);
            FileEditLock current = FILE_EDIT_LOCKS.get(fileName);
            if (current == null || current.isOwnedBy(editor) || forceTakeover)
            {
                FileEditLock next = new FileEditLock(fileName, editor.getUserId(), editor.getUsername(), now,
                        now + LOCK_TIMEOUT_MILLIS);
                FILE_EDIT_LOCKS.put(fileName, next);
                return LockOperationResult.success(next);
            }
            return LockOperationResult.failure(current,
                    "File is currently being edited by " + current.getOwnerUsername() + ".");
        }
    }

    private LockOperationResult refreshFileLock(String fileName, EditorIdentity editor)
    {
        long now = System.currentTimeMillis();
        synchronized (FILE_EDIT_LOCK_MONITOR)
        {
            cleanupExpiredLocks(now);
            FileEditLock current = FILE_EDIT_LOCKS.get(fileName);
            if (current == null)
            {
                return LockOperationResult.failure(null, "Edit lock expired. Please re-open the file.");
            }
            if (!current.isOwnedBy(editor))
            {
                return LockOperationResult.failure(current,
                        "File is currently being edited by " + current.getOwnerUsername() + ".");
            }
            current.setExpiresAt(now + LOCK_TIMEOUT_MILLIS);
            current.setUpdatedAt(now);
            return LockOperationResult.success(current);
        }
    }

    private LockOperationResult releaseFileLock(String fileName, EditorIdentity editor, boolean forceRelease)
    {
        long now = System.currentTimeMillis();
        synchronized (FILE_EDIT_LOCK_MONITOR)
        {
            cleanupExpiredLocks(now);
            FileEditLock current = FILE_EDIT_LOCKS.get(fileName);
            if (current == null)
            {
                return LockOperationResult.success(null);
            }
            if (!current.isOwnedBy(editor) && !forceRelease)
            {
                return LockOperationResult.failure(current,
                        "File is currently being edited by " + current.getOwnerUsername() + ".");
            }
            FILE_EDIT_LOCKS.remove(fileName);
            return LockOperationResult.success(null);
        }
    }

    private FileEditLock getActiveLock(String fileName)
    {
        long now = System.currentTimeMillis();
        synchronized (FILE_EDIT_LOCK_MONITOR)
        {
            cleanupExpiredLocks(now);
            return FILE_EDIT_LOCKS.get(fileName);
        }
    }

    private LockValidationResult validateLockForWrite(String fileName, EditorIdentity editor)
    {
        long now = System.currentTimeMillis();
        synchronized (FILE_EDIT_LOCK_MONITOR)
        {
            cleanupExpiredLocks(now);
            FileEditLock current = FILE_EDIT_LOCKS.get(fileName);
            if (current == null)
            {
                return LockValidationResult.failure(null,
                        "Edit lock not found. Please reopen the workbook.");
            }
            if (!current.isOwnedBy(editor))
            {
                return LockValidationResult.failure(current,
                        "File is currently being edited by " + current.getOwnerUsername() + ".");
            }
            current.setExpiresAt(now + LOCK_TIMEOUT_MILLIS);
            current.setUpdatedAt(now);
            return LockValidationResult.success(current);
        }
    }

    private void cleanupExpiredLocks(long now)
    {
        Iterator<Map.Entry<String, FileEditLock>> iterator = FILE_EDIT_LOCKS.entrySet().iterator();
        while (iterator.hasNext())
        {
            Map.Entry<String, FileEditLock> entry = iterator.next();
            FileEditLock lock = entry.getValue();
            if (lock == null || lock.getExpiresAt() <= now)
            {
                iterator.remove();
            }
        }
    }

    private Map<String, Object> buildLockResponseData(String fileName, FileEditLock lock, EditorIdentity editor,
            boolean acquired)
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.put("fileName", fileName);
        data.put("timeoutMs", LOCK_TIMEOUT_MILLIS);
        data.put("acquired", acquired);
        if (lock == null)
        {
            data.put("locked", false);
            data.put("self", false);
            return data;
        }

        long now = System.currentTimeMillis();
        data.put("locked", true);
        data.put("self", lock.isOwnedBy(editor));
        data.put("ownerUserId", lock.getOwnerUserId());
        data.put("ownerUsername", lock.getOwnerUsername());
        data.put("createdAt", lock.getCreatedAt());
        data.put("updatedAt", lock.getUpdatedAt());
        data.put("expiresAt", lock.getExpiresAt());
        data.put("remainingMs", Math.max(lock.getExpiresAt() - now, 0L));
        return data;
    }

    private Map<String, Object> buildSheetView(Workbook workbook, Sheet sheet, DataFormatter formatter,
            FormulaEvaluator evaluator)
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        data.put("name", sheet.getSheetName());
        Map<String, Object> freezePane = buildFreezePaneView(sheet);
        if (!freezePane.isEmpty())
        {
            data.put("freezePane", freezePane);
        }
        List<Map<String, Object>> validations = buildValidationView(workbook, sheet, formatter, evaluator);
        if (!validations.isEmpty())
        {
            data.put("validations", validations);
        }
        List<Map<String, Object>> images = buildSheetImageView(workbook, sheet);
        if (!images.isEmpty())
        {
            data.put("images", images);
        }

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
                Cell cell = row == null ? null : row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
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
        data.put("style", buildStyleView(workbook, sheet, row, cell, columnIndex));
        return data;
    }

    private Map<String, Object> buildStyleView(Workbook workbook, Sheet sheet, Row row, Cell cell, int columnIndex)
    {
        Map<String, Object> style = new LinkedHashMap<String, Object>();
        CellStyle cellStyle = resolveCellStyle(sheet, row, cell, columnIndex);
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
        else
        {
            style.put("whiteSpace", "nowrap");
        }

        style.putAll(buildBorderStyle(workbook, cellStyle));
        return style;
    }

    private void applyWorkbookChanges(Path workbookPath, List<CellPatch> changes, List<MergeRegionPatch> mergeRegions,
            List<RowHeightPatch> rowHeights, List<ColumnWidthPatch> columnWidths) throws Exception
    {
        if (changes != null && changes.size() > MAX_PATCH_CHANGES)
        {
            throw new IllegalArgumentException("Too many changed cells in one save request");
        }

        Path tempFile = null;
        try (InputStream inputStream = Files.newInputStream(workbookPath);
                Workbook workbook = WorkbookFactory.create(inputStream))
        {
            DataFormatter formatter = new DataFormatter(Locale.getDefault());
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            if (changes != null)
            {
                for (CellPatch change : changes)
                {
                    applySingleChange(workbook, change, formatter, evaluator);
                }
            }

            if (mergeRegions != null)
            {
                applyWorkbookMergeLayout(workbook, mergeRegions);
            }

            if (rowHeights != null)
            {
                applyWorkbookRowHeights(workbook, rowHeights);
            }

            if (columnWidths != null)
            {
                applyWorkbookColumnWidths(workbook, columnWidths);
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

    private void applyWorkbookMergeLayout(Workbook workbook, List<MergeRegionPatch> mergeRegions)
    {
        Map<String, List<CellRangeAddress>> nextRegionsBySheet = new HashMap<String, List<CellRangeAddress>>();
        Set<String> dedupeKeys = new HashSet<String>();

        for (MergeRegionPatch mergeRegion : mergeRegions)
        {
            CellRangeAddress region = toCellRangeAddress(mergeRegion);
            if (region == null)
            {
                continue;
            }

            Sheet sheet = workbook.getSheet(mergeRegion.getSheetName());
            if (sheet == null)
            {
                continue;
            }

            String dedupeKey = sheet.getSheetName() + ":" + buildMergeRegionKey(region);
            if (!dedupeKeys.add(dedupeKey))
            {
                continue;
            }

            nextRegionsBySheet.computeIfAbsent(sheet.getSheetName(), key -> new ArrayList<CellRangeAddress>()).add(region);
        }

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            List<CellRangeAddress> nextRegions = nextRegionsBySheet.getOrDefault(sheet.getSheetName(),
                    new ArrayList<CellRangeAddress>());
            nextRegions.sort(Comparator.comparingInt(CellRangeAddress::getFirstRow)
                    .thenComparingInt(CellRangeAddress::getFirstColumn).thenComparingInt(CellRangeAddress::getLastRow)
                    .thenComparingInt(CellRangeAddress::getLastColumn));

            Set<String> originalRegionKeys = new HashSet<String>();
            for (int mergeIndex = 0; mergeIndex < sheet.getNumMergedRegions(); mergeIndex++)
            {
                originalRegionKeys.add(buildMergeRegionKey(sheet.getMergedRegion(mergeIndex)));
            }

            clearMergedRegions(sheet);

            for (CellRangeAddress region : nextRegions)
            {
                if (!originalRegionKeys.contains(buildMergeRegionKey(region)))
                {
                    clearMergedCoveredCells(sheet, region);
                }
                sheet.addMergedRegion(region);
            }
        }
    }

    private void applyWorkbookRowHeights(Workbook workbook, List<RowHeightPatch> rowHeights)
    {
        for (RowHeightPatch patch : rowHeights)
        {
            if (patch == null || StringUtils.isEmpty(patch.getSheetName()) || patch.getRowIndex() == null
                    || patch.getHeightPx() == null)
            {
                continue;
            }

            Sheet sheet = workbook.getSheet(patch.getSheetName());
            if (sheet == null)
            {
                continue;
            }

            int rowIndex = patch.getRowIndex().intValue();
            if (rowIndex < 0 || rowIndex > resolveLastActiveRowIndex(sheet))
            {
                throw new IllegalArgumentException("Row out of sheet bounds: " + sheet.getSheetName() + "!"
                        + (rowIndex + 1));
            }

            int safeHeightPx = Math.max(MIN_ROW_HEIGHT_PX, Math.min(MAX_ROW_HEIGHT_PX, patch.getHeightPx().intValue()));
            Row row = sheet.getRow(rowIndex);
            if (row == null)
            {
                row = sheet.createRow(rowIndex);
            }
            row.setHeightInPoints(pixelToPoints(safeHeightPx));
        }
    }

    private void applyWorkbookColumnWidths(Workbook workbook, List<ColumnWidthPatch> columnWidths)
    {
        for (ColumnWidthPatch patch : columnWidths)
        {
            if (patch == null || StringUtils.isEmpty(patch.getSheetName()) || patch.getColIndex() == null
                    || patch.getWidthPx() == null)
            {
                continue;
            }

            Sheet sheet = workbook.getSheet(patch.getSheetName());
            if (sheet == null)
            {
                continue;
            }

            int colIndex = patch.getColIndex().intValue();
            int maxColumnCount = resolveMaxColumnCount(sheet);
            if (colIndex < 0 || colIndex >= maxColumnCount)
            {
                throw new IllegalArgumentException("Column out of sheet bounds: " + sheet.getSheetName() + "!"
                        + toColumnLabel(colIndex));
            }

            int safeWidthPx = Math.max(MIN_COLUMN_WIDTH_PX, Math.min(MAX_COLUMN_WIDTH_PX, patch.getWidthPx().intValue()));
            sheet.setColumnWidth(colIndex, pixelToExcelWidthUnits(safeWidthPx));
        }
    }

    private float pixelToPoints(int heightPx)
    {
        return (float) (heightPx * 72D / 96D);
    }

    private int pixelToExcelWidthUnits(int widthPx)
    {
        int units = (int) Math.round(((Math.max(widthPx, 1) - 12D) / 7D) * 256D);
        int minUnits = 256;
        int maxUnits = 255 * 256;
        return Math.max(minUnits, Math.min(maxUnits, units));
    }

    private CellRangeAddress toCellRangeAddress(MergeRegionPatch mergeRegion)
    {
        if (mergeRegion == null || StringUtils.isEmpty(mergeRegion.getSheetName()) || mergeRegion.getFirstRow() == null
                || mergeRegion.getLastRow() == null || mergeRegion.getFirstColumn() == null
                || mergeRegion.getLastColumn() == null)
        {
            return null;
        }

        int firstRow = Math.max(mergeRegion.getFirstRow().intValue(), 0);
        int lastRow = Math.max(mergeRegion.getLastRow().intValue(), firstRow);
        int firstColumn = Math.max(mergeRegion.getFirstColumn().intValue(), 0);
        int lastColumn = Math.max(mergeRegion.getLastColumn().intValue(), firstColumn);
        return new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
    }

    private void clearMergedRegions(Sheet sheet)
    {
        for (int mergeIndex = sheet.getNumMergedRegions() - 1; mergeIndex >= 0; mergeIndex--)
        {
            sheet.removeMergedRegion(mergeIndex);
        }
    }

    private void clearMergedCoveredCells(Sheet sheet, CellRangeAddress region)
    {
        for (int rowIndex = region.getFirstRow(); rowIndex <= region.getLastRow(); rowIndex++)
        {
            for (int columnIndex = region.getFirstColumn(); columnIndex <= region.getLastColumn(); columnIndex++)
            {
                if (rowIndex == region.getFirstRow() && columnIndex == region.getFirstColumn())
                {
                    continue;
                }

                Row row = sheet.getRow(rowIndex);
                if (row == null)
                {
                    row = sheet.createRow(rowIndex);
                }
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setBlank();
            }
        }
    }

    private String buildMergeRegionKey(CellRangeAddress region)
    {
        return region.getFirstRow() + ":" + region.getLastRow() + ":" + region.getFirstColumn() + ":"
                + region.getLastColumn();
    }

    private void applySingleChange(Workbook workbook, CellPatch change, DataFormatter formatter,
            FormulaEvaluator evaluator)
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

        int rowIndex = change.getRowIndex().intValue();
        int columnIndex = change.getColIndex().intValue();
        if (rowIndex < 0 || columnIndex < 0)
        {
            throw new IllegalArgumentException("Cell index cannot be negative");
        }

        int maxRowIndex = resolveLastActiveRowIndex(sheet);
        int maxColumnCount = resolveMaxColumnCount(sheet);
        int maxColumnIndex = Math.max(maxColumnCount - 1, 0);
        if (rowIndex > maxRowIndex || columnIndex > maxColumnIndex)
        {
            throw new IllegalArgumentException("Cell out of sheet bounds: " + sheet.getSheetName() + "!"
                    + formatCellAddress(rowIndex, columnIndex));
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null)
        {
            row = sheet.createRow(rowIndex);
        }

        ValidationRule validationRule = resolveListValidationRuleForCell(workbook, sheet, rowIndex, columnIndex, formatter,
                evaluator);
        if (validationRule != null && !validationRule.getOptions().isEmpty())
        {
            String rawValue = change.getValue() == null ? "" : change.getValue();
            String normalizedValue = rawValue.trim();
            if (StringUtils.isEmpty(normalizedValue))
            {
                if (!validationRule.isAllowBlank())
                {
                    throw new IllegalArgumentException(
                            "Cell " + sheet.getSheetName() + "!" + formatCellAddress(rowIndex, columnIndex)
                                    + " does not allow blank value");
                }
            }
            else if (!validationRule.getOptions().contains(normalizedValue))
            {
                throw new IllegalArgumentException(
                        "Cell " + sheet.getSheetName() + "!" + formatCellAddress(rowIndex, columnIndex)
                                + " must use one of the dropdown options");
            }
        }

        Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        applyCellValue(cell, change.getValue());
    }

    private void applyCellValue(Cell cell, String inputValue)
    {
        String safeValue = inputValue == null ? "" : inputValue;
        if (safeValue.length() > MAX_CELL_TEXT_LENGTH)
        {
            throw new IllegalArgumentException("Cell text exceeds Excel limit of " + MAX_CELL_TEXT_LENGTH
                    + " characters");
        }
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

    private CellStyle resolveCellStyle(Sheet sheet, Row row, Cell cell, int columnIndex)
    {
        if (cell != null)
        {
            return cell.getCellStyle();
        }
        if (row != null && row.isFormatted())
        {
            CellStyle rowStyle = row.getRowStyle();
            if (rowStyle != null)
            {
                return rowStyle;
            }
        }
        if (sheet != null)
        {
            CellStyle columnStyle = sheet.getColumnStyle(columnIndex);
            if (columnStyle != null)
            {
                return columnStyle;
            }
        }
        return null;
    }

    private Map<String, Object> buildFreezePaneView(Sheet sheet)
    {
        Map<String, Object> data = new LinkedHashMap<String, Object>();
        PaneInformation pane = sheet.getPaneInformation();
        if (pane == null || !pane.isFreezePane())
        {
            return data;
        }

        data.put("xSplit", pane.getVerticalSplitPosition());
        data.put("ySplit", pane.getHorizontalSplitPosition());
        data.put("leftColumn", pane.getVerticalSplitLeftColumn());
        data.put("topRow", pane.getHorizontalSplitTopRow());
        return data;
    }

    private List<Map<String, Object>> buildValidationView(Workbook workbook, Sheet sheet, DataFormatter formatter,
            FormulaEvaluator evaluator)
    {
        List<Map<String, Object>> items = new ArrayList<Map<String, Object>>();
        List<? extends DataValidation> validations = sheet.getDataValidations();
        for (DataValidation validation : validations)
        {
            DataValidationConstraint constraint = validation.getValidationConstraint();
            if (constraint == null || constraint.getValidationType() != DataValidationConstraint.ValidationType.LIST)
            {
                continue;
            }

            List<String> options = resolveValidationOptions(workbook, sheet, constraint, formatter, evaluator);
            if (options.isEmpty())
            {
                continue;
            }

            CellRangeAddressList regions = validation.getRegions();
            if (regions == null || regions.countRanges() == 0)
            {
                continue;
            }

            for (CellRangeAddress region : regions.getCellRangeAddresses())
            {
                Map<String, Object> item = new LinkedHashMap<String, Object>();
                item.put("firstRow", region.getFirstRow());
                item.put("lastRow", region.getLastRow());
                item.put("firstColumn", region.getFirstColumn());
                item.put("lastColumn", region.getLastColumn());
                item.put("options", options);
                item.put("allowBlank", validation.getEmptyCellAllowed());
                items.add(item);
            }
        }
        return items;
    }

    private ValidationRule resolveListValidationRuleForCell(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex,
            DataFormatter formatter, FormulaEvaluator evaluator)
    {
        List<? extends DataValidation> validations = sheet.getDataValidations();
        for (DataValidation validation : validations)
        {
            DataValidationConstraint constraint = validation.getValidationConstraint();
            if (constraint == null || constraint.getValidationType() != DataValidationConstraint.ValidationType.LIST)
            {
                continue;
            }

            CellRangeAddressList regions = validation.getRegions();
            if (regions == null || regions.countRanges() == 0)
            {
                continue;
            }

            boolean inRegion = false;
            for (CellRangeAddress region : regions.getCellRangeAddresses())
            {
                if (rowIndex >= region.getFirstRow() && rowIndex <= region.getLastRow() && columnIndex >= region.getFirstColumn()
                        && columnIndex <= region.getLastColumn())
                {
                    inRegion = true;
                    break;
                }
            }
            if (!inRegion)
            {
                continue;
            }

            List<String> options = resolveValidationOptions(workbook, sheet, constraint, formatter, evaluator);
            if (options.isEmpty())
            {
                continue;
            }
            return new ValidationRule(options, validation.getEmptyCellAllowed());
        }
        return null;
    }

    private List<String> resolveValidationOptions(Workbook workbook, Sheet sheet, DataValidationConstraint constraint,
            DataFormatter formatter, FormulaEvaluator evaluator)
    {
        String[] explicitValues = constraint.getExplicitListValues();
        if (explicitValues != null && explicitValues.length > 0)
        {
            return normalizeValidationOptions(explicitValues);
        }

        String formula = constraint.getFormula1();
        if (StringUtils.isEmpty(formula))
        {
            return new ArrayList<String>();
        }

        String normalizedFormula = formula.trim();
        if (normalizedFormula.startsWith("\"") && normalizedFormula.endsWith("\"") && normalizedFormula.length() >= 2)
        {
            String csvText = normalizedFormula.substring(1, normalizedFormula.length() - 1);
            if (StringUtils.isEmpty(csvText))
            {
                return new ArrayList<String>();
            }
            return normalizeValidationOptions(csvText.split(","));
        }

        List<String> namedRangeOptions = resolveValidationOptionsFromNamedRange(workbook, sheet, normalizedFormula,
                formatter, evaluator);
        if (!namedRangeOptions.isEmpty())
        {
            return namedRangeOptions;
        }

        return resolveValidationOptionsFromRange(workbook, sheet, normalizedFormula, formatter, evaluator);
    }

    private List<String> resolveValidationOptionsFromNamedRange(Workbook workbook, Sheet currentSheet, String expression,
            DataFormatter formatter, FormulaEvaluator evaluator)
    {
        String nameKey = expression == null ? "" : expression.trim();
        if (StringUtils.isEmpty(nameKey))
        {
            return new ArrayList<String>();
        }
        if (nameKey.startsWith("="))
        {
            nameKey = nameKey.substring(1).trim();
        }
        if (StringUtils.isEmpty(nameKey) || nameKey.contains("!") || nameKey.contains("("))
        {
            return new ArrayList<String>();
        }

        int currentSheetIndex = workbook.getSheetIndex(currentSheet);
        Name matched = null;
        for (Name name : workbook.getAllNames())
        {
            if (name == null || StringUtils.isEmpty(name.getNameName()))
            {
                continue;
            }
            if (!name.getNameName().equalsIgnoreCase(nameKey))
            {
                continue;
            }
            int sheetIndex = name.getSheetIndex();
            if (sheetIndex == currentSheetIndex)
            {
                matched = name;
                break;
            }
            if (matched == null && sheetIndex == -1)
            {
                matched = name;
            }
        }

        if (matched == null || StringUtils.isEmpty(matched.getRefersToFormula()))
        {
            return new ArrayList<String>();
        }

        String referred = matched.getRefersToFormula().trim();
        if (StringUtils.isEmpty(referred) || referred.equalsIgnoreCase(nameKey))
        {
            return new ArrayList<String>();
        }
        return resolveValidationOptionsFromRange(workbook, currentSheet, referred, formatter, evaluator);
    }

    private List<String> normalizeValidationOptions(String[] values)
    {
        Set<String> dedupe = new LinkedHashSet<String>();
        if (values == null)
        {
            return new ArrayList<String>();
        }
        for (String value : values)
        {
            String normalized = value == null ? "" : value.trim();
            if (StringUtils.isNotEmpty(normalized))
            {
                dedupe.add(normalized);
            }
        }
        return new ArrayList<String>(dedupe);
    }

    private List<String> resolveValidationOptionsFromRange(Workbook workbook, Sheet currentSheet, String formula,
            DataFormatter formatter, FormulaEvaluator evaluator)
    {
        String expression = formula == null ? "" : formula.trim();
        if (StringUtils.isEmpty(expression))
        {
            return new ArrayList<String>();
        }

        if (expression.startsWith("="))
        {
            expression = expression.substring(1).trim();
        }

        String sourceSheetName = currentSheet.getSheetName();
        String rangeRef = expression;
        int delimiter = expression.lastIndexOf('!');
        if (delimiter >= 0)
        {
            sourceSheetName = normalizeSheetName(expression.substring(0, delimiter));
            rangeRef = expression.substring(delimiter + 1);
        }

        Sheet sourceSheet = workbook.getSheet(sourceSheetName);
        if (sourceSheet == null || StringUtils.isEmpty(rangeRef))
        {
            return new ArrayList<String>();
        }

        String normalizedRangeRef = rangeRef.replace("$", "");
        String[] refs = normalizedRangeRef.split(":");
        if (refs.length == 0 || refs.length > 2)
        {
            return new ArrayList<String>();
        }

        CellReference firstRef;
        CellReference lastRef;
        try
        {
            firstRef = new CellReference(refs[0]);
            lastRef = refs.length == 2 ? new CellReference(refs[1]) : firstRef;
        }
        catch (Exception e)
        {
            return new ArrayList<String>();
        }

        int firstRow = Math.max(Math.min(firstRef.getRow(), lastRef.getRow()), 0);
        int lastRow = Math.max(firstRef.getRow(), lastRef.getRow());
        int firstCol = Math.max(Math.min(firstRef.getCol(), lastRef.getCol()), 0);
        int lastCol = Math.max(firstRef.getCol(), lastRef.getCol());

        long totalCells = (long) (lastRow - firstRow + 1) * (long) (lastCol - firstCol + 1);
        if (totalCells > 5000L)
        {
            return new ArrayList<String>();
        }

        Set<String> dedupe = new LinkedHashSet<String>();
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        {
            Row row = sourceSheet.getRow(rowIndex);
            if (row == null)
            {
                continue;
            }
            for (int columnIndex = firstCol; columnIndex <= lastCol; columnIndex++)
            {
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                String value = resolveDisplayValue(cell, formatter, evaluator);
                if (StringUtils.isNotEmpty(value))
                {
                    dedupe.add(value.trim());
                }
            }
        }
        return new ArrayList<String>(dedupe);
    }

    private String normalizeSheetName(String rawSheetName)
    {
        String sheetName = rawSheetName == null ? "" : rawSheetName.trim();
        if (sheetName.startsWith("'") && sheetName.endsWith("'") && sheetName.length() > 1)
        {
            sheetName = sheetName.substring(1, sheetName.length() - 1).replace("''", "'");
        }
        return sheetName;
    }

    private List<Map<String, Object>> buildSheetImageView(Workbook workbook, Sheet sheet)
    {
        if (!(sheet instanceof XSSFSheet))
        {
            return new ArrayList<Map<String, Object>>();
        }
        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        List<Map<String, Object>> images = new ArrayList<Map<String, Object>>();
        images.addAll(collectAnchoredImages(sheet, xssfSheet));
        images.addAll(collectBackgroundImages(workbook, sheet, xssfSheet));
        return images;
    }

    private List<Map<String, Object>> collectAnchoredImages(Sheet sheet, XSSFSheet xssfSheet)
    {
        List<Map<String, Object>> images = new ArrayList<Map<String, Object>>();
        XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
        if (drawing == null)
        {
            return images;
        }
        for (XSSFShape shape : drawing.getShapes())
        {
            if (!(shape instanceof XSSFPicture))
            {
                continue;
            }
            XSSFPicture picture = (XSSFPicture) shape;
            Map<String, Object> image = buildAnchoredImageView(sheet, picture);
            if (!image.isEmpty())
            {
                images.add(image);
            }
        }
        return images;
    }

    private Map<String, Object> buildAnchoredImageView(Sheet sheet, XSSFPicture picture)
    {
        XSSFClientAnchor anchor = picture.getClientAnchor();
        XSSFPictureData pictureData = picture.getPictureData();
        if (anchor == null || pictureData == null)
        {
            return new LinkedHashMap<String, Object>();
        }

        int row1 = Math.max(anchor.getRow1(), 0);
        int col1 = Math.max(anchor.getCol1(), 0);
        int row2 = Math.max(anchor.getRow2(), row1);
        int col2 = Math.max(anchor.getCol2(), col1);

        int dx1Px = emuToPixel(anchor.getDx1());
        int dx2Px = emuToPixel(anchor.getDx2());
        int dy1Px = emuToPixel(anchor.getDy1());
        int dy2Px = emuToPixel(anchor.getDy2());

        int widthPx = resolveImageWidthPx(sheet, col1, col2, dx1Px, dx2Px);
        int heightPx = resolveImageHeightPx(sheet, row1, row2, dy1Px, dy2Px);
        if (widthPx <= 0 || heightPx <= 0)
        {
            int[] intrinsic = resolveIntrinsicImageSize(pictureData.getData());
            if (widthPx <= 0)
            {
                widthPx = intrinsic[0];
            }
            if (heightPx <= 0)
            {
                heightPx = intrinsic[1];
            }
        }

        Map<String, Object> image = new LinkedHashMap<String, Object>();
        image.put("kind", "anchored");
        image.put("row1", row1);
        image.put("col1", col1);
        image.put("row2", row2);
        image.put("col2", col2);
        image.put("dx1Px", Math.max(dx1Px, 0));
        image.put("dx2Px", Math.max(dx2Px, 0));
        image.put("dy1Px", Math.max(dy1Px, 0));
        image.put("dy2Px", Math.max(dy2Px, 0));
        image.put("widthPx", Math.max(widthPx, 1));
        image.put("heightPx", Math.max(heightPx, 1));
        image.put("src", buildImageDataUrl(pictureData));
        image.put("mimeType", pictureData.getMimeType());
        image.put("extension", pictureData.suggestFileExtension());
        image.put("description", picture.getShapeName());
        return image;
    }

    private List<Map<String, Object>> collectBackgroundImages(Workbook workbook, Sheet sheet, XSSFSheet xssfSheet)
    {
        List<Map<String, Object>> images = new ArrayList<Map<String, Object>>();
        String relationshipId = resolveBackgroundPictureRelationId(xssfSheet);
        if (StringUtils.isEmpty(relationshipId))
        {
            return images;
        }

        POIXMLDocumentPart relationPart = xssfSheet.getRelationById(relationshipId);
        if (!(relationPart instanceof XSSFPictureData))
        {
            return images;
        }

        XSSFPictureData pictureData = (XSSFPictureData) relationPart;
        int[] intrinsic = resolveIntrinsicImageSize(pictureData.getData());
        int fallbackWidth = resolveSheetWidthPx(sheet);
        int fallbackHeight = resolveSheetHeightPx(sheet);

        Map<String, Object> image = new LinkedHashMap<String, Object>();
        image.put("kind", "background");
        image.put("row1", 0);
        image.put("col1", 0);
        image.put("row2", Math.max(resolveLastActiveRowIndex(sheet), 0));
        image.put("col2", Math.max(resolveMaxColumnCount(sheet) - 1, 0));
        image.put("dx1Px", 0);
        image.put("dx2Px", 0);
        image.put("dy1Px", 0);
        image.put("dy2Px", 0);
        image.put("widthPx", Math.max(intrinsic[0], fallbackWidth));
        image.put("heightPx", Math.max(intrinsic[1], fallbackHeight));
        image.put("src", buildImageDataUrl(pictureData));
        image.put("mimeType", pictureData.getMimeType());
        image.put("extension", pictureData.suggestFileExtension());
        image.put("description", "SheetBackground");
        images.add(image);
        return images;
    }

    private String resolveBackgroundPictureRelationId(XSSFSheet xssfSheet)
    {
        if (xssfSheet == null || xssfSheet.getPackagePart() == null)
        {
            return "";
        }
        try (InputStream inputStream = xssfSheet.getPackagePart().getInputStream())
        {
            String worksheetXml = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
            Matcher matcher = BACKGROUND_PICTURE_RELATION_PATTERN.matcher(worksheetXml);
            if (matcher.find())
            {
                return matcher.group(1);
            }
        }
        catch (IOException e)
        {
            log.debug("Failed to resolve sheet background picture relation, sheet: {}", xssfSheet.getSheetName(), e);
        }
        return "";
    }

    private String buildImageDataUrl(XSSFPictureData pictureData)
    {
        String mimeType = StringUtils.isNotEmpty(pictureData.getMimeType()) ? pictureData.getMimeType()
                : MediaType.APPLICATION_OCTET_STREAM_VALUE;
        String base64Data = Base64.getEncoder().encodeToString(pictureData.getData());
        return "data:" + mimeType + ";base64," + base64Data;
    }

    private int emuToPixel(int value)
    {
        return (int) Math.round(value / 9525D);
    }

    private int resolveImageWidthPx(Sheet sheet, int col1, int col2, int dx1Px, int dx2Px)
    {
        if (col2 < col1)
        {
            return 0;
        }
        if (col1 == col2)
        {
            return Math.max(dx2Px - dx1Px, 0);
        }

        int total = Math.max(approximateColumnWidthPx(sheet, col1) - dx1Px, 0);
        for (int columnIndex = col1 + 1; columnIndex < col2; columnIndex++)
        {
            total += approximateColumnWidthPx(sheet, columnIndex);
        }
        total += Math.max(dx2Px, 0);
        return total;
    }

    private int resolveImageHeightPx(Sheet sheet, int row1, int row2, int dy1Px, int dy2Px)
    {
        if (row2 < row1)
        {
            return 0;
        }
        if (row1 == row2)
        {
            return Math.max(dy2Px - dy1Px, 0);
        }

        int total = Math.max(resolveRowHeightPx(sheet, row1) - dy1Px, 0);
        for (int rowIndex = row1 + 1; rowIndex < row2; rowIndex++)
        {
            total += resolveRowHeightPx(sheet, rowIndex);
        }
        total += Math.max(dy2Px, 0);
        return total;
    }

    private int resolveSheetWidthPx(Sheet sheet)
    {
        int maxColumnCount = resolveMaxColumnCount(sheet);
        int total = 0;
        for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
        {
            total += approximateColumnWidthPx(sheet, columnIndex);
        }
        return Math.max(total, 320);
    }

    private int resolveSheetHeightPx(Sheet sheet)
    {
        int lastRowIndex = resolveLastActiveRowIndex(sheet);
        int total = 0;
        for (int rowIndex = 0; rowIndex <= lastRowIndex; rowIndex++)
        {
            total += resolveRowHeightPx(sheet, rowIndex);
        }
        return Math.max(total, 160);
    }

    private int[] resolveIntrinsicImageSize(byte[] data)
    {
        int[] result = new int[] { 240, 160 };
        if (data == null || data.length == 0)
        {
            return result;
        }
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(data))
        {
            BufferedImage image = ImageIO.read(inputStream);
            if (image != null && image.getWidth() > 0 && image.getHeight() > 0)
            {
                result[0] = image.getWidth();
                result[1] = image.getHeight();
            }
        }
        catch (IOException e)
        {
            // Ignore image parse failure and use fallback dimensions.
        }
        return result;
    }

    private Map<String, Object> buildBorderStyle(Workbook workbook, CellStyle cellStyle)
    {
        Map<String, Object> style = new LinkedHashMap<String, Object>();
        String top = resolveBorderCss(workbook, cellStyle, cellStyle.getBorderTop(), "top");
        if (StringUtils.isNotEmpty(top))
        {
            style.put("borderTop", top);
        }

        String right = resolveBorderCss(workbook, cellStyle, cellStyle.getBorderRight(), "right");
        if (StringUtils.isNotEmpty(right))
        {
            style.put("borderRight", right);
        }

        String bottom = resolveBorderCss(workbook, cellStyle, cellStyle.getBorderBottom(), "bottom");
        if (StringUtils.isNotEmpty(bottom))
        {
            style.put("borderBottom", bottom);
        }

        String left = resolveBorderCss(workbook, cellStyle, cellStyle.getBorderLeft(), "left");
        if (StringUtils.isNotEmpty(left))
        {
            style.put("borderLeft", left);
        }

        return style;
    }

    private String resolveBorderCss(Workbook workbook, CellStyle cellStyle, BorderStyle borderStyle, String side)
    {
        if (borderStyle == null || borderStyle == BorderStyle.NONE)
        {
            return "";
        }
        String lineStyle = toBorderCssLineStyle(borderStyle);
        String lineWidth = toBorderCssWidth(borderStyle);
        String color = resolveBorderColor(workbook, cellStyle, side);
        if (StringUtils.isEmpty(color))
        {
            color = "#D9D9D9";
        }
        return lineWidth + " " + lineStyle + " " + color;
    }

    private String resolveBorderColor(Workbook workbook, CellStyle cellStyle, String side)
    {
        if (cellStyle instanceof XSSFCellStyle)
        {
            XSSFCellStyle xssfStyle = (XSSFCellStyle) cellStyle;
            XSSFColor color = null;
            if ("top".equals(side))
            {
                color = xssfStyle.getTopBorderXSSFColor();
            }
            else if ("right".equals(side))
            {
                color = xssfStyle.getRightBorderXSSFColor();
            }
            else if ("bottom".equals(side))
            {
                color = xssfStyle.getBottomBorderXSSFColor();
            }
            else if ("left".equals(side))
            {
                color = xssfStyle.getLeftBorderXSSFColor();
            }
            return toCssColor(color);
        }

        if (cellStyle instanceof HSSFCellStyle && workbook instanceof HSSFWorkbook)
        {
            HSSFCellStyle hssfStyle = (HSSFCellStyle) cellStyle;
            short colorIndex = 0;
            if ("top".equals(side))
            {
                colorIndex = hssfStyle.getTopBorderColor();
            }
            else if ("right".equals(side))
            {
                colorIndex = hssfStyle.getRightBorderColor();
            }
            else if ("bottom".equals(side))
            {
                colorIndex = hssfStyle.getBottomBorderColor();
            }
            else if ("left".equals(side))
            {
                colorIndex = hssfStyle.getLeftBorderColor();
            }
            HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            HSSFColor color = palette == null ? null : palette.getColor(colorIndex);
            return toCssColor(color);
        }
        return "";
    }

    private String toBorderCssLineStyle(BorderStyle borderStyle)
    {
        switch (borderStyle)
        {
            case DOUBLE:
                return "double";
            case DASH_DOT:
            case DASH_DOT_DOT:
            case DOTTED:
                return "dotted";
            case DASHED:
            case MEDIUM_DASH_DOT:
            case MEDIUM_DASH_DOT_DOT:
            case MEDIUM_DASHED:
            case SLANTED_DASH_DOT:
                return "dashed";
            default:
                return "solid";
        }
    }

    private String toBorderCssWidth(BorderStyle borderStyle)
    {
        switch (borderStyle)
        {
            case MEDIUM:
            case MEDIUM_DASH_DOT:
            case MEDIUM_DASH_DOT_DOT:
            case MEDIUM_DASHED:
            case DOUBLE:
                return "2px";
            case THICK:
                return "3px";
            default:
                return "1px";
        }
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

    private String formatCellAddress(int rowIndex, int columnIndex)
    {
        return toColumnLabel(columnIndex) + (rowIndex + 1);
    }

    private String toColumnLabel(int columnIndex)
    {
        int value = Math.max(columnIndex, 0) + 1;
        StringBuilder builder = new StringBuilder();
        while (value > 0)
        {
            int remainder = (value - 1) % 26;
            builder.insert(0, (char) ('A' + remainder));
            value = (value - 1) / 26;
        }
        return builder.length() > 0 ? builder.toString() : "A";
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

    private static class ValidationRule
    {
        private final List<String> options;
        private final boolean allowBlank;

        ValidationRule(List<String> options, boolean allowBlank)
        {
            this.options = options == null ? new ArrayList<String>() : options;
            this.allowBlank = allowBlank;
        }

        List<String> getOptions()
        {
            return options;
        }

        boolean isAllowBlank()
        {
            return allowBlank;
        }
    }

    public static class WorkbookPatchRequest
    {
        private String fileName;

        private String version;

        private List<CellPatch> changes;

        private List<MergeRegionPatch> mergeRegions;

        private List<RowHeightPatch> rowHeights;

        private List<ColumnWidthPatch> columnWidths;

        public String getFileName()
        {
            return fileName;
        }

        public void setFileName(String fileName)
        {
            this.fileName = fileName;
        }

        public String getVersion()
        {
            return version;
        }

        public void setVersion(String version)
        {
            this.version = version;
        }

        public List<CellPatch> getChanges()
        {
            return changes;
        }

        public void setChanges(List<CellPatch> changes)
        {
            this.changes = changes;
        }

        public List<MergeRegionPatch> getMergeRegions()
        {
            return mergeRegions;
        }

        public void setMergeRegions(List<MergeRegionPatch> mergeRegions)
        {
            this.mergeRegions = mergeRegions;
        }

        public List<RowHeightPatch> getRowHeights()
        {
            return rowHeights;
        }

        public void setRowHeights(List<RowHeightPatch> rowHeights)
        {
            this.rowHeights = rowHeights;
        }

        public List<ColumnWidthPatch> getColumnWidths()
        {
            return columnWidths;
        }

        public void setColumnWidths(List<ColumnWidthPatch> columnWidths)
        {
            this.columnWidths = columnWidths;
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

    public static class MergeRegionPatch
    {
        private String sheetName;

        private Integer firstRow;

        private Integer lastRow;

        private Integer firstColumn;

        private Integer lastColumn;

        public String getSheetName()
        {
            return sheetName;
        }

        public void setSheetName(String sheetName)
        {
            this.sheetName = sheetName;
        }

        public Integer getFirstRow()
        {
            return firstRow;
        }

        public void setFirstRow(Integer firstRow)
        {
            this.firstRow = firstRow;
        }

        public Integer getLastRow()
        {
            return lastRow;
        }

        public void setLastRow(Integer lastRow)
        {
            this.lastRow = lastRow;
        }

        public Integer getFirstColumn()
        {
            return firstColumn;
        }

        public void setFirstColumn(Integer firstColumn)
        {
            this.firstColumn = firstColumn;
        }

        public Integer getLastColumn()
        {
            return lastColumn;
        }

        public void setLastColumn(Integer lastColumn)
        {
            this.lastColumn = lastColumn;
        }
    }

    public static class RowHeightPatch
    {
        private String sheetName;

        private Integer rowIndex;

        private Integer heightPx;

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

        public Integer getHeightPx()
        {
            return heightPx;
        }

        public void setHeightPx(Integer heightPx)
        {
            this.heightPx = heightPx;
        }
    }

    public static class ColumnWidthPatch
    {
        private String sheetName;

        private Integer colIndex;

        private Integer widthPx;

        public String getSheetName()
        {
            return sheetName;
        }

        public void setSheetName(String sheetName)
        {
            this.sheetName = sheetName;
        }

        public Integer getColIndex()
        {
            return colIndex;
        }

        public void setColIndex(Integer colIndex)
        {
            this.colIndex = colIndex;
        }

        public Integer getWidthPx()
        {
            return widthPx;
        }

        public void setWidthPx(Integer widthPx)
        {
            this.widthPx = widthPx;
        }
    }

    public static class FileLockRequest
    {
        private String fileName;

        private boolean force;

        public String getFileName()
        {
            return fileName;
        }

        public void setFileName(String fileName)
        {
            this.fileName = fileName;
        }

        public boolean isForce()
        {
            return force;
        }

        public void setForce(boolean force)
        {
            this.force = force;
        }
    }

    private static class EditorIdentity
    {
        private final Long userId;

        private final String username;

        EditorIdentity(Long userId, String username)
        {
            this.userId = userId;
            this.username = username;
        }

        Long getUserId()
        {
            return userId;
        }

        String getUsername()
        {
            return username;
        }
    }

    private static class FileEditLock
    {
        private final String fileName;

        private final Long ownerUserId;

        private final String ownerUsername;

        private final long createdAt;

        private long updatedAt;

        private long expiresAt;

        FileEditLock(String fileName, Long ownerUserId, String ownerUsername, long createdAt, long expiresAt)
        {
            this.fileName = fileName;
            this.ownerUserId = ownerUserId;
            this.ownerUsername = ownerUsername;
            this.createdAt = createdAt;
            this.updatedAt = createdAt;
            this.expiresAt = expiresAt;
        }

        boolean isOwnedBy(EditorIdentity editor)
        {
            if (editor == null)
            {
                return false;
            }
            if (ownerUserId != null && editor.getUserId() != null)
            {
                return ownerUserId.equals(editor.getUserId());
            }
            if (StringUtils.isEmpty(ownerUsername) || StringUtils.isEmpty(editor.getUsername()))
            {
                return false;
            }
            return ownerUsername.equals(editor.getUsername());
        }

        String getFileName()
        {
            return fileName;
        }

        Long getOwnerUserId()
        {
            return ownerUserId;
        }

        String getOwnerUsername()
        {
            return ownerUsername;
        }

        long getCreatedAt()
        {
            return createdAt;
        }

        long getUpdatedAt()
        {
            return updatedAt;
        }

        void setUpdatedAt(long updatedAt)
        {
            this.updatedAt = updatedAt;
        }

        long getExpiresAt()
        {
            return expiresAt;
        }

        void setExpiresAt(long expiresAt)
        {
            this.expiresAt = expiresAt;
        }
    }

    private static class LockOperationResult
    {
        private final boolean success;

        private final FileEditLock lock;

        private final String message;

        private LockOperationResult(boolean success, FileEditLock lock, String message)
        {
            this.success = success;
            this.lock = lock;
            this.message = message;
        }

        static LockOperationResult success(FileEditLock lock)
        {
            return new LockOperationResult(true, lock, "");
        }

        static LockOperationResult failure(FileEditLock lock, String message)
        {
            return new LockOperationResult(false, lock, message);
        }

        boolean isSuccess()
        {
            return success;
        }

        FileEditLock getLock()
        {
            return lock;
        }

        String getMessage()
        {
            return message;
        }
    }

    private static class LockValidationResult
    {
        private final boolean allowed;

        private final FileEditLock lock;

        private final String message;

        private LockValidationResult(boolean allowed, FileEditLock lock, String message)
        {
            this.allowed = allowed;
            this.lock = lock;
            this.message = message;
        }

        static LockValidationResult success(FileEditLock lock)
        {
            return new LockValidationResult(true, lock, "");
        }

        static LockValidationResult failure(FileEditLock lock, String message)
        {
            return new LockValidationResult(false, lock, message);
        }

        boolean isAllowed()
        {
            return allowed;
        }

        FileEditLock getLock()
        {
            return lock;
        }

        String getMessage()
        {
            return message;
        }
    }
}
