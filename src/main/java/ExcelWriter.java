import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;


public class ExcelWriter {

    private static final String RECEIPT_TEMPLATE_ORIG_XLSX_FILE_PATH = "./receipt_template_original.xlsx";
    private static final String RECEIPT_TEMPLATE_XLSX_FILE_PATH = "./receipt_template.xlsx";
    private static final String RECEIPT_OUTPUT_XLSX_FILE_PATH = "./receipt_output.xlsx";

    // for macro, xlsm file
    //    private static final String RECEIPT_TEMPLATE_ORIG_XLSX_FILE_PATH = "./receipt_template_macro_original.xlsm";
    //    private static final String RECEIPT_TEMPLATE_XLSX_FILE_PATH = "./receipt_template_macro.xlsm";
    //    private static final String RECEIPT_OUTPUT_XLSX_FILE_PATH = "./receipt_output_macro.xlsm";

    private static final Receipt sample = new Receipt(
            "10002-0010005",
            "2016年7月23日",
            "株式会社今北産業",
            "再発行",
            "\\174,000-",
            "ショッピングオンライン（仮称）",
            "00001169010002",
            "1100000000001031,1100000000001161\n1100000000001162,1100000000001200",
            "2016年7月20日",
            "○○○○○○○○○○○○○○○○○○○○\n○○○○○○○○○○○○○○○○○○○○",
            "株式会社ホールディングス○○○○\n○○○○○○○○○○○○○○○○○○○○",
            "ショッピングパークオンラインECショップ",
            "ショッピングパークカスタマーセンター",
            "TEL：03-1111-2222");

    public static void main() throws IOException, InvalidFormatException {
        modify(new File(RECEIPT_TEMPLATE_ORIG_XLSX_FILE_PATH), new File(RECEIPT_OUTPUT_XLSX_FILE_PATH), sample);
    }

    // Example to modify an existing excel file
    private static void modify(File originalTemplate, File output, Object bean) throws InvalidFormatException, IOException {

        Path templatePath = java.nio.file.Files.copy(FileSystems.getDefault().getPath(originalTemplate.getPath()), FileSystems.getDefault().getPath(RECEIPT_TEMPLATE_XLSX_FILE_PATH));
        Path outputPath = java.nio.file.Files.copy(FileSystems.getDefault().getPath(originalTemplate.getPath()), FileSystems.getDefault().getPath(output.getPath()));

        Map<Field, String> fieldMethodMap = getMap(bean);

        // Obtain a workbook from the excel file
        final Workbook workbook = WorkbookFactory.create(templatePath.toFile());
        try {
            // Get Sheet at index 0
            Sheet sheet = workbook.getSheetAt(0);

            fieldMethodMap.keySet().forEach(field -> {
                String fieldName = field.getName();
                if (field.getAnnotation(Direction.class) != null) {
                    fieldName = field.getAnnotation(Direction.class).column();
                }
                findCell(sheet, "${" + fieldName + "}").ifPresent(cell -> {
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(fieldMethodMap.get(field));
                });
            });

            // Update the cell's value
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(outputPath.toFile());
            workbook.write(fileOut);
            fileOut.close();

        } finally {
            // Closing the workbook
            workbook.close();
        }
    }

    private static Map<Field, String> getMap(Object bean) { // fieldName, value
        final Map<Field, String> map = new HashMap<>();
        Stream.of(bean.getClass().getDeclaredFields()).forEach(field -> Stream.of(bean.getClass().getMethods()).filter(method -> {
            return method.getName().equalsIgnoreCase("get" + field.getName().toLowerCase());
        }).findFirst().ifPresent(method -> {
            try {
                Object o = method.invoke(bean);
                if (o != null) {
                    map.put(field, o.toString());
                }
            } catch (IllegalAccessException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }));
        return map;
    }

    private static Optional<Cell> findCell(Sheet sheet, String cellContent) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellTypeEnum() == CellType.STRING && cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                    return Optional.of(cell);
                }
            }
        }
        return Optional.empty();
    }
}
