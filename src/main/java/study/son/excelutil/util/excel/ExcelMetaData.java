package study.son.excelutil.util.excel;

import java.lang.reflect.Field;
import java.util.*;

public class ExcelMetaData<T> {

    private List<Field> fieldNames = new ArrayList<>();
    private Map<String, Field> columns = new LinkedHashMap<>();
    private Class<T> type;

    public ExcelMetaData(Class<T> type) {
        this.type = type;
        Field[] fields = type.getDeclaredFields();
        for (Field field : fields) {
            if (field.getAnnotation(ExcelColumn.class) != null) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                this.fieldNames.add(field);
                this.columns.put(excelColumn.fieldName(), field);
            }
        }
    }

    public T getInstance() throws Exception {
        return (T) this.type.getConstructor().newInstance();
    }

    public List<String> getFieldNames() {
        return new ArrayList<>(this.columns.keySet());
    }

    public Optional<Field> getField(String fieldName) {
        return Optional.ofNullable(this.columns.get(fieldName));
    }

}
