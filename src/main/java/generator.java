import org.apache.commons.text.StringSubstitutor;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collector;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class generator {

    public static void invokeSpringToExcel() throws IOException {
        FileInputStream file = new FileInputStream(new File("/Users/nattapat/Downloads/TEST1.xlsx"));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        data.get(new Integer(i)).add(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i).add(cell.getDateCellValue() + "");
                        } else {
                            data.get(i).add(cell.getNumericCellValue() + "");
                        }
                        break;
                    case BOOLEAN:
                        data.get(i).add(cell.getBooleanCellValue() + "");
                        break;
                    case FORMULA:
                        data.get(i).add(cell.getCellFormula() + "");
                        break;
                    default:
                        data.get(new Integer(i)).add(" ");
                }
            }
            i++;
        }

        String importText = "import com.fasterxml.jackson.annotation.JsonInclude;\n" +
                "import com.fasterxml.jackson.annotation.JsonProperty;\n" +
                "import com.fasterxml.jackson.annotation.JsonPropertyOrder;\n" +
                "import io.swagger.annotations.ApiModel;\n" +
                "import io.swagger.annotations.ApiModelProperty;\n" +
                "import javax.validation.constraints.NotNull;\n \n"+
                "public class accountLoan {";
//    public String getCardId() {
//    return cardId;
//    }
//
//    }
//    public void setCardId(String cardId) {
//        this.cardId = cardId;

        BufferedWriter writer = new BufferedWriter(new FileWriter("/Users/nattapat/Downloads/TEST3.txt"));
        Map<Integer, List<String>> filtered = data.entrySet().stream().filter(x->x.getKey()>=3).collect(Collectors.toMap(a->a.getKey(), a->a.getValue()));
        List<String> combinedFieldList = new ArrayList<>();
        List<String> combinedGetSetList = new ArrayList<>();
        String combinedString ="";
        String combinedGetSet ="";
        filtered.forEach((x,y) -> combinedFieldList.add("private "+firstCharToUpperCase(y.get(2))+' '+y.get(1)+';'));
        filtered.forEach((x,y) -> combinedGetSetList.add(String.format("public void get%s(){%n return %s; %n } %n public %s set%s"
                ,firstCharToUpperCase(y.get(1))
                ,y.get(1)
                ,firstCharToUpperCase(y.get(2))
                ,firstCharToUpperCase(y.get(1))
                )));
        System.out.println(combinedFieldList.size());
        for(int j = 0;j<combinedFieldList.size();j++){
            combinedString=combinedString+"\n"+combinedFieldList.get(j);
            combinedGetSet+="\n"+combinedGetSetList.get(j);
        }
        //get and set





        String combinedModel = importText+combinedString+combinedGetSet+"\n}";
        System.out.println(combinedModel);
        writer.write(combinedModel);
        writer.close();

    }
    public static void main(String[] args) throws IOException {
        invokeSpringToExcel();
    }
    public static String firstCharToUpperCase(String str){
        String cap = str.substring(0, 1).toUpperCase() + str.substring(1);
        return cap;
    }
}