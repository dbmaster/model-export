import com.branegy.dbmaster.model.Table
import com.branegy.dbmaster.model.View
import com.branegy.dbmaster.model.Column
import org.apache.poi.ss.usermodel.Sheet
import com.branegy.dbmaster.model.ModelObject
import com.branegy.dbmaster.database.api.ModelService
import com.branegy.dbmaster.service.ModelExporter.Statistics
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import com.branegy.dbmaster.custom.field.server.api.ICustomFieldService
import com.branegy.dbmaster.service.ModelExporter

logger.info("Model=${p_model_name} version=${p_model_version}")

logger.info("FileName=${p_filename}")


ModelExporter helper = new ModelExporter(null, null)

def workBook = new SXSSFWorkbook(100);
def fields = new java.util.ArrayList(dbm.getService(ICustomFieldService.class).getProjectCustomConfigList())
def sheet = workBook.createSheet("objects") //getSheet
def sheetColumn = workBook.createSheet("columns") //getSheet


def statisticMap = [:]

  def exportColumns(sheet, helper, statisticMap, rowNumber, object, columns) {
        for (column in columns) {
            row = sheet.getRow(rowNumber);
            if (row==null){
                row = sheet.createRow(rowNumber);
            }

            String subjectArea = helper.getSubjectArea(object);
            String description = column.getCustomData("Description");

            def statistics = helper.getStatistics(statisticMap, subjectArea);
            if (object instanceof Table) {
                statistics.countTableColumn++;
                if (description!=null && description.length()>0) {
                    statistics.countTableColumnDescribed++;
                }
            } else if (object instanceof View) {
                statistics.countViewColumn++;
                if (description!=null && description.length()>0) {
                    statistics.countViewColumnDescribed++;
                }
            }
            int j=0;
            helper.setValue(row, j++, object.getName());
            helper.setValue(row, j++, column.getName());
            helper.setValue(row, j++, column.getPrettyType());
            helper.setValue(row, j++, column.isNullable());
            helper.setValue(row, j++, column.getDefaultValue());
            helper.setValue(row, j++, column.getExtraDefinition());
            helper.setValue(row, j++, column.getCustomData("Logical Name"));
            helper.setValue(row, j++, column.getCustomData("Constraints"));
            helper.setValue(row, j++, column.getCustomData("Data Domain"));
            helper.setValue(row, j++, column.getCustomData("Sensitive Data"));
            helper.setValue(row, j++, description);
            helper.setValue(row, j++, column.getCustomData("Notes"));
            helper.setValue(row, j++, column.getCustomData("TODO Items"));
            helper.setValue(row, j++, column.getCustomData("Used By"));

            rowNumber++;    
        }
        return rowNumber;
    }






helper.setHeader(sheet, null, fields, "Type", "Subject Area", "Object Name", 
                "Logical Name", "Source of Records", "Used By", "Author",
                "Description", "Notes", "TODO Items");

helper.setHeader(sheetColumn, null, fields, "Table/View Name", "Column Name", 
                "Type", "Nullable", "Default Value", "Extra Definition",
                "Logical Name","Constraints","Data Domain",  
                "Sensitive Data","Description", "Notes", "TODO Items");

def i=1;
def z=1;
def current = 0;


def modelService = dbm.getService(ModelService.class)
// def modelList = modelService.getModelList(null, null)
def model = modelService.findModelByName(p_model_name, p_model_version, com.branegy.dbmaster.model.Model.FETCH_TREE)

if (model==null) {
  logger.error("Model ${p_model} was not found")
  return
}

def total=model.getTables().size()+model.getViews().size()+model.getProcedures().size()+1;

// step 1.1: Export tables
helper.setProgressMsg("Export tables");
for (table in model.getTables()) {
    def row = sheet.getRow(i);
    if (row==null) {
    	row = sheet.createRow(i);
    }

    String subjectArea = helper.getSubjectArea(table);
    statistics = helper.getStatistics(statisticMap, subjectArea);
    Object description = table.getCustomData("Description");
    if (description!=null && description.toString().length()>0) {
        statistics.countTablesDescribed++;
    }

    statistics.countTables++;

    int j=0;
    helper.setValue(row, j++, "Table");
    helper.setValue(row, j++, subjectArea);
    helper.setValue(row, j++, table.getName());
    helper.setValue(row, j++, table.getCustomData("Logical Name"));
    helper.setValue(row, j++, table.getCustomData("Source of Records"));
    helper.setValue(row, j++, table.getCustomData("Used By"));
    helper.setValue(row, j++, table.getCustomData("Author"));
    helper.setValue(row, j++, description);
    helper.setValue(row, j++, table.getCustomData("Notes"));
    helper.setValue(row, j++, table.getCustomData("TODO Items"));
    i++;

    z = exportColumns(sheetColumn, helper, statisticMap, z, table, table.getColumns());
    helper.setProgressDone(++current/total); 
    if (helper.isCanceled()){ 
	return;
    }
}

// step 1.2: Export views
helper.setProgressMsg("Export views");
for (view in model.getViews()) {
    row = sheet.getRow(i);
    if (row==null){
        row = sheet.createRow(i);
    }
    String subjectArea = helper.getSubjectArea(view);
    statistics = helper.getStatistics(statisticMap, subjectArea);
    statistics.countViews++;
    Object description = view.getCustomData("Description");
    if (description!=null && description.toString().length()>0) {
        statistics.countViewDescribed++;
    }

    int j=0;
    helper.setValue(row, j++, "View");
    helper.setValue(row, j++, subjectArea);
    helper.setValue(row, j++, view.getName());
    helper.setValue(row, j++, view.getCustomData("Logical Name"));
    helper.setValue(row, j++, view.getCustomData("Source of Records"));
    helper.setValue(row, j++, view.getCustomData("Used By"));
    helper.setValue(row, j++, view.getCustomData("Author"));
    helper.setValue(row, j++, description);
    helper.setValue(row, j++, view.getCustomData("Notes"));
    helper.setValue(row, j++, view.getCustomData("TODO Items"));

    i++;

    z = exportColumns(sheetColumn, statisticMap, z, view, view.getColumns());
    helper.setProgressDone(++current/total);
    if (helper.isCanceled()){ 
	return;
    }
}

// step 1.3: Export procedures
helper.setProgressMsg("Export procedures");
for (procedure in model.getProcedures()) {
    row = sheet.getRow(i);
    if (row==null){
        row = sheet.createRow(i);
    }
    String subjectArea = helper.getSubjectArea(procedure);
    statistics = helper.getStatistics(statisticMap, subjectArea);
    statistics.countProcedures++;
    Object description = procedure.getCustomData("Description");
    if (description!=null && description.toString().length()>0) {
        statistics.countProceduresDescribed++;
    }
    int j=0;
    helper.setValue(row, j++, "SP");
    helper.setValue(row, j++, subjectArea);
    helper.setValue(row, j++, procedure.getName());
    helper.setValue(row, j++, procedure.getCustomData("Logical Name"));
    helper.setValue(row, j++, procedure.getCustomData("Source of Records"));
    helper.setValue(row, j++, procedure.getCustomData("Used By"));
    helper.setValue(row, j++, procedure.getCustomData("Author"));
    helper.setValue(row, j++, description);
    helper.setValue(row, j++, procedure.getCustomData("Notes"));
    helper.setValue(row, j++, procedure.getCustomData("TODO Items"));

    i++;
    helper.setProgressDone(++current/total);
    if (helper.isCanceled()){ 
	return;
    }
}


for (columnIndex in 0..6) sheet.autoSizeColumn(columnIndex)

// step 2 Export Statistics

sheet = workBook.createSheet("statistics"); //getSheet

helper.setHeader(sheet, null, null,
   "Subject Area","Tables","Tables\nDescribed","Table\nColumns",
   "Table\nColumns\nDescribed","Views","Views\nDescribed","View\nColumns",
   "View\nColumns\nDescribed","Procedures","Procedures\nDescribed","Progress","%");

row = sheet.getRow(0);

   //to enable newlines you need set a cell styles with wrap=true

    cs = workBook.createCellStyle();
    cs.setWrapText(true);
    cs.setVerticalAlignment(CellStyle.VERTICAL_TOP);
    cs.setFillBackgroundColor(HSSFColor.GREEN.index);
//    cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

    for (columnIndex in 0..12) row.getCell(columnIndex).setCellStyle(cs)

    // cell.setCellStyle(cs);

    //increase row height to accomodate two lines of text
    row.setHeightInPoints((float)3.0*sheet.getDefaultRowHeightInPoints());

i=1;
helper.setProgressMsg("Export statistics");

    font = workBook.createFont();
    font.setFontHeightInPoints((short)10);
    font.setFontName("Arial Unicode MS");
    font.setColor(HSSFColor.GREEN.index);

    // Fonts are set into a style so create a new one to use.
    style = workBook.createCellStyle();
    style.setFont(font);

    style2 = workBook.createCellStyle();
    style2.setDataFormat(workBook.createDataFormat().getFormat("0%"));

for (statistics in statisticMap.values()) {
    row = sheet.getRow(i);
    if (row==null){
        row = sheet.createRow(i);
    }
    int j=0;
    helper.setValue(row, j++, statistics.subjectArea);
    helper.setValue(row, j++, statistics.countTables);
    helper.setValue(row, j++, statistics.countTablesDescribed);
    helper.setValue(row, j++, statistics.countTableColumn);
    helper.setValue(row, j++, statistics.countTableColumnDescribed);
    helper.setValue(row, j++, statistics.countViews);
    helper.setValue(row, j++, statistics.countViewDescribed);
    helper.setValue(row, j++, statistics.countViewColumn);
    helper.setValue(row, j++, statistics.countViewColumnDescribed);
    helper.setValue(row, j++, statistics.countProcedures);
    helper.setValue(row, j++, statistics.countProceduresDescribed);



    helper.setFormula(row, j++, "REPT(\"|\",M"+(i+1)+"*30)")
    row.getCell(j-1).setCellStyle(style)

    u = i+1
    helper.setFormula(row, j++, "=IF(B"+u+"+D"+u+"+F"+u+"+H"+u+"+J"+u+"=0,1,(C"+u+"+E"+u+"+G"+u+"+I"+u+"+K"+u+")/(B"+u+"+D"+u+"+F"+u+"+H"+u+"+J"+u+"))");
    row.getCell(j-1).setCellStyle(style2)
    
    //            setValue(row, j++, statistics.countParameters);
    //            setValue(row, j++, statistics.countParametersDescribed);
    i++;
    if (helper.isCanceled()){ 
	return;
    }	
}

for (columnIndex in 0..10) sheet.autoSizeColumn(columnIndex)
sheet.setColumnWidth(11,16*256)
sheet.setColumnWidth(12,10*256)

// step 4 Export Field Description

   sheet = workBook.createSheet("field_description"); //getSheet
   i=1;
   helper.setProgressMsg("Exporting fields");
   helper.setHeader(sheet, null, null,"Object", "Custom Field Name", "Custom Field Description");

   cs = workBook.createCellStyle();
   cs.setWrapText(true);

   fields.sort([compare:{a,b -> a.clazz.equals(b.clazz) ? a.name.compareTo(b.name) : a.clazz.compareTo(b.clazz) }] as Comparator)

   for (field in fields) {
      row = sheet.getRow(i);
      if (row==null) { row = sheet.createRow(i); }
      int j=0;
      helper.setValue(row, j++, field.clazz)
      helper.setValue(row, j++, field.name)
      helper.setValue(row, j++, field.description)
      i++;
   }
   sheet.autoSizeColumn(0)
   sheet.autoSizeColumn(1)
   sheet.setColumnWidth(2,40*256)

  def fileService = dbm.getService(com.branegy.files.FileService)
  def file = fileService.createFile(p_filename, "model-export")  
  def outputStream = file.getOutputStream()

  workBook.write(outputStream)
  outputStream.close()
  helper.setProgressDone(1)

  println "Export completed. Download file ${p_filename} from the 'Files' tab"