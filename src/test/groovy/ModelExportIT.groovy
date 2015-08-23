import io.dbmaster.testng.BaseToolTestNGCase;
import io.dbmaster.testng.OverridePropertyNames;
import static org.testng.Assert.assertTrue;

import org.testng.annotations.Test
import org.testng.annotations.Parameters;

import com.branegy.tools.api.ExportType;


@OverridePropertyNames(project="project.dictionary")
public class ModelExportIT extends BaseToolTestNGCase {
    
    @Test
    @Parameters(["model-export.p_model_name","model-export.p_model_version","model-export.p_filename"])
    public void testModelExport(String p_model_name,String p_model_version, String p_filename) {
        def parameters = [ "p_model_name"  :  p_model_name,
                           "p_model_version" : p_model_version,
                           "p_filename" : p_filename 
                         ]
        String result = tools.toolExecutor("model-export", parameters).execute()
        assertTrue(result.contains("Export completed"), "Unexpected search results ${result}");
    }
}

