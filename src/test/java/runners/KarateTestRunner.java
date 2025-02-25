package runners;
import com.intuit.karate.Results;
import com.intuit.karate.junit5.Karate;
import com.intuit.karate.Runner;
import org.junit.jupiter.api.Test;
import utils.ExcelReader;
import java.util.List;
import java.util.Map;

public class KarateTestRunner {

        @Test
         public void  testAll() {
            Results results = Runner.path("classpath:features")
                    .configDir("classpath:karate-config")
                    .karateEnv("dev")
                    //  .systemProperty("testData", testData.toString())
                    .parallel(1);
           // org.junit.jupiter.api.Assertions.assertEquals(0, results.getFailCount(), "There are test failures!");

        }
}

