package technologies.zip4j;

import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.FileHeader;
import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Test the Zip4J library.
 */
public class Zip4JTests
{
    @Test
    public void readPowerPointSlide() throws IOException
    {
        Path path = Paths.get("../example-template/example-template.pptx");
        ZipFile zipFile = new ZipFile(path.toFile());
        FileHeader fileHeader = zipFile.getFileHeader("ppt/slides/slide1.xml");

        try (var inputStream = zipFile.getInputStream(fileHeader))
        {
            String content = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
            assertTrue(content.contains("Einstein"));
        }
    }
}
