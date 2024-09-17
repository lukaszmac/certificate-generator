package technologies.jackson;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Tests the Jackson JSON serialization library.
 */
public class JacksonTests
{
    @Test
    public void jsonCTest() throws JsonProcessingException
    {
        JsonMapper mapper = new JsonMapper();
        mapper.enable(JsonParser.Feature.ALLOW_COMMENTS);
        var data = mapper.readValue("//Hello World - I have a comment\n { \"hello\":\"world\"}", LinkedHashMap.class);
        assertEquals("world", data.get("hello"));
    }

    @Test
    public void JsonNodeTest()
    {
        JsonNodeFactory factory = new JsonNodeFactory(true);

        var root = factory.objectNode();
        root.set("Hello", factory.textNode("World"));

        var expected = "{\"Hello\":\"World\"}";
        assertEquals(expected, root.toString());
    }
}
