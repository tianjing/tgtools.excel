package tgtools.excel.util;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * @author 田径
 * @Title
 * @Description
 * @date 14:25
 */
public class JsonUtil {
    private static ObjectMapper mObjectMapper =new ObjectMapper();

    public static ObjectNode getEmptyObjectNode()
    {
        return mObjectMapper.createObjectNode();
    }
    public static ArrayNode getEmptyArrayNode()
    {
        return mObjectMapper.createArrayNode();
    }
}
