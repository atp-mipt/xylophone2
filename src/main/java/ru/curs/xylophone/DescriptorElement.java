package ru.curs.xylophone;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

final class DescriptorElement {
    private final String name;
    private final List<DescriptorOutputBase> sub_elements;

    DescriptorElement(String name) {
        this.name = name;
        this.sub_elements = new LinkedList<>();
    }

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorElement(
            @JsonProperty("name") String name,
            @JsonProperty("output-steps") List<DescriptorOutputBase> sub_elements)
    {
        this.name = name;
        this.sub_elements = sub_elements;
    }

    String getName() {
        return name;
    }

    List<DescriptorOutputBase> getSubElements() {
        return sub_elements;
    }

    public static DescriptorElement jsonDeserialize(InputStream json_stream) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readValue(json_stream,DescriptorElement.class);
    }

}
