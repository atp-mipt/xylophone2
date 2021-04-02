package ru.curs.xylophone;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.LinkedList;
import java.util.List;

final class DescriptorIteration extends DescriptorOutputBase {
    private final int index;
    private final int merge;
    private final boolean horizontal;
    private final String regionName;
    private final List<DescriptorElement> elements;

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorIteration(
            @JsonProperty("index")      int index,
            @JsonProperty("mode")       String mode,
            @JsonProperty("merge")      int merge,
            @JsonProperty("regionName") String regionName,
            @JsonProperty("element")    List<DescriptorElement> elements)
    {
        this.index = index;
        this.merge = merge;
        this.horizontal = "horizontal".equals(mode);
        this.regionName = regionName;
        this.elements = elements;
    }

    DescriptorIteration(int index, boolean horizontal, int merge,
                        String regionName) {
        this.index = index;
        this.horizontal = horizontal;
        this.merge = merge;
        this.regionName = regionName;
        this.elements = new LinkedList<>();
    }

    int getIndex() {
        return index;
    }

    boolean isHorizontal() {
        return horizontal;
    }

    List<DescriptorElement> getElements() {
        return elements;
    }

    public int getMerge() {
        return merge;
    }

    public String getRegionName() {
        return regionName;
    }
}
