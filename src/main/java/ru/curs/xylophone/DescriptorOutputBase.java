package ru.curs.xylophone;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.WRAPPER_OBJECT)
@JsonSubTypes({
        @JsonSubTypes.Type(value = DescriptorIteration.class, name = "iteration"),
        @JsonSubTypes.Type(value = DescriptorOutput.class, name = "output")})
abstract class DescriptorOutputBase {
}
