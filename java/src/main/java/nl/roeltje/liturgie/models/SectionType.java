package nl.roeltje.liturgie.models;

import com.fasterxml.jackson.annotation.JsonValue;

/** Type of a {@link LiturgySection}. Mirrors Python's {@code SectionType} enum. */
public enum SectionType {
    REGULAR("regular"),
    SONG("song");

    private final String value;

    SectionType(String value) { this.value = value; }

    @JsonValue
    public String getValue() { return value; }

    public static SectionType fromValue(String v) {
        for (SectionType t : values()) {
            if (t.value.equalsIgnoreCase(v)) return t;
        }
        return REGULAR;
    }
}
