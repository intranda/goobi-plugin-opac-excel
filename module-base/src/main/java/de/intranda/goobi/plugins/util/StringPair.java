package de.intranda.goobi.plugins.util;

import lombok.AllArgsConstructor;
import lombok.Getter;

public @Getter  @AllArgsConstructor class StringPair {

    String classification;
    String label;

    // TODO
    @Override
    public boolean equals(Object obj) {
        if (obj instanceof StringPair) {
            StringPair sp = (StringPair) obj;
            return sp.getClassification().equals(classification);
        } else if (obj instanceof String) {
            return classification.equals((String) obj);
        }
        return false;
    }

}
