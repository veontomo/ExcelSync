package com.company;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * Updates a workbook with data from other workbooks
 */
public class XUpdater {

    private final HashMap<String, Integer> target;
    private final List<HashMap<String, Integer>> sources;


    /**
     * list of keys that are present in both target and sources
     */

    private HashMap<String, Integer> duplicates;
    /**
     * list of keys that are present in target and not present in any of the sources
     */
    private List<String> missing;

    /**
     * list of keys that are present in one of the sources and not present in the target
     */
    private HashMap<String, Integer> extra;


    public XUpdater(final HashMap<String, Integer> target, final List<HashMap<String, Integer>> sources) {
        this.target = target;
        this.sources = sources;
    }


    /**
     * Finds the sources that contain a key from the target.
     * <p>
     * If a key is found in multiple sources, an exception is thrown.
     * <p>
     * Returns a hash map from a string to an integer that is the ordinal number of the source in the source list in which
     * the key is found.
     *
     * @return
     */
    public void analyze() throws Exception {
        duplicates = new HashMap<>();
        missing = new ArrayList<>();
        extra = new HashMap<>();
        boolean isFoundInSources;
        int sourcesLen = sources.size();
        // first pass: iterate over the target and control the presence in the sources
        for (String key : target.keySet()) {
            isFoundInSources = false;
            for (int i = 0; i < sourcesLen; i++) {
                if (sources.get(i).containsKey(key)) {
                    if (duplicates.containsKey(key)) {
                        throw new Exception("key " + key + " has already been found in source n. " + i + ". Resolve to proceed.");
                    }
                    isFoundInSources = true;
                    duplicates.put(key, i);
                }
            }
            if (!isFoundInSources) {
                missing.add(key);
            }
        }
        // second pass: iterate over the sources and control if they contain keys that are not in the target
        for (int i = 0; i < sourcesLen; i++) {
            for (String key : sources.get(i).keySet()){
                if (target.containsKey(key)){
                    // cross check: the variable "duplicates" must contain this key as well.
                    if (duplicates.containsKey(key) && duplicates.get(key) == i){
//                        System.out.println("cross-check is OK");
                    } else {
                        System.out.println("cross-check is not OK for key " + key + " that is supposed to be in set " + i);
                    }
                } else {
                    extra.put(key, i);
                }
            }
        }

    }

    public HashMap<String, Integer> getDuplicates() {
        return duplicates;
    }

    public List<String> getMissing() {
        return missing;
    }

    public HashMap<String, Integer> getExtra() {
        return extra;
    }
}
