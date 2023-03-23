package com.opendynamic;

import java.util.List;
import java.util.Map;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

public class OdResult {
    private Map<String, Object> result;

    @SuppressWarnings("unchecked")
    public OdResult(String data) {
        Gson gson = new GsonBuilder().create();
        result = gson.fromJson(data, Map.class);
    }

    public Map<String, Object> getResult() {
        return result;
    }

    public boolean getSuccess() {
        return (Boolean) result.get("success");
    }

    public int getTotal() {
        return ((Double) result.get("total")).intValue();
    }

    @SuppressWarnings("unchecked")
    public Map<String, Object> getObject(String key) {
        return (Map<String, Object>) result.get(key);
    }

    @SuppressWarnings("unchecked")
    public List<Map<String, Object>> getObjectList(String key) {
        return (List<Map<String, Object>>) result.get(key);
    }
}