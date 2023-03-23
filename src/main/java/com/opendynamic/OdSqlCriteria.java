package com.opendynamic;

import java.util.Map;

public class OdSqlCriteria {
    private String sql;
    private Map<String, Object> paramMap;

    public OdSqlCriteria(String sql, Map<String, Object> paramMap) {
        this.sql = sql;
        this.paramMap = paramMap;
    }

    public String getSql() {
        return sql;
    }

    public Map<String, Object> getParamMap() {
        return paramMap;
    }
}