package com.joyenglish.reportgenerator;

import java.util.List;

public class ProjectReportRequest {
    private String name;
    private List<String> items;

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
    
    public List<String> getItems() { return items; }
    public void setItems(List<String> items) { this.items = items; }
}