package org.example.gov.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;
import java.util.Map;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class SheetJson {

   private String sheetName;
   private int headerRowIndex;
   private List<String> columns;
   private List<Map<String, Object>> mapList;
}
