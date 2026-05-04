package org.example.gov.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class WorkbookJson {

   private String fileName;
   private int sheetCount;
   private List<SheetJson> sheetJsonJsons = new ArrayList<>();


}
