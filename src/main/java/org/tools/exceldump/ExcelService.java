package org.tools.exceldump;

import java.io.File;
import java.io.FileNotFoundException;
import java.nio.file.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.List;

import javax.annotation.PostConstruct;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.jdbc.support.rowset.SqlRowSet;
import org.springframework.jdbc.support.rowset.SqlRowSetMetaData;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {
	
	@Autowired
    private ApplicationArguments args;
	
	@Autowired
	private NamedParameterJdbcTemplate namedjdbctemplate;
	
	private MapSqlParameterSource paramsource;
	
	private void dumpquery(String query) {
		System.out.println("running query:" + query);
		String outputloc = "dump.xlsx";
		paramsource = new MapSqlParameterSource();
		SqlRowSet toret = namedjdbctemplate.queryForRowSet(query, paramsource);
		SqlRowSetMetaData rsmd =  toret.getMetaData();
		System.out.println(rsmd.getColumnCount());		
		
		SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
		Sheet sheet = wb.createSheet();
	       
		Row headerRow = sheet.createRow(0);
		for(int i=1; i<=rsmd.getColumnCount();i++) {
			Cell cell = headerRow.createCell(i-1);
			cell.setCellValue(rsmd.getColumnName(i));
		}
		
		Integer rowNum = 1;
		while(toret.next()) {
			Row row = sheet.createRow(rowNum);
			HashMap<String,Object> datarow = new HashMap<String,Object>();
			for(int i=1; i<=rsmd.getColumnCount();i++) {				
				row.createCell(i-1).setCellValue(toret.getString(i));
			}
			rowNum+=1;
		}
		
		try {
			FileOutputStream fileOut = new FileOutputStream(outputloc);
			wb.write(fileOut);
			fileOut.close();			
			wb.dispose();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static String readFileAsString(String fileName) throws Exception  {
	    String data = "";
	    data = new String(Files.readAllBytes(Paths.get(fileName)));
	    return data;
	}
	
	@PostConstruct
	public void init() {
		
		String sqlcmd="select 1";
		boolean gotcmd = false;
		
		if(args.containsOption("sqlfile")) 
        {
            //Get argument values
            List<String> values = args.getOptionValues("sqlfile");
            System.out.println("sqlfile:" + values);
            File tempFile = new File(values.get(0));            
            boolean exists = tempFile.exists();
            if(exists) {
	            try {	            	
					sqlcmd = readFileAsString(values.get(0));
					System.out.println("Using sql from file");
					gotcmd = true;
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}           
            }
        }
		if(!gotcmd && args.containsOption("sqlcmd")) 
        {
            //Get argument values
            List<String> values = args.getOptionValues("sqlcmd");
            System.out.println("sqlcmd:" + values);
            try {
				sqlcmd = values.get(0);
				System.out.println("Using sql from cmdline");
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}  
        }
		
		dumpquery(sqlcmd);
		System.out.println("Here iam ");
	}

}
