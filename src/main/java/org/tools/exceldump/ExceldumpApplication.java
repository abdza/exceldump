package org.tools.exceldump;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExceldumpApplication {
	
	public static void main(String[] args) {
		System.out.println("Start");
		for(String arg:args) {
			System.out.println("argss:" + arg);
		}
		System.out.println("end");
		SpringApplication.run(ExceldumpApplication.class, args);		
	}

}
