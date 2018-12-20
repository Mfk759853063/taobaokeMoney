package com.vbn.taobaokeMoney;

import com.vbn.taobaokeMoney.excelOperation.OperationExcel;
import com.vbn.taobaokeMoney.excelOperation.PoiUtils;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

//@SpringBootApplication
public class TaobaokeMoneyApplication {

	public static void main(String[] args) {
		String path = args[0];
		OperationExcel.readExcel(path);
//		SpringApplication.run(TaobaokeMoneyApplication.class, args);
	}

}

