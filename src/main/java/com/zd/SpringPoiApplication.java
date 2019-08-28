package com.zd;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class SpringPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(SpringPoiApplication.class, args);
        System.out.println("启动成功！");
    }

}
