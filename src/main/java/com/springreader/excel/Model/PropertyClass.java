package com.springreader.excel.Model;

import java.util.Iterator;
import java.util.Properties;

public class PropertyClass {
    public void usageOfProperties(){
        Properties properties = new Properties();
        properties.setProperty("email","ntwariegide2@gmail.com");
        properties.setProperty("username","egide");
        properties.setProperty("password","123");

        // in order to get the properties
        Iterator keysIterator = properties.keySet().iterator();
        while (keysIterator.hasNext()){
            String email = (String) keysIterator.next();
            String username = (String) keysIterator.next();

            System.out.println(email);
            System.out.println(username);
        }
        System.out.println(properties.getProperty("password"));
    }
}
