package com.sioi.test_excel_in.poi;


public class Person {
    private int id;
    private String name;
    private String sex;
    private int age;
    private String phoneNumber;
    
    public int getId() {
    
        return id;
    }
    
    public void setId(int id) {
    
        this.id = id;
    }
    
    public String getName() {
    
        return name;
    }
    
    public void setName(String name) {
    
        this.name = name;
    }
    
    public String getSex() {
    
        return sex;
    }
    
    public void setSex(String sex) {
    
        this.sex = sex;
    }
    
    public int getAge() {
    
        return age;
    }
    
    public void setAge(int age) {
    
        this.age = age;
    }
    
    public String getPhoneNumber() {
    
        return phoneNumber;
    }
    
    public void setPhoneNumber(String phoneNumber) {
    
        this.phoneNumber = phoneNumber;
    }

    @Override
    public String toString() {

        return "Person [id=" + id + ", name=" + name + ", sex=" + sex + ", age=" + age + ", phoneNumber=" + phoneNumber
                + "]";
    }

    
    
    

}
