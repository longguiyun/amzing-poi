package com.amazing.poi.excel;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class SimplePOJO {

    private String name;

    private Integer age;

    private String gender;

//    public SimplePOJO() {
//        this.name = "";
//        this.age = 0;
//        this.gender = "";
//    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    @Override
    public String toString() {
        return "SimplePOJO{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", gender='" + gender + '\'' +
                '}';
    }
}
