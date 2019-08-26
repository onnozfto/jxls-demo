package com.self.jxlsdemo.pojo;

import java.util.ArrayList;
import java.util.List;

/**
 * @Description
 * @Author will
 * @Date 2019/08/26
 */
public class Department {

  private String name;


  private Person chief = new Person();//必须创建

  private List<Person> staff = new ArrayList<>();//必须创建

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public Person getChief() {
    return chief;
  }

  public void setChief(Person chief) {
    this.chief = chief;
  }

  public List<Person> getStaff() {
    return staff;
  }

  public void setStaff(List<Person> staff) {
    this.staff = staff;
  }
}
