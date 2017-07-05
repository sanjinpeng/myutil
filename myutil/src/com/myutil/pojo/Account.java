package com.myutil.pojo;

import java.util.Date;

public class Account {
	private String name;
	private Integer age;
	private Date creatDate;
	private int count;
	private Double account;
	private double newcount;
	private float xishu;
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
	public Date getCreatDate() {
		return creatDate;
	}
	public void setCreatDate(Date creatDate) {
		this.creatDate = creatDate;
	}
	public int getCount() {
		return count;
	}
	public void setCount(int count) {
		this.count = count;
	}
	public Double getAccount() {
		return account;
	}
	public void setAccount(Double account) {
		this.account = account;
	}
	public double getNewcount() {
		return newcount;
	}
	public void setNewcount(double newcount) {
		this.newcount = newcount;
	}
	public float getXishu() {
		return xishu;
	}
	public void setXishu(float xishu) {
		this.xishu = xishu;
	}
	@Override
	public String toString() {
		return "Account [name=" + name + ", age=" + age + ", creatDate=" + creatDate + ", count=" + count + ", account="
				+ account + ", newcount=" + newcount + ", xishu=" + xishu + "]";
	}
	
	
	
}
