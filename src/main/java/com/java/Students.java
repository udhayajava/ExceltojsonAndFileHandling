package com.java;
public class Students {
        private int rollNum;
        private String name;
        private int age;
        private int mark;
    public Students() {
    }

    public Students( int rollNum, String name,int age, int mark){
        this.rollNum = rollNum;
        this.name = name;
        this.age = age;
        this.mark = mark;
    }

        public int getRollNum () {
        return rollNum;
    }

        public void setRollNum ( int rollNum){
        this.rollNum = rollNum;
    }

        public String getName () {
        return name;
    }

        public void setName (String name){
        this.name = name;
    }

        public int getAge () {
        return age;
    }

        public void setAge ( int age){
        this.age = age;
    }

        public int getMark () {
        return mark;
    }

        public void setMark ( int mark){
        this.mark = mark;
    }
        @Override
        public String toString () {
        return "Students{" +
                "rollNum=" + rollNum +
                ", name='" + name + '\'' +
                ", age=" + age +
                ", mark=" + mark +
                '}';
    }

    }
