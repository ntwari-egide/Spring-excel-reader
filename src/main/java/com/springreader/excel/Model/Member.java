package com.springreader.excel.Model;

public class Member {
    private String memberNo;
    private String firstName;
    private String secondName;
    private String marks;
    private String position;

    public Member(String memberNo, String firstName, String secondName, String marks, String position) {
        this.memberNo = memberNo;
        this.firstName = firstName;
        this.secondName = secondName;
        this.marks = marks;
        this.position = position;
    }

    public String getMemberNo() {
        return memberNo;
    }

    public void setMemberNo(String memberNo) {
        this.memberNo = memberNo;
    }

    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public String getSecondName() {
        return secondName;
    }

    public void setSecondName(String secondName) {
        this.secondName = secondName;
    }

    public String getMarks() {
        return marks;
    }

    public void setMarks(String marks) {
        this.marks = marks;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }
}
