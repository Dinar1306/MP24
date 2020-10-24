package ru.rkb2ufa;

public class Contact {
    public Contact(String name, int contact, String city) {
        this.name = name;
        this.contact = contact;
        this.city = city;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setContact(int contact) {
        this.contact = contact;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getName() {
        return name;
    }

    public int getContact() {
        return contact;
    }

    public String getCity() {
        return city;
    }

    String name;
    int contact;
    String city;
}
