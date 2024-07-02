package org.example;

import java.util.HashSet;
import java.util.Set;

public class Orders {
    String orderNo;
    String Email;
    int quantity;
    String product;
    Double price;
    Double totalPrice = 0.0;
    String sku;
    String Name;
    String Address;
    String Country;
    String phone;
    Set<String> notes = new HashSet<>();
    String cancelledAt;
    String disCountCode;
    String fulfillmentStatus;
    String fulfillmentAt;
    String fulfillmentNote;
    boolean isMerged = false;

    // getters 和 setters

    public String getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(String customerId) {
        this.orderNo = customerId;
    }

    public String getEmail() {
        return Email;
    }

    public void setEmail(String email) {
        Email = email;
    }

    public int getQuantity() {
        return quantity;
    }

    public void setQuantity(int quantity) {
        this.quantity = quantity;
    }

    public String getProduct() {
        return product;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public Double getPrice() {
        return price;
    }

    public void setPrice(Double price) {
        this.price = price;
    }

    public Double getTotalPrice() {
        return totalPrice;
    }

    public void setTotalPrice(Double totalPrice) {
        this.totalPrice = totalPrice;
    }

    public String getSku() {
        return sku;
    }

    public void setSku(String sku) {
        this.sku = sku;
    }

    public String getName() {
        return Name;
    }

    public void setName(String name) {
        Name = name;
    }

    public String getAddress() {
        return Address;
    }

    public void setAddress(String address) {
        Address = address;
    }

    public String getCountry() {
        return Country;
    }

    public void setCountry(String country) {
        Country = country;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public Set<String> getNotes() {
        return notes;
    }

    public void setNotes(Set<String> notes) {
        this.notes = notes;
    }

    public String getCancelledAt() {
        return cancelledAt;
    }

    public void setCancelledAt(String cancelledAt) {
        this.cancelledAt = cancelledAt;
    }

    public String getDisCountCode() {
        return disCountCode;
    }

    public void setDisCountCode(String disCountCode) {
        this.disCountCode = disCountCode;
    }

    public String getFulfillmentStatus() {
        return fulfillmentStatus;
    }

    public void setFulfillmentStatus(String fulfillmentStatus) {
        this.fulfillmentStatus = fulfillmentStatus;
    }

    public String getFulfillmentAt() {
        return fulfillmentAt;
    }

    public void setFulfillmentAt(String fulfillmentAt) {
        this.fulfillmentAt = fulfillmentAt;
    }

    public String getFulfillmentNote() {
        return fulfillmentNote;
    }

    public String setFulfillmentNote(String fulfillmentNote) {
        this.fulfillmentNote = fulfillmentNote;
        return fulfillmentNote;
    }

    public boolean isMerged() {
        return isMerged;
    }

    public void setMerged(boolean merged) {
        isMerged = merged;
    }
}
