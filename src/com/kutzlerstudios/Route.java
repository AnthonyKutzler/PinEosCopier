package com.kutzlerstudios;

public class Route {

    private String name;
    private String route;
    private String town;
    private int pkgCount;

    public Route(String name, String route){
        this.name = name;
        this.route = route.trim();
    }

    public Route(String name, String route, int pkgCount, String address){
        this.name = name;
        this.route = route.trim();
        this.pkgCount = pkgCount;
        this.town = address.split(",")[1].trim();
    }

    public Route(String route, int pkgCount){
        this.route = route;
        this.pkgCount = pkgCount;
    }

    int getPkgCount() {
        return pkgCount;
    }

    String getRoute() {
        return route;
    }

    String getName() {
        return name;
    }

    String getTown() {
        return town;
    }
}
