/*
 */
package com.iesvdc.acceso.excelapi.excelapi;

/**
 * Esta clase almacena información del texto de
 * una hoja de cálculo.
 * 
 * @author Daniel Morillo Expósito
 */
public class Hoja {
    private String[][] datos;
    private String nombre;
public Hoja(){
    this.datos = new String[0][0];
    this.nombre = "";
}
public Hoja(int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre="";
    }     
    public Hoja(String nombre,int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre=nombre;
    }

    public String getDato(int fila, int columna) {
        return datos[fila][columna];
    }

    public void setDato(String dato,int fila,int columna) {
        this.datos[fila][columna] = dato;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    
    
    
}
