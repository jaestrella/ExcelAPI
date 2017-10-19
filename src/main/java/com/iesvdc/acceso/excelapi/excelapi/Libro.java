/*
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.util.List;

/**
 * Esta clase alacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas.
 * @author Daniel Morillo Expósito
 */
public class Libro {
    private List<Hoja> hojas;
    private String nombreArchivo;
}
