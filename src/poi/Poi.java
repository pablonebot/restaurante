/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package poi;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.jayway.jsonpath.Configuration;
import com.jayway.jsonpath.JsonPath;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import com.jayway.jsonpath.InvalidJsonException;
import com.jayway.jsonpath.JsonPathException;
import java.io.FileOutputStream;
import net.minidev.json.JSONArray;
import net.minidev.json.JSONObject;
import net.minidev.json.JSONStyle;
import net.minidev.json.JSONValue;
import net.minidev.json.parser.JSONParser;
import net.minidev.json.parser.ParseException;
import net.minidev.json.writer.JsonReaderI;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Pablo nebot
 */
public class Poi {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        // TODO code application logic here

        JsonObject json = leerJson("restaurante.json");

     ArrayList<Comida> menu=crearMenu(json);
               generarExcell(menu);
               try{  
          Runtime.getRuntime().exec("cmd /c start "+"productos.xls");
          }catch(IOException  e){
              e.printStackTrace();
          }  
    }
    public static void generarExcell( ArrayList<Comida> comida) throws FileNotFoundException, IOException{
        Workbook wb = new HSSFWorkbook();
    
    Sheet sheet = wb.createSheet("menu");

    Row row = sheet.createRow(0);

    Cell cell0 = row.createCell(0);
     Cell cell1 = row.createCell(1);
      Cell cell2 = row.createCell(2);
       Cell cell3 = row.createCell(3);
    cell0.setCellValue("nombre");
      cell1.setCellValue("precio");
        cell2.setCellValue("categoria");
          cell3.setCellValue("tamaños");

        for (int i = 0; i < comida.size(); i++) {
           
              Row row1 = sheet.createRow(i+1);
              Cell cn=row1.createCell(0);
              Cell cp=row1.createCell(1);
              Cell  cc=row1.createCell(2);
              Cell ct=row1.createCell(3);
              cn.setCellValue(comida.get(i).getNombre());
              cp.setCellValue(comida.get(i).getPrecio());
               cc.setCellValue(comida.get(i).getCategoria());
               String tam=comida.get(i).getTam()[0];
               for (int j = 1; j < comida.get(i).getTam().length; j++) {
                tam+=","+comida.get(i).getTam()[j];
                
            }
               ct.setCellValue(tam);
        }
 

 sheet. autoSizeColumn(4);

    
    FileOutputStream fileOut = new FileOutputStream("productos.xls");
    wb.write(fileOut);
        System.out.println("excel generaado con exito");
    fileOut.close();
    
    }
    public static ArrayList<Comida> crearMenu(JsonObject json){
        
          JsonElement  elemento1 = json.get("menu");
       JsonObject json2=elemento1.getAsJsonObject();
                  JsonElement  elemento2 = json2.get("items");
                JsonArray arr=elemento2.getAsJsonArray();
                ArrayList<Comida> menu=new  ArrayList<Comida>();
                for (int i = 0; i < arr.size(); i++) {
                    
                Comida nuevacomida=new Comida();
             JsonObject  json3=(JsonObject) arr.get(i);
            JsonElement  elementocomida=json3.get("name");
           
                    String nombre=elementocomida.getAsString();
                     nuevacomida.setNombre(nombre);
                     JsonElement  elementoprecio=json3.get("qty");
                    String precio=elementoprecio.getAsString();
                     nuevacomida.setPrecio(precio);
                     JsonElement  elementocategoria=json3.get("category");
                    String categoria=elementocategoria.getAsString();
                    nuevacomida.setCategoria(categoria);
                     JsonElement  elementotamanio=json3.get("sizes");
                    JsonArray arr2=elementotamanio.getAsJsonArray();
                    String[] tamanio = new String[arr2.size()];
                   for (int j = 0; j < arr2.size(); j++){
                       
                   tamanio[j]=arr2.get(j).getAsString();
                }
                   nuevacomida.setTam(tamanio);
                   menu.add(nuevacomida);
                }
                return menu;
    }
    public static JsonObject leerJson(String file) {

        JsonObject jsonObject = new JsonObject();

        try {
            JsonParser parser = new JsonParser();
            JsonElement jsonElement = parser.parse(new FileReader(file));
            jsonObject = jsonElement.getAsJsonObject();
        } catch (FileNotFoundException e) {

        } catch (IOException ioe) {

        }
        return jsonObject;
    }
}
