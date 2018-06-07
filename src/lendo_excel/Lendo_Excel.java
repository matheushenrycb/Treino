/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package lendo_excel;

import java.io.File;

import jxl.Cell;

import jxl.Sheet;


/**
 *
 * @author Wander Luiz
 */
public class Lendo_Excel {
    
    private Workbook planilha; //objeto que receberá a instancia da planilha estudada
    private Sheet aba; //objeto que será a aba
    private File arquivo; //arquivo .xls que será lido
    
    public Lendo_Excel(){
        try{
            arquivo = new File("F:\\teste.xls"); //instancia a planilha
            
            planilha = Workbook.getWorkbook(arquivo); //obtendo as abas da planilha
            
            Sheet[] abas = planilha.getSheets(); //loop para o vetor de abas
            
            aba = planilha.getSheet(0); //pega a primeira aba ou seja de indice 0.
            
            String[][] matriz = new String[aba.getRows()][aba.getColumns()];
            //matriz.length -> representa as linhas da matriz
            //matriz[0].length -> pega o tamanho da linha [0] ou seja pega numeros de colunas
            
            //Cell[] linha = aba.getRow(0); // pega a primeira linha, ou seja, linhas de indice 0
            
            Cell[] cel; //Instancia um array de celulas que ira auxiliar no povoamento de matriz
            
            //for(Cell c : linha){
              //  System.out.println(c.getContents());
            //}
            for (int i = 0; i < matriz.length; i++){
            
            cel = aba.getRow(i);
            
            for(int j = 0; j < matriz[0].length; j++){ //pega os dados da cel[j] e adiciona na matriz
                
                matriz[i][j] = cel[j].getContents();
                
            }
            
        }
            
            //imprime os dados da matriz
            
            for(int i = 0; i < matriz.length; i++){
                
                for(int j = 0; j < matriz[0].length; j++){
                    
                    System.out.println(matriz[i][j] + "\t");
                }
                
                System.out.println("");
            }
        
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        new Lendo_Excel();
    }
    
}
