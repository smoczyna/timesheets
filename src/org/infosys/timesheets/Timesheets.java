/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.infosys.timesheets;

/**
 *
 * @author 58128
 */
public class Timesheets {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        TimeTemplate template = new TimeTemplate();
        template.prepareTimesheet(args);
    }
    
}
