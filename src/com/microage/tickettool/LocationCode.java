/*
 * The location code is used by Clearview to select the proper location.
 */
package com.microage.tickettool;

/**
 *
 * @author Brian Milne
 */
public class LocationCode {
    
    public static int LocationCode() {
        int floor = 1;
        if (Main.main.location.getSelectedItem().equals("OMP") || 
            Main.main.location.getSelectedItem().equals("333") || 
            Main.main.location.getSelectedItem().equals("350")) 
        {
            floor = Integer.parseInt(Main.main.floor.getSelectedItem().toString());
        }
        if (Main.main.location.getSelectedItem().equals("OMP")) {
            switch (floor) {
                case 1: return 1;
                case 2: return 2;
                case 3: return 3;
                case 4: return 4;
                case 5: return 5;
                case 6: return 6;
                case 7: return 7;
                case 8: return 8;
                case 9: return 9;
                case 10: return 10;
                case 11: return 11;
                case 12: return 12;
                case 13: return 13;
                case 14: return 14;
                case 15: return 15;
                case 16: return 16;
                default: return 1;
            }
        }
        else if (Main.main.location.getSelectedItem().equals("333 Bloor")) {
            switch (floor) {
                case 1: return 17;
                case 2: return 18;
                case 3: return 19;
                case 4: return 20;
                case 5: return 21;
                case 6: return 22;
                case 7: return 23;
                case 8: return 24;
                case 9: return 25;
                case 10: return 26;
                default: return 17;
            }
        }
        else if (Main.main.location.getSelectedItem().equals("350 Bloor")) { 
            switch (floor) {
                case 1: return 27;
                case 2: return 28;
                case 3: return 29;
                case 4: return 30;
                case 5: return 31;
                case 6: return 32;
                case 7: return 33;
                default: return 27;
            }
        }
        else if (Main.main.location.getSelectedItem().equals("Brampton")) {
            if (Main.main.floor.getSelectedItem().toString().contains("City Center") || Main.main.floor.getSelectedItem().toString().isEmpty()) 
                return 36;
            else if (Main.main.floor.getSelectedItem().toString().contains("South"))
                return 37;
            else if (Main.main.floor.getSelectedItem().toString().contains("East"))
                return 38;
            else if (Main.main.floor.getSelectedItem().toString().contains("North Building"))
                return 39;
            else if (Main.main.floor.getSelectedItem().toString().contains("Networking"))
                return 40;             
            else
                return 36;
        }
        else if (Main.main.location.getSelectedItem().equals("Other")) {
            if (Main.main.floor.getSelectedItem().equals("1 Blue Jays Way"))
                return 35;
            else if (Main.main.floor.getSelectedItem().equals("855 York Mills"))
                return 41;
            else
                return 1;
        }
        return 1;
    }
}

