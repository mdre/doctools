/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.github.mdre.templater;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Stream;

/**
 *
 * @author Marcelo D. Ré {@literal <marcelo.re@gmail.com>}
 */
public class FillerCommand {
    private final static Logger LOGGER = Logger.getLogger(FillerCommand.class .getName());
    static {
        if (LOGGER.getLevel() == null) {
            LOGGER.setLevel(Level.INFO);
        }
    }
    
    public enum COMMAND {
        REPLACE
    }
    
    List<Command> comands;

    public FillerCommand() {
        this.comands = new ArrayList<>();
    }
    
    
    public FillerCommand add(String pattern, String replaString) {
        this.comands.add(new Command(COMMAND.REPLACE, pattern, replaString));
        return this;
    }
    
    public Iterator<Command> iterator() {
        return this.comands.iterator();
    }
    
    public Stream<Command> stream() {
        return this.comands.stream();
    }
    
   
    
}

class Command {

    public Command(FillerCommand.COMMAND command, String pattern, String replace) {
        this.command = command;
        this.pattern = pattern;
        this.replace = replace;
    }
    
    public FillerCommand.COMMAND command;
    public String pattern;
    public String replace;
}