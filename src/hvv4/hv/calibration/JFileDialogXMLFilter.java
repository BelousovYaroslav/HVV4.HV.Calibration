/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package hvv4.hv.calibration;

import java.io.File;
import javax.swing.filechooser.FileFilter;

/**
 *
 * @author yaroslav
 */
public class JFileDialogXMLFilter  extends FileFilter {

    @Override
    public boolean accept( File f) {
        if (f.isDirectory()) {
            return true;
        }

        String extension = Utils.getExtension( f);
        if( extension != null) {
            if( extension.equals( Utils.xml))
                return true;
            else
                return false;
        }

        return false;
    }

    @Override
    public String getDescription() {
        return "XML files";
    }
}
