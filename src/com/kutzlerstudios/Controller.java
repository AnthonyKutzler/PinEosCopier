package com.kutzlerstudios;



import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;

import java.io.File;

public class Controller {

    @FXML
    Label runningLabel, finishLabel, exceptionLabel;
    @FXML
    Node root;
    @FXML
    TextField quads;
    @FXML
    TextField manifest;


    private File maniFile;

    public void browseManifest(){
        maniFile = fileBrowser("Select Route Manifest");
        manifest.setText(maniFile.getAbsolutePath());
    }


    public void finishSetup() {
        try {
            runLabel();
            new EosExcel(new AssignmentsExcel().setupDriverRouteData(new BulkAddExcel().setupFinalBulkAdd()));
            finishedLabel();
        }catch (Exception e){
            errorLabel(e.getLocalizedMessage());
        }
    }

    public void initialSetup() {
        try {
            runLabel();
            new BulkAddExcel(maniFile).setupInitBulkAdd(quads.getText().split(","));
            finishedLabel();
        } catch (Exception e) {
            errorLabel(e.getLocalizedMessage());
        }
    }


    private File fileBrowser(String title){
        FileChooser chooser = new FileChooser();
        chooser.setTitle(title);
        return chooser.showOpenDialog(root.getScene().getWindow());
    }


    private void runLabel(){
        exceptionLabel.setVisible(false);
        runningLabel.setVisible(true);
        finishLabel.setVisible(false);
    }

    private void finishedLabel(){
        exceptionLabel.setVisible(false);
        runningLabel.setVisible(false);
        finishLabel.setVisible(true);
    }

    private void errorLabel(String msg){
        exceptionLabel.setVisible(true);
        runningLabel.setVisible(false);
        finishLabel.setVisible(false);
        exceptionLabel.setText(msg);
    }

}
