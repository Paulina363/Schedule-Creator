/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel.fx;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.concurrent.Task;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.effect.Light.Distant;
import javafx.scene.effect.Lighting;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import javax.swing.JOptionPane;


/**
 *
 * @author Paulina Chametka
 */
public class EXCELFX extends Application {
    //public Label lbl = new Label();
    //final ProgressBar pb = new ProgressBar(-1f);
    @Override
    public void start(Stage primaryStage) {
        
        final Button btn0 = new Button();
        btn0.setText("Responses" );
        btn0.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {            
            try {
                Desktop.getDesktop().open(new File(Generator.FileNameResponses));
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null,"CREATE SCHEDULE BEFORE USING THIS FEATURE","Error",JOptionPane.ERROR_MESSAGE);
            }
        }
        });
        
        final Button btn1 = new Button();
        btn1.setText("Schedule");
        btn1.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
         try {
               Desktop.getDesktop().open(new File(Generator.FileNameSchedule));
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null,"CREATE SCHEDULE BEFORE USING THIS FEATURE","Error",JOptionPane.ERROR_MESSAGE);
            }
            
        }
        });
        
        final Button btn2 = new Button();
        btn2.setText("Labels");
        btn2.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
         try {
              Desktop.getDesktop().open(new File(Generator.FileNameLabels));
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null,"CREATE SCHEDULE BEFORE USING THIS FEATURE","Error",JOptionPane.ERROR_MESSAGE);
            }
            
        }
        });
        
        final Button btn3 = new Button();
        btn3.setText("Subject");
        btn3.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
         try {
                Desktop.getDesktop().open(new File(Generator.FileNameSubject));
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null,"CREATE SCHEDULE BEFORE USING THIS FEATURE","Error",JOptionPane.ERROR_MESSAGE);
            }   
            
        }
        });
        
        final Button btn4 = new Button();
        btn4.setText("Schedule Type");
        btn4.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
         try {
                Desktop.getDesktop().open(new File(Generator.FileNameScheduleType));
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(null,"CREATE SCHEDULE BEFORE USING THIS FEATURE","Error",JOptionPane.ERROR_MESSAGE);
            }  
            
        }
        });
        
        final HBox hbTop = new HBox();
        hbTop.setSpacing(5);
        hbTop.setAlignment(Pos.CENTER);
        hbTop.getChildren().addAll(btn0, btn1, btn2,btn3,btn4);
        
        File file = new File("images/ayj.jpg");
        //Image image = new Image(file.toURI().toString());
        final ImageView imv = new ImageView();
        final Image image2 = new Image(file.toURI().toString());
        imv.setImage(image2);
        
        final Label lblMode = new Label("Select schedule mode:");        
        final ToggleGroup group = new ToggleGroup();

        //A radio button with an empty string for its label
        final RadioButton rb1 = new RadioButton();
        rb1.setToggleGroup(group);
        rb1.setSelected(true);
        //Setting a text label
        rb1.setText("By Date");
        //A radio button with the specified label
        final RadioButton rb2 = new RadioButton("Randomly");
        rb2.setToggleGroup(group);
        
        final HBox hb = new HBox();
        hb.setSpacing(5);
        hb.setAlignment(Pos.CENTER);
        hb.getChildren().addAll(lblMode, rb1, rb2);

         final ProgressBar pb = new ProgressBar(-1f);
         pb.setVisible(false);
         pb.setMinWidth(200);
        
         final Label lbl = new Label();
         
        final Button btnCreate = new Button();
        btnCreate.setText("Create Schedule");
        btnCreate.setOnAction(new EventHandler<ActionEvent>() {
        @Override
        public void handle(ActionEvent event) {
    
            final Generator  generator = new Generator();
            System.out.println("Start application!");
                       
            Task task;
            final BooleanProperty isTaskRunning = new SimpleBooleanProperty(false);
            final BooleanProperty isTaskNotRunning = new SimpleBooleanProperty(true);
            task = new Task<Void>() {
                @Override
                public Void call() throws InterruptedException, IOException { 
                    
                    //Calculate  the Random mode
                    Boolean bRandom = false;
                    if(rb2.isSelected())
                         bRandom = true;
                    
                    isTaskRunning.set(true);
                    isTaskNotRunning.set(false);
                    
                     updateMessage("Reading Subjects...");
                     generator.readSubjects();
                     updateMessage("Reading Responses...");
                     generator.readResponses(bRandom); //converts spreadsheet to array list
                     updateMessage("Reading Schedule Type...");
                     generator.readScheduleType();
                     updateMessage("Generating Schedule...");
                     
                     generator.generateSchedule();
                     updateMessage("Generating Labels...");
                     generator.generateLabels(); //coverts schedule array list to label array list
                     updateMessage("Writing Schedule...");
                     generator.writeSchedule();//creates spreadsheet from schedule array list 
                     updateMessage("Writing Label...");
                     generator.writeLabel();//creates spreadsheet from label array list 
                     updateMessage("Completed...Wait..."); 
                     isTaskRunning.set(false); 
                     Thread.sleep(2000);
                     updateMessage("");
                     isTaskNotRunning.set(true);
                     
                    return null;
                }
            };
            lbl.textProperty().bind(task.messageProperty());
            btnCreate.visibleProperty().bind(isTaskNotRunning);
            pb.visibleProperty().bind(isTaskRunning);
            new Thread(task).start();  
        }
    });
    
    final VBox hb1 = new VBox();
        hb1.setSpacing(5);
        hb1.setMinWidth(200);
        hb1.setAlignment(Pos.CENTER);
        hb1.getChildren().addAll(hbTop,imv,lighting(),hb,pb,lbl,btnCreate);
     
     Scene scene = new Scene(hb1, 500, 350);       
    primaryStage.setTitle("Schedule Creator");
    primaryStage.setScene(scene);
    primaryStage.show();
        
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }
    
    static Node lighting() {
        Distant light = new Distant();
        light.setAzimuth(-135.0f);
 
        Lighting l = new Lighting();
        l.setLight(light);
        l.setSurfaceScale(5.0f);
 
        Text t = new Text();
        t.setText("Schedule Generator");
        t.setFill(Color.BLUE);
        t.setFont(Font.font("null", FontWeight.BOLD, 40));
        t.setX(10.0f);
        t.setY(10.0f);
        t.setTextOrigin(VPos.TOP);
 
        t.setEffect(l);
 
        t.setTranslateX(0);
        t.setTranslateY(0);
 
        return t;
    }
    
}
