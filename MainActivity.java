package com.example.tryexcel;

import android.Manifest;
import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.EditText;
import android.widget.ImageButton;
import android.widget.LinearLayout;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.FileProvider;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileOutputStream;


public class MainActivity extends AppCompatActivity {
    private ImageButton btn;
    private LinearLayout linear;

    private Bitmap bitmap;
    EditText editTextExcel;
    private File filePath = new File(Environment.getExternalStorageDirectory() + "/Demo.xls");


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
                Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);

        editTextExcel = findViewById(R.id.editText);
    }
    private static Cell cell;

    public void buttonCreateExcel(View view){

        cell = null;
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        HSSFSheet hssfSheet = hssfWorkbook.createSheet("Custom Sheet");


        HSSFRow hssfRow = hssfSheet.createRow(0);
        HSSFRow hssfRow1 = hssfSheet.createRow(1);
        HSSFCell hssfCell = hssfRow.createCell(0);
        HSSFCell hssfCell1 = hssfRow1.createCell(0);

        hssfCell.setCellValue(editTextExcel.getText().toString());

        hssfCell1.setCellValue("Hello Rashmi");
        try {
            File sdCard = Environment.getExternalStorageDirectory();
            File newFolder = new File(Environment.getExternalStorageDirectory(),
                    "Examples");
            if (!filePath.exists()){
                filePath.createNewFile();
            }

            FileOutputStream fileOutputStream= new FileOutputStream(filePath);
            hssfWorkbook.write(fileOutputStream);
            Toast.makeText(this, "successfully File created", Toast.LENGTH_SHORT).show();

            if (fileOutputStream!=null){
                fileOutputStream.flush();
                fileOutputStream.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        //hssfWorkbook.close();
        //Toast.makeText(this, "successfully pdf created", Toast.LENGTH_SHORT).show();

        openPdf();

    }
    private void openPdf() {
        File file = new File(Environment.getExternalStorageDirectory() + "", "Demo.xls");
        Log.d("xlsFIle", "" + file);

        if (file.exists()) {
            Toast.makeText(this, "Existed", Toast.LENGTH_SHORT).show();
            Uri uriPdfPath = FileProvider.getUriForFile(this, getApplicationContext().getPackageName() + ".provider", file);
            Log.d("xlsPath", "" + uriPdfPath);
            Intent pdfOpenIntent = new Intent(Intent.ACTION_VIEW);
            pdfOpenIntent.setFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP);
            pdfOpenIntent.setClipData(ClipData.newRawUri("", uriPdfPath));
            pdfOpenIntent.setDataAndType(uriPdfPath, "application/xls");
            pdfOpenIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);

            try {
                startActivity(pdfOpenIntent);
            } catch (ActivityNotFoundException e) {

                Toast.makeText(this, "No Application for xls view", Toast.LENGTH_SHORT).show();
            }
        }

    }}
