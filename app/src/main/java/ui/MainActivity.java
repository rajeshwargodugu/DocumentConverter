package ui;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import android.content.Intent;
import android.content.pm.PackageManager;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import com.example.documentconverter.R;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    public static final String TAG = MainActivity.class.getSimpleName();

    private Button browseButton;
    private TextView sourceFileName;

    //Uri to store the image uri
    private Uri sourceUri;

    private int PICK_XLS_REQUEST = 1;

    //storage permission code
    private static final int STORAGE_PERMISSION_CODE = 100;
    private static final int STORAGE_WRITE_PERMISSION_CODE = 200;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        browseButton = findViewById(R.id.select_button);
        sourceFileName = findViewById(R.id.source_file_name);
        browseButton.setOnClickListener(this);
    }

    private void browseDirectory() {
        Log.d(TAG, "BrowseDirectory");
        Intent intent = new Intent();
        intent.setType("application/*");
        intent.setAction(Intent.ACTION_GET_CONTENT);
        String state = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(state)) {
            if (Build.VERSION.SDK_INT >= 23) {
                if (checkReadPermission()) {
                    startActivityForResult(Intent.createChooser(intent, "Select Document"), PICK_XLS_REQUEST);
                } else {
                    requestReadPermission();
                }
            }
        }


    }

    @Override
    public void onClick(View v) {
        browseDirectory();
    }

    private void convertToPdf(String path) {
        try {
            File storageDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
            FileInputStream input_document = new FileInputStream(new File(path));
            // Read workbook into HSSFWorkbook
            XSSFWorkbook my_xls_workbook = new XSSFWorkbook(input_document);
            // HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document);
            // Read worksheet into HSSFSheet
            XSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
            // HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
            // To iterate over the rows
            Iterator<Row> rowIterator = my_worksheet.iterator();
            //We will create output PDF document objects at this point
            Document iText_xls_2_pdf = new Document();
            if (checkWritePermission()) {
                String outFilePath = storageDir + "/Excel_To_PDF_Output.pdf";
                PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream(outFilePath));
                iText_xls_2_pdf.open();
                //we have two columns in the Excel sheet, so we create a PDF table with two columns
                PdfPTable my_table = new PdfPTable(32);
                //We will use the object below to dynamically add new data to the table
                PdfPCell table_cell;
                //Loop through rows.

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next(); //Fetch CELL
                        switch (cell.getCellType()) { //Identify CELL type
                            //you need to add more code here based on
                            //your requirement / transformations
                            case Cell.CELL_TYPE_STRING:
                                //Push the data from Excel to PDF Cell
                                table_cell = new PdfPCell(new Phrase(cell.getStringCellValue()));

                                //feel free to move the code below to suit to your needs
                                my_table.addCell(table_cell);
                                break;
                        }
                        //next line
                    }

                }

                //Finally add the table to PDF document
                iText_xls_2_pdf.add(my_table);
                iText_xls_2_pdf.close();
                //we created our pdf file..
                input_document.close(); //close xls
            } else {
                requestWritePermission();
            }
        } catch (Exception e) {
            Log.d(TAG, e.getMessage());
        }


    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (requestCode == PICK_XLS_REQUEST && resultCode == RESULT_OK && data != null && data.getData() != null) {
            sourceUri = data.getData();
            Log.d(TAG, sourceUri.getPath());
            String filename = getPath(sourceUri);
            sourceFileName.setText(filename.substring(filename.lastIndexOf("/") + 1));
            // convertToPdf(getPath(sourceUri));
        }
    }

    //method to get the file path from uri
    public String getPath(Uri uri) {
        Cursor cursor = getContentResolver().query(uri, null, null, null, null);
        cursor.moveToFirst();
        String document_id = cursor.getString(0);
        document_id = document_id.substring(document_id.lastIndexOf(":") + 1);
        cursor.close();
        /*cursor = getContentResolver().query(
                android.provider.MediaStore.Images.Media.EXTERNAL_CONTENT_URI,
                null, MediaStore.Images.Media._ID + " = ? ", new String[]{document_id}, null);
        cursor.moveToFirst();
        String path = cursor.getString(cursor.getColumnIndex(MediaStore.Images.Media.DATA));
        cursor.close();
*/
        return document_id;
    }

    private boolean checkReadPermission() {
        int result = ContextCompat.checkSelfPermission(MainActivity.this, android.Manifest.permission.READ_EXTERNAL_STORAGE);
        if (result == PackageManager.PERMISSION_GRANTED)
            return true;
        else
            return false;
    }

    private boolean checkWritePermission() {
        int result = ContextCompat.checkSelfPermission(MainActivity.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE);

        if (result == PackageManager.PERMISSION_GRANTED)
            return true;
        else
            return false;
    }

    private void requestReadPermission() {

        if (ActivityCompat.shouldShowRequestPermissionRationale(MainActivity.this, android.Manifest.permission.READ_EXTERNAL_STORAGE)) {
            Toast.makeText(MainActivity.this, "Read External Storage permission allows us to read files. Please allow this permission in App Settings.", Toast.LENGTH_LONG).show();
        } else {
            ActivityCompat.requestPermissions(MainActivity.this, new String[]{android.Manifest.permission.READ_EXTERNAL_STORAGE}, STORAGE_PERMISSION_CODE);
        }
    }

    private void requestWritePermission() {

        if (ActivityCompat.shouldShowRequestPermissionRationale(MainActivity.this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
            Toast.makeText(MainActivity.this, "Write External Storage permission allows us to write files. Please allow this permission in App Settings.", Toast.LENGTH_LONG).show();
        } else {
            ActivityCompat.requestPermissions(MainActivity.this, new String[]{android.Manifest.permission.WRITE_EXTERNAL_STORAGE,}, STORAGE_WRITE_PERMISSION_CODE);
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String permissions[], int[] grantResults) {
        switch (requestCode) {
            case STORAGE_PERMISSION_CODE:
                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                    Log.e(TAG, "Permission Granted, Now you can use local drive .");
                } else {
                    Log.e(TAG, "Permission Denied, You cannot use local drive .");
                }
                break;
            case STORAGE_WRITE_PERMISSION_CODE:
                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                    Log.e(TAG, "Permission Granted, Now you can write local drive .");
                } else {
                    Log.e(TAG, "Permission Denied, You cannot write local drive .");
                }
                break;
        }
    }

}
