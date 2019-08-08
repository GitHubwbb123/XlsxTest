package com.mrwxb.xlsxtest;
import android.Manifest;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.EditText;
import android.widget.Toast;

import com.yanzhenjie.permission.AndPermission;
import com.yanzhenjie.permission.PermissionYes;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.List;

    public class MainActivity extends AppCompatActivity {
        EditText output;
        String filePath=Environment.getExternalStorageDirectory().getPath()+"/file.xlsx";
        @Override
        protected void onCreate(Bundle savedInstanceState) {
            super.onCreate(savedInstanceState);
            setContentView(R.layout.activity_main);
            output = (EditText) findViewById(R.id.textOut);
            if (AndPermission.hasPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {

            } else {
                AndPermission.with(this)
                        .permission(Manifest.permission.WRITE_EXTERNAL_STORAGE)
                        .requestCode(100)
                        .send();
            }
        }

        public void onReadClick(View view) {
            //InputStream stream = getResources().openRawResource(R.raw.test1);
            ExcelPOIUtil.read(output,Environment.getExternalStorageDirectory().getPath()+"/file.xlsx");

        }
        public void onWriteClick(View view) {
            ExcelPOIUtil.write(filePath);
        }
        public void onUpdateClick(View view) {
            ExcelPOIUtil.update(filePath,0,3,1,"我是追加的内容");//一行中间不能有空白，否则空白后的单元就读不出来

        }




        @Override
        public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
            AndPermission.onRequestPermissionsResult(this, requestCode, permissions, grantResults);
        }

        @PermissionYes(100)
        private void getPermission(List<String> grantedPermissions) {
            Toast.makeText(MainActivity.this, "接受权限", Toast.LENGTH_SHORT).show();
        }

    }
